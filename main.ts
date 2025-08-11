import { 
    App, 
    Plugin, 
    PluginSettingTab, 
    Setting, 
    WorkspaceLeaf, 
    TFile, 
    Notice, 
    ItemView,
    FileView,
    normalizePath,
    Platform,
    FileSystemAdapter,
    requestUrl,
    ViewStateResult,
    Modal
} from 'obsidian';
import * as path from 'path';
import * as fs from 'fs';
import { SignJWT } from 'jose';
import * as http from 'http';
import * as mime from 'mime-types'; // Add: npm install mime-types

// Define the interface for plugin settings - RADICALLY SIMPLIFIED
interface OnlyOfficePluginSettings {
    onlyOfficeServerPort: number;
    jwtSecret: string;
    localServerAddress: string;
    htmlServerPort?: number;
    callbackServerPort?: number;
    useRequestToken?: boolean;
    useSystemSaveAs?: boolean; // open download in external browser instead of in-app prompt
    hideDownloadAs?: boolean; // hide Download/Download copy menu entry
    promptNameOnCreate?: boolean; // prompt user for name before creating new file
}

// Default settings - RADICALLY SIMPLIFIED
const DEFAULT_SETTINGS: OnlyOfficePluginSettings = {
    onlyOfficeServerPort: 8080,
    jwtSecret: 'your-secret-key-please-change',
    localServerAddress: 'host.docker.internal',
    htmlServerPort: 0, // 0 = dynamic starting at 8081
    callbackServerPort: 0, // 0 = dynamic starting at 8082
    useRequestToken: true,
    useSystemSaveAs: false,
    hideDownloadAs: true,
    promptNameOnCreate: true,
};

// Define the OnlyOffice document view (reuse single view for multiple formats)
const VIEW_TYPE_ONLYOFFICE = 'onlyoffice-docx'; // keep legacy id for existing workspaces

// Add this interface to define the WebviewTag for TypeScript
interface WebviewTag extends HTMLElement {
    src: string;
    webpreferences: string;
    executeJavaScript(code: string, userGesture?: boolean): Promise<any>;
    addEventListener(type: 'did-finish-load', listener: () => void): void;
    addEventListener(type: 'did-fail-load', listener: (error: {errorCode: number, errorDescription: string, validatedURL: string}) => void): void;
    addEventListener(type: 'console-message', listener: (event: { level: number, message: string, line: number, sourceId: string }) => void): void;
    addEventListener(type: 'ipc-message', listener: (event: { channel: string, args: any[] }) => void): void;
}

// Add type declaration for DocsAPI on window object
declare global {
    interface Window {
        DocsAPI?: {
            DocEditor: new (id: string, config: any) => any;
        };
    }
};

// Forward declaration of plugin interface
interface IOnlyOfficePlugin extends Plugin {
    settings: OnlyOfficePluginSettings;
    httpServer: http.Server | null;
    callbackServer: http.Server | null;
    localServerPort: number;
    callbackServerPort: number;
    openDocxFile(file: TFile): Promise<void>;
    openOnlyOfficeFile(file: TFile): Promise<void>; // generic multi-format opener
    openNewDocument(): Promise<void>;
    saveAsPromise?: {
        resolve: (data: ArrayBuffer) => void,
        reject: (reason?: any) => void,
        keyPrefix: string
    } | null;
    keyFileMap?: Record<string,string>;
}

/**
 * OnlyOffice Document View class
 * This class handles rendering of DOCX/XLSX/PPTX/PDF files in the OnlyOffice editor
 */
class OnlyOfficeDocumentView extends FileView {
    // The `file` property is now inherited from ItemView and managed by Obsidian's core.
    // We no longer need to declare it here.
    public webview: WebviewTag | null = null;
    private iframe: HTMLIFrameElement | null = null;
    private plugin: IOnlyOfficePlugin;
    public viewId: string; // Make viewId public and stable
    private boundMessageHandler: ((event: MessageEvent) => void) | null = null;
    private fallbackEditor: HTMLElement | null = null;
    public isDirty: boolean = false; // Add a dirty flag
    private pendingSaveAsName: string | null = null; // store name chosen before download phase
    private saveButton: HTMLButtonElement | null = null;
    private saveAsButton: HTMLButtonElement | null = null;
    private dirtyCheckInterval: number | null = null; // For polling dirty state
    private systemSaveAsRetryPending: boolean = false; // guard for new-doc system Save As
    private performingSystemSaveAsFallback: boolean = false; // prevent concurrent fallback injections
    private currentExt: string = 'docx'; // track active document extension

    // Add these properties for event listeners
    private _didFinishLoadListener: (() => void) | null = null;
    private _didFailLoadListener: ((error: any) => void) | null = null;

    constructor(leaf: WorkspaceLeaf, plugin: IOnlyOfficePlugin) {
        super(leaf);
        this.plugin = plugin;
        this.viewId = Math.random().toString(36).substring(2, 15); // Assign a unique ID to each view instance
        // `this.file` is now managed by Obsidian. No need to initialize here.
        // The handleSaveMessage method and its binding are removed.
    }

    // Return the type for this view
    public getViewType() {
        return VIEW_TYPE_ONLYOFFICE;
    }

    // Return the display text for this view (the file name)
    public getDisplayText() {
        // This is now safe because Obsidian's core manages `this.file`
        return this.file?.name || 'Document';
    }

    // Get the icon for this view
    getIcon(): string {
        return 'file-text';   // Changed to a document icon for clarity
    }

    // The custom setState method is removed. The default implementation is sufficient when using `openFile`.

    // This method is called by Obsidian when a file is loaded into the view.
    async onLoadFile(file: TFile): Promise<void> {
        try {
            // --- Detect template state for new document ---
            const state = this.leaf?.getViewState();
            // Template flow removed: we now always create a real file before opening
            const isTemplate = false;

            // --- Early validation of file and container ---
            const container = this.contentEl as HTMLElement; // CHANGED: use contentEl instead of containerEl.children[1]
            if (!container) {
                console.error("OnlyOffice: Could not find view content container. Aborting onLoadFile.");
                return;
            }
            
            // Check if we have a valid file or template state
            if (!file && !isTemplate) {
                container.empty();
                const errorDiv = container.createEl('div', { cls: 'onlyoffice-error' });
                errorDiv.createEl('h2', { text: 'File Error' });
                errorDiv.createEl('p', { text: "Could not open document: file not found or view was opened without a file." });
                return;
            }

            // --- CRITICAL: Update this.file to the current file ---
            if (file) {
                (this as any).file = file;
            }

            // --- Prevent recursive onLoadFile calls by handling view state updates first ---
            if (this.leaf && file) {
                const state = this.leaf.getViewState();
                if (!state.state || state.state.file !== file.path) {
                    await this.leaf.setViewState({
                        ...state,
                        state: { ...(state.state || {}), file: file.path }
                    });
                    // Exit early - the setViewState will trigger another onLoadFile call
                    return;
                }
            }

            // --- Clean up existing webview to avoid memory leaks ---
            if (this.webview) {
                    if (this._didFinishLoadListener) {
                        this.webview.removeEventListener('did-finish-load', this._didFinishLoadListener as any);
                        this._didFinishLoadListener = null;
                }
                if (this._didFailLoadListener) {
                    this.webview.removeEventListener('did-fail-load', this._didFailLoadListener as any);
                    this._didFailLoadListener = null;
                }
                this.webview.remove();
                this.webview = null;
            }
            
            // Reset dirty flag when loading a new file
            this.isDirty = false;
            
            // Clear any polling intervals to avoid duplicate intervals
            if (this.dirtyCheckInterval) {
                clearInterval(this.dirtyCheckInterval);
                this.dirtyCheckInterval = null;
            }

            const localHostBaseUrl = `http://127.0.0.1:${this.plugin.localServerPort}`;
            const dockerHostBaseUrl = `http://${this.plugin.settings.localServerAddress || 'host.docker.internal'}:${this.plugin.localServerPort}`;
            const _editorUrl = `${localHostBaseUrl}/editor.html`;

            // --- More robust server availability check ---
            let htmlOk = false;
            let waited = 0;
            const maxWait = 5000; // Increase timeout for slower systems
            const checkInterval = 250; // Check less frequently
            
            while (waited < maxWait && !htmlOk) {
                try {
                    await requestUrl({ 
                        url: _editorUrl, 
                        method: 'HEAD',
                        headers: { 'Cache-Control': 'no-cache' }
                    });
                    htmlOk = true;
                } catch (e) {
                    await new Promise(res => setTimeout(res, checkInterval));
                    waited += checkInterval;
                }
            }
            
            if (!htmlOk) {
                const error = "OnlyOffice: editor.html is not available from the internal server.";
                console.error(error);
                new Notice(error);
                this.showEditorError(container, "Could not connect to the local server. Please check that the server is running.");
                return;
            }

            let effectiveFile = file;
            if (file) {
                // Wait for the DOCX file if not a template
                const relPathWait = file.path.replace(/\\/g, '/').split('/').map(encodeURIComponent).join('/');
                const docxUrl = `${localHostBaseUrl}/${relPathWait}`;
                let docxOk = false;
                waited = 0;
                
                while (waited < maxWait && !docxOk) {
                    try {
                        await requestUrl({ 
                            url: docxUrl, 
                            method: 'HEAD',
                            headers: { 'Cache-Control': 'no-cache' }
                        });
                        docxOk = true;
                    } catch (e) {
                        await new Promise(res => setTimeout(res, checkInterval));
                        waited += checkInterval;
                    }
                }
                
                if (!docxOk) {
                    const error = `OnlyOffice: The selected DOCX file '${file.name}' is not available from the internal server.`;
                    console.error(error);
                    new Notice(error);
                    this.showEditorError(container, `File '${file.name}' could not be accessed by the server. Please try again or check the file permissions.`);
                    return;
                }
            }

            // --- Always clear container before adding new content ---
            container.empty();

            // --- Generate a safe unique key (avoid slashes/spaces) and store mapping for callback resolution ---
            const randomId = Math.random().toString(36).slice(2,10);
            const mtimeOrNow = effectiveFile.stat?.mtime || Date.now();
            const baseForKey = effectiveFile.path + '|' + mtimeOrNow + '|' + randomId;
            let hash = 0; for (let i=0;i<baseForKey.length;i++){ hash = (hash*31 + baseForKey.charCodeAt(i)) >>> 0; }
            const uniqueKey = 'oo_' + hash.toString(36) + '_' + randomId;
            if (!this.plugin.keyFileMap) this.plugin.keyFileMap = {};
            this.plugin.keyFileMap[uniqueKey] = effectiveFile.path;

            // --- Build the correct file URLs for OnlyOffice ---
            const relPath2Eff = effectiveFile.path.replace(/\\/g, '/').split('/').map(encodeURIComponent).join('/');
            const hostFileUrlEff = `${localHostBaseUrl}/${relPath2Eff}`;
            const dockerFileUrlEff = `${dockerHostBaseUrl}/${relPath2Eff}`;
            const callbackUrlEff = `http://${this.plugin.settings.localServerAddress || 'host.docker.internal'}:${this.plugin.callbackServerPort}/callback`;

            // --- Determine if this is the Start.docx template (virtual or not) ---
            const isStartDocx = false; // direct file open only

            // Multi-format: ensure extension supported
            const ext = (effectiveFile?.extension || '').toLowerCase();
            this.currentExt = ext || 'docx';
            const supported = ['docx','xlsx','pptx','pdf'];
            if (!supported.includes(ext)) {
                this.showEditorError(container, `Unsupported file type .${ext}. Supported: ${supported.join(', ')}`);
                return;
            }

            await this.loadDirectOnlyOfficeInterface(
                container,
                hostFileUrlEff,
                dockerFileUrlEff,
                dockerHostBaseUrl,
                callbackUrlEff,
                effectiveFile,
                isStartDocx,
                uniqueKey
            );
        } catch (error) {
            console.error("OnlyOffice: Error in onLoadFile:", error);
            new Notice(`OnlyOffice: Failed to load document - ${error.message || 'Unknown error'}`);
            
            // Attempt to show error in the container
            try {
                const container = this.contentEl as HTMLElement;
                if (container) {
                    this.showEditorError(container, `An error occurred while loading the document: ${error.message || 'Unknown error'}`);
                }
            } catch (e) {
                // Last resort error handling
                console.error("OnlyOffice: Failed to show error in container:", e);
            }
        }
    }

    // This method is called by Obsidian when the file is about to be closed.
    async onUnloadFile(file: TFile): Promise<void> {
        // Trigger a save on close to prevent data loss, but only if the file is dirty.
        if (this.isDirty) {
            console.log("OnlyOffice: onUnloadFile triggered for dirty file, attempting to save.");
            try {
                // Use the reliable callback mechanism for saving on close
                this.webview?.executeJavaScript(`window.docEditor?.serviceCommand("c:forcesave", "");`);
                new Notice("Saving document on close...");
            } catch (error) {
                console.error("OnlyOffice: Failed to trigger save on close.", error);
            }
        } else {
            console.log("OnlyOffice: onUnloadFile triggered for clean file, no save needed.");
        }

        // Clean up resources here.
        if (this.boundMessageHandler) {
            window.removeEventListener('message', this.boundMessageHandler);
            this.boundMessageHandler = null;
        }
        if (this.dirtyCheckInterval) {
            clearInterval(this.dirtyCheckInterval);
            this.dirtyCheckInterval = null;
        }
        if (this.webview) {
            this.webview.remove();
            this.webview = null;
        }
        this.contentEl.empty();
    }

    // Load the OnlyOffice editor directly in the interface
    private async loadDirectOnlyOfficeInterface(
        container: HTMLElement,
        hostFileUrl: string,
        dockerFileUrl: string,
        dockerBaseUrl: string,
        callbackUrl: string,
        file: TFile,
        isStartDocx: boolean,
        uniqueKey?: string
    ): Promise<void> {
        console.log('Loading OnlyOffice editor interface');
    console.log('OnlyOffice debug: hostFileUrl', hostFileUrl, 'dockerFileUrl', dockerFileUrl, 'callbackUrl', callbackUrl);

        // --- REMOVE unused editorHtmlUrl variable ---
        // const editorHtmlUrl = `http://127.0.0.1:${this.plugin.localServerPort}/editor.html?doc=${encodeURIComponent(file.path)}&v=${uniqueKey}`;

        // Clear container and set proper styling for flex layout
        container.empty();
        container.style.cssText = 'height: 100%; width: 100%; display: flex; flex-direction: column;';
        

    // (Removed legacy host toolbar with extra New / Save As buttons; OnlyOffice's built-in UI now handles these actions.)

        // Create container for webview
        const webviewContainer = container.createEl('div', { 
            attr: { style: 'flex-grow: 1; position: relative; width: 100%; height: 100%;' } 
        });

        // Add a loading indicator
        const loadingDiv = container.createEl('div', {
            attr: { style: 'position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); z-index: 10; background: rgba(255,255,255,0.9); padding: 24px 40px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); font-size: 1.2em; color: #333;' }
        });
        loadingDiv.textContent = 'Loading OnlyOffice editor...';

        // --- REMOVE old listener properties ---
        this._didFinishLoadListener = null;
        this._didFailLoadListener = null;

        // --- Remove the postMessage listener, as we are switching to polling ---
        if (this.boundMessageHandler) {
            window.removeEventListener('message', this.boundMessageHandler);
            this.boundMessageHandler = null;
        }

        // When creating the webview:
        const webview = document.createElement('webview') as WebviewTag;
        this.webview = webview;
    webview.setAttribute('webpreferences', 'contextIsolation=false, nodeIntegration=true');
    webview.setAttribute('partition', 'persist:onlyoffice');
    webview.setAttribute('allowpopups', 'true');
        webview.style.cssText = 'width: 100%; height: 100%; border: none;';
    webview.src = `http://127.0.0.1:${this.plugin.localServerPort}/embedded-editor.html`;
        webviewContainer.appendChild(webview);

        // --- REMOVED ipc-message listener ---

    // --- Use LAN IP or host.docker.internal for document.url ---
    // This is the URL OnlyOffice in Docker will use to fetch the file!
    // If this is a template/new doc, point to bundled Start.docx so the server can actually serve a file.
    const documentUrl = isStartDocx ? `${dockerBaseUrl}/assets/Start.docx` : dockerFileUrl; // NOT hostFileUrl
    const ext = (file?.extension || 'docx').toLowerCase();
    const documentType = ext === 'xlsx' ? 'cell' : ext === 'pptx' ? 'slide' : ext === 'pdf' ? 'word' : 'word'; // pdf opened in word viewer mode

        // Build the base config ON THE HOST (without events) so we can sign exactly what the editor will use
        const baseConfig = {
            documentType: documentType,
            document: {
                fileType: ext || 'docx',
                key: uniqueKey,
                title: file ? file.name : (documentUrl.split('/').pop() || 'Document.docx'),
                url: documentUrl,
                permissions: {
                    edit: true,
                    download: true,
                    print: true,
                    chat: false,
                    comment: true,
                    copy: true,
                    fillForms: true,
                    modifyContent: true,
                    modifyFilter: true,
                    review: true,
                    rename: true
                }
            },
            editorConfig: {
                mode: "edit",
                lang: "en",
                user: {
                    id: "obsidian-static-user",
                    name: "Obsidian User"
                },
                customization: {
                    autosave: false,
                    toolbarNoTabs: false,
                    plugins: true,
                    comments: false,
                    compactChat: false,
                    forcesave: true
                },
                callbackUrl: callbackUrl,
                // Keep these fallbacks as part of signed payload to avoid post-sign changes
                createUrl: `http://127.0.0.1:${this.plugin.localServerPort}/api/createNew?fileType=${documentType}`,
                saveAsUrl: `http://127.0.0.1:${this.plugin.localServerPort}/api/saveAs`,
                coEditing: {
                    mode: "strict",
                    change: false,
                    fastCoAuthoring: false
                }
            },
            width: "100%",
            height: "100%"
        } as any;

        // If a JWT secret is provided, generate tokens:
        // 1) requestToken for request verification, appended to URLs
        // 2) signedToken for the editor config payload
        let signedToken: string | null = null;
        if (this.plugin.settings.jwtSecret && this.plugin.settings.jwtSecret.trim() !== '') {
            try {
                const secret = new TextEncoder().encode(this.plugin.settings.jwtSecret);
                // Generate request-verification token (independent payload)
                let requestToken: string | null = null;
                if (this.plugin.settings.useRequestToken !== false) {
                    requestToken = await new SignJWT({ key: uniqueKey, ts: Math.floor(Date.now() / 1000) })
                        .setProtectedHeader({ alg: 'HS256', typ: 'JWT' })
                        .sign(secret);
                    // Append request token to URLs BEFORE signing the editor config
                    const append = (u: string) => u + (u.includes('?') ? '&' : '?') + 'token=' + encodeURIComponent(requestToken!);
                    baseConfig.document.url = append(baseConfig.document.url);
                    baseConfig.editorConfig.callbackUrl = append(baseConfig.editorConfig.callbackUrl);
                    baseConfig.editorConfig.createUrl = append(baseConfig.editorConfig.createUrl);
                    baseConfig.editorConfig.saveAsUrl = append(baseConfig.editorConfig.saveAsUrl);
                    console.log('OnlyOffice: appended request token to URLs');
                } else {
                    console.log('OnlyOffice: request token appending disabled');
                }

                const payload = {
                    document: baseConfig.document,
                    editorConfig: { ...baseConfig.editorConfig },
                    iat: Math.floor(Date.now() / 1000)
                } as any;
                signedToken = await new SignJWT(payload)
                    .setProtectedHeader({ alg: 'HS256', typ: 'JWT' })
                    .sign(secret);
            } catch (err) {
                this.showEditorError(container, `Failed to initialize editor: ${err.message}`);
                return;
            }
        }

        // --- Inject config and API script URL, then load the simplified HTML ---
        webview.addEventListener('did-finish-load', () => {
            console.log('OnlyOffice webview did-finish-load, injecting init script');
            // Define the API script URL
            const apiScriptUrl = `http://${this.plugin.settings.localServerAddress || 'host.docker.internal'}:${this.plugin.settings.onlyOfficeServerPort}/web-apps/apps/api/documents/api.js`;
            
            console.log("OnlyOffice DEBUG: apiScriptUrl:", apiScriptUrl);
            
            // Pass the local server port and initialize the editor with pre-signed config
            const initScript = `
                window.localServerPort = ${this.plugin.localServerPort};
                const __BASE64 = '${Buffer.from(JSON.stringify(baseConfig)).toString('base64')}';
                const __SIGNED_TOKEN = ${signedToken ? `'${signedToken}'` : 'null'};
                
                // Load the API script dynamically
                const script = document.createElement('script');
                script.src = '${apiScriptUrl}';
                script.onload = function() {
                    console.log('API script loaded');
                    try {
                        const { ipcRenderer } = (typeof require === 'function' ? require('electron') : { ipcRenderer: null });
                        const __sendToHost = (channel, data) => {
                            try {
                                if (ipcRenderer && typeof ipcRenderer.sendToHost === 'function') {
                                    ipcRenderer.sendToHost(channel, data);
                                    return;
                                }
                            } catch (e) {}
                            try {
                                console.log('OO-EVT:' + JSON.stringify({ channel, data }));
                            } catch (e) {}
                        };
                        // Parse base config and set token without mutating signed fields
                        const config = JSON.parse(atob(__BASE64));
                        if (__SIGNED_TOKEN) config.token = __SIGNED_TOKEN;

                        // Add event handlers as functions (per OnlyOffice API: use root-level config.events)
                        config.events = {
                            onDocumentStateChange: function(event) {
                                window.isOnlyOfficeDirty = event.data;
                __sendToHost('onlyoffice-dirty-state', { isDirty: !!event.data });
                            },
                            onDocumentContentChanged: function(event) {
                                window.isOnlyOfficeDirty = true;
                __sendToHost('onlyoffice-dirty-state', { isDirty: true });
                            },
                            onDownloadAs: function(event) {
                                // Fallback: when downloadAs is used, forward to host as Save As
                                try { __sendToHost('onlyoffice-save-as', event && event.data ? event.data : {}); } catch(_) {}
                            },
                            onError: function(event) {
                                try {
                                    console.log('OnlyOffice onError raw event', JSON.stringify(event));
                                    __sendToHost('onlyoffice-error', event && event.data ? event.data : {});
                                    // If the error suggests unsaved changes or callback failure, attempt a silent forcesave
                                    try {
                                        if (window.docEditor && event && event.data && (event.data.errorCode === 6 || event.data.errorCode === 7)) {
                                            console.log('OnlyOffice: attempting silent forcesave after errorCode', event.data.errorCode);
                                            window.docEditor.serviceCommand && window.docEditor.serviceCommand('c:forcesave', '');
                                        }
                                    } catch(e) { console.warn('OnlyOffice: silent forcesave retry failed', e); }
                                } catch(_) {}
                            },
                            onAppReady: function(event) {
                                console.log("OnlyOffice App is ready");
                                try { window.docEditor.serviceCommand('TrackRevisions:Change', false); } catch(e) {}
                                try { window.docEditor.serviceCommand("Mode:Edit", true); } catch(e) {}
                                try { __sendToHost('onlyoffice-app-ready', {}); } catch(e) {}
                            },
                            onRequestCreateNew: function(event) {
                                console.log('OnlyOffice onRequestCreateNew');
                __sendToHost('onlyoffice-create-new', { fileType: (event && event.data && event.data['fileType']) || 'docx' });
                            },
                            onRequestSaveAs: function(event) {
                                console.log('OnlyOffice onRequestSaveAs', event);
                __sendToHost('onlyoffice-save-as', event && event.data ? event.data : {});
                            }
                        };

                        // Also mirror events under editorConfig.events for compatibility with older builds
                        config.editorConfig.events = config.events;

                        console.log('OnlyOffice: events wired', Object.keys(config.events||{}));

                        // Initialize the editor with the constructed config
                        window.docEditor = new window.DocsAPI.DocEditor("placeholder", config);
                        console.log("OnlyOffice editor initialized with dynamic config");
                        // Force-hide any 'Download' / 'Download as' / 'Download copy' menu entries (no user setting)
                        (function hideDownloadMenu(){
                            // Extend labels and include ones with spacing/case variants
                            const LABELS = new Set([
                                'download','download as','download copy','download a copy','save copy','save copy as','download as..','download as...'
                            ]);
                            const matchText = (t) => {
                                if (!t) return false; const s = t.trim().toLowerCase().replace(/\s+/g,' ');
                                if (LABELS.has(s)) return true;
                                // Partial heuristics
                                return /^download as/.test(s) || /^download copy/.test(s);
                            };
                            const hide = () => {
                                try {
                                    const els = Array.from(document.querySelectorAll('span,div,button,a,li'));
                                    els.forEach(el => {
                                        if (el.getAttribute('data-oo-hidden')==='download') return;
                                        const txt = (el.textContent||'');
                                        if (matchText(txt)) {
                                            const container = el.closest('li,div,button') || el;
                                            container.style.display='none';
                                            container.setAttribute('data-oo-hidden','download');
                                        }
                                    });
                                } catch(e){ console.warn('OnlyOffice: hideDownloadMenu scan error', e); }
                            };
                            // Monkey patch downloadAs to no-op (prevent keyboard / programmatic)
                            try {
                                const guardProp = '__ooDownloadDisabled';
                                const patch = () => {
                                    try {
                                        if (window.docEditor && !window.docEditor[guardProp] && window.docEditor.downloadAs) {
                                            const orig = window.docEditor.downloadAs.bind(window.docEditor);
                                            window.docEditor.downloadAs = function(){ console.log('OnlyOffice: downloadAs blocked'); return null; };
                                            window.docEditor[guardProp] = true;
                                        }
                                    } catch(e){}
                                };
                                patch();
                                setInterval(patch, 1500); // ensure after internal reloads
                            } catch(e){ console.warn('OnlyOffice: failed to patch downloadAs', e); }
                            hide();
                            const obs = new MutationObserver(hide);
                            try { obs.observe(document.body,{childList:true,subtree:true}); } catch(_){}
                            // Timed passes; stop observer later
                            [400,800,1600,3000,5000,8000,12000].forEach(t=> setTimeout(hide, t));
                            setTimeout(()=>{ try { obs.disconnect(); } catch(_){} }, 18000);
                        })();
                        
                        // Bridge messages coming from /api/createNew and /api/saveAs HTML pages
                        window.addEventListener('message', function(ev) {
                            try {
                                if (!ev || !ev.data || !ev.data.type) return;
                                if (ev.data.type === 'onlyoffice-save-as') {
                                    __sendToHost('onlyoffice-save-as', ev.data.data || {});
                                } else if (ev.data.type === 'onlyoffice-create-new') {
                                    __sendToHost('onlyoffice-create-new', ev.data);
                                }
                            } catch (_) {}
                        });
                        
                        // (Removed legacy attachEvent attempts to reduce errors)
                    } catch (e) {
                        console.error("Error initializing editor with dynamic config:", e);
                        document.getElementById('placeholder').innerHTML = '<div style="padding: 20px; color: red;">Failed to initialize editor: ' + e.message + '</div>';
                    }
                };
                document.head.appendChild(script);
            `;
            
            // Execute the initialization script in the webview
            webview.executeJavaScript(initScript)
                .then(() => {
                    loadingDiv.remove();
                }).catch(err => {
                    this.showEditorError(container, `Failed to initialize editor: ${err.message}`);
                    loadingDiv.remove();
                });
        });

        webview.addEventListener('did-fail-load', (error) => {
            this.showEditorError(container, `Failed to load editor view: ${error.errorDescription}`);
        });

        // Add error handling for the webview itself:
        webview.addEventListener('console-message', (event) => {
            const msg = event.message || '';
            if (typeof msg === 'string' && msg.startsWith('OO-EVT:')) {
                try {
                    const payload = JSON.parse(msg.substring('OO-EVT:'.length));
                    const channel = payload.channel;
                    const data = payload.data;
                    if (channel === 'onlyoffice-save-as') {
                        if (data && data.url) {
                            // Phase 2: file data available
                            this.systemSaveAsRetryPending = false;
                            this.performingSystemSaveAsFallback = false;
                            this.saveAsDocument(false, data);
                        } else {
                            // Phase 1: user requested Save As via toolbar/menu inside editor
                            if (this.plugin.settings.useSystemSaveAs) {
                                // If OnlyOffice did not supply a URL (common for very new docs), fall back to manual vault-based download.
                                if (!this.performingSystemSaveAsFallback) {
                                    this.performingSystemSaveAsFallback = true;
                                    this.fallbackSystemSaveAsFromVault();
                                }
                            } else {
                                this.saveAsDocument(true);
                            }
                        }
                    } else if (channel === 'onlyoffice-create-new') {
                        this.plugin.openNewDocument();
                    } else if (channel === 'onlyoffice-dirty-state') {
                        this.isDirty = !!(data && data.isDirty);
                    }
                } catch (e) {
                    console.log('[OnlyOffice webview]', event.message);
                }
            } else {
                console.log("[OnlyOffice webview]", event.message);
            }
        });

        // Bridge messages from webview (Docs API events) back to the plugin host
        webview.addEventListener('ipc-message', (event: any) => {
            try {
                if (!event) return;
                const channel = event.channel;
                const arg = (event.args && event.args[0]) || {};
                if (channel === 'onlyoffice-save-as') {
                    new Notice('OnlyOffice: Save As event received');
                    // arg expected: { url, title, ... }
                    this.saveAsDocument(this.plugin.settings.useSystemSaveAs ? false : true, arg);
                } else if (channel === 'onlyoffice-create-new') {
                    new Notice('OnlyOffice: Create New event received');
                    // Open a new OnlyOffice view using template flow
                    const leaf = this.leaf;
                    if (leaf) {
                        leaf.setViewState({ type: VIEW_TYPE_ONLYOFFICE, state: { onlyOfficeTemplate: true } });
                    } else {
                        this.plugin.openNewDocument();
                    }
                } else if (channel === 'onlyoffice-app-ready') {
                    new Notice('OnlyOffice: App Ready');
                } else if (channel === 'onlyoffice-dirty-state') {
                    this.isDirty = !!arg.isDirty;
                } else if (channel === 'onlyoffice-error') {
                    const msg = (arg && (arg.message || arg.code)) ? `${arg.code || ''} ${arg.message || ''}`.trim() : 'Unknown error';
                    console.warn('OnlyOffice error:', arg);
                    new Notice('OnlyOffice error: ' + msg);
                }
            } catch (e) {
                console.error('OnlyOffice: ipc-message handler error', e);
            }
        });
    }

    public onSaveComplete() {
        this.isDirty = false;
        if (this.saveButton) this.saveButton.disabled = true;
        if (this.saveAsButton) this.saveAsButton.disabled = true;
        // The polling mechanism will now handle re-enabling buttons if the user edits again.
        // No need to execute JS here.
    }

    // Load OnlyOffice API script
    private async loadOnlyOfficeAPI(): Promise<void> {
        return new Promise((resolve, reject) => {
            if (window.DocsAPI) {
                resolve();
                return;
            }
            
            // Use the API URL that matches your OnlyOffice server
            const apiScriptUrl = `http://localhost:8080/web-apps/apps/api/documents/api.js`;
            console.log('Loading OnlyOffice API script from:', apiScriptUrl);
            
            const script = document.createElement('script');
            script.src = apiScriptUrl;
            script.onload = () => resolve();
            script.onerror = (e) => {
                console.error('Failed to load OnlyOffice API:', e);
                reject(new Error('Failed to load OnlyOffice API'));
            };
            document.head.appendChild(script);
        });
    }

    // Show error message in the editor container
    private showEditorError(container: HTMLElement, message: string): void {
        container.empty();
        const errorDiv = container.createEl('div', { cls: 'onlyoffice-error' });
        errorDiv.createEl('h2', { text: 'OnlyOffice Editor Error' });
        errorDiv.createEl('p', { text: message });
        
        // Add fallback editor option
        const fallbackButton = errorDiv.createEl('button', { 
            text: 'Use Fallback Editor',
            attr: { style: 'margin-top: 10px; padding: 8px 16px; background: #007ACC; color: white; border: none; border-radius: 4px; cursor: pointer;' }
        });
        fallbackButton.addEventListener('click', () => {
            this.loadFallbackEditor(container);
        });
    }

    // Load a simple fallback editor
    private loadFallbackEditor(container: HTMLElement): void {
        container.empty();
        
        // Create a simple rich text editor
        const editorDiv = container.createEl('div', {
            attr: {
                style: 'width: 100%; height: 100%; display: flex; flex-direction: column; background: white;'
            }
        });
        
        // Add toolbar
        const toolbar = editorDiv.createEl('div', {
            attr: {
                style: 'padding: 8px; border-bottom: 1px solid #ccc; background: #f5f5f5; display: flex; gap: 8px; flex-shrink: 0;'
            }
        });
        
        // Add toolbar buttons
        this.addToolbarButton(toolbar, 'Bold', () => document.execCommand('bold'));
        this.addToolbarButton(toolbar, 'Italic', () => document.execCommand('italic'));
        this.addToolbarButton(toolbar, 'Underline', () => document.execCommand('underline'));
        
        // Add save button
        const saveButton = toolbar.createEl('button', {
            text: 'Save to Obsidian',
            attr: {
                style: 'margin-left: auto; padding: 4px 12px; background: #007ACC; color: white; border: none; border-radius: 4px; cursor: pointer;'
            }
        });
        saveButton.addEventListener('click', () => this.saveFallbackToObsidian());
        
        // Create editor area
        this.fallbackEditor = editorDiv.createEl('div', {
            attr: {
                contenteditable: 'true',
                style: 'flex: 1; padding: 20px; font-family: Arial, sans-serif; font-size: 14px; line-height: 1.5; overflow-y: auto; background: white;'
            }
        });
        this.fallbackEditor.innerHTML = '<p>Start typing your document here...</p>';
        this.fallbackEditor.focus();
    }

    // Add toolbar button helper
    private addToolbarButton(toolbar: HTMLElement, text: string, action: () => void): void {
        const button = toolbar.createEl('button', {
            text: text,
            attr: {
                style: 'padding: 4px 8px; border: 1px solid #ccc; background: white; cursor: pointer; border-radius: 3px;'
            }
        });
        button.addEventListener('click', action);
    }

    // Save fallback editor content to Obsidian
    private async saveFallbackToObsidian(): Promise<void> {
        if (!this.fallbackEditor) {
            new Notice('No content to save');
            return;
        }
        try {
            const htmlContent = this.fallbackEditor.innerHTML;
            const markdownContent = OnlyOfficeDocumentView.htmlToMarkdown(htmlContent);
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
            const fileName = `OnlyOffice Document ${timestamp}.md`;
            await this.app.vault.create(fileName, markdownContent);
            new Notice(`Document saved as: ${fileName}`);
            const file = this.app.vault.getAbstractFileByPath(fileName);
            if (file instanceof TFile) {
                await this.app.workspace.openLinkText(fileName, '');
            }
        } catch (error) {
            console.error('Error saving to Obsidian:', error);
            new Notice('Error saving document: ' + error.message);
        }
    }

    // Static HTML to Markdown converter for fallback editor
    private static htmlToMarkdown(html: string): string {
        let markdown = html
            .replace(/<strong[^>]*>(.*?)<\/strong>/gi, '**$1**')
            .replace(/<b[^>]*>(.*?)<\/b>/gi, '**$1**')
            .replace(/<em[^>]*>(.*?)<\/em>/gi, '*$1*')
            .replace(/<i[^>]*>(.*?)<\/i>/gi, '*$1*')
            .replace(/<u[^>]*>(.*?)<\/u>/gi, '<u>$1</u>')
            .replace(/<p[^>]*>(.*?)<\/p>/gi, '$1\n\n')
            .replace(/<br[^>]*>/gi, '\n')
            .replace(/<div[^>]*>(.*?)<\/div>/gi, '$1\n')
            .replace(/<[^>]*>/g, '') // Remove remaining HTML tags
            .replace(/&nbsp;/g, ' ')
            .replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/\n\s*\n\s*\n/g, '\n\n')
            .trim();
        return markdown || 'Empty document';
    }

    // Add a manual save method as fallback
    private async manualSaveFromEditor(): Promise<void> {
        if (!this.file) {
            // This can happen on unload, so don't show a notice.
            console.log("Cannot save: No file is open.");
            return;
        }

        try {
            const arrayBuffer = await this.getEditorContent();
            if (arrayBuffer) {
                await this.app.vault.modifyBinary(this.file, arrayBuffer);
                new Notice('Document saved successfully');
            } else {
                // Avoid showing a notice if the view is closing.
                if (this.leaf.view) {
                    new Notice("Could not get document from editor. Save failed.");
                }
            }
        } catch (error) {
            console.error("Error in manual save:", error);
            if (this.leaf.view) {
                new Notice("Save failed. " + (error instanceof Error ? error.message : "Please check the console for details."));
            }
        }
    }

    // Add the missing saveAsDocument method with improved functionality
    async saveAsDocument(forceSaveAs = false, saveAsData?: any): Promise<void> {
        // For the fallback editor, save content to Obsidian
        if (this.fallbackEditor) {
            await this.saveFallbackToObsidian();
            return;
        }

        try {
            // If we have data from the OnlyOffice Save As event
            const activeExt = this.currentExt || (this.file?.extension?.toLowerCase()) || 'docx';
            if (saveAsData && saveAsData.url) {
                if (this.plugin.settings.useSystemSaveAs) {
                    // Proactively trigger a download inside the webview so the OS Save dialog appears.
                    let fileName = this.pendingSaveAsName || saveAsData.title || `Document.${activeExt}`;
                    this.pendingSaveAsName = null;
                    if (!fileName.toLowerCase().endsWith('.'+activeExt)) fileName += '.'+activeExt;
                    const js = `(() => { try {\n  const url = ${JSON.stringify(saveAsData.url)};\n  const name = ${JSON.stringify(fileName)};\n  // Create hidden anchor to force download\n  const a = document.createElement('a');\n  a.href = url;\n  a.download = name;\n  a.style.display='none';\n  document.body.appendChild(a);\n  a.click();\n  setTimeout(()=>{ try { a.remove(); } catch(_){} }, 4000);\n} catch(e){ console.error('System Save As injection failed', e); } })();`;
                    this.webview?.executeJavaScript(js).catch(e=>console.error('executeJavaScript download inject failed', e));
                    new Notice('Opening system Save dialog...');
                    return;
                }

                new Notice('Downloading document...');
                // Get the file from the provided URL (internal Save As mode)
                const response = await requestUrl({ url: saveAsData.url, method: 'GET' });
                if (response.status !== 200 || !response.arrayBuffer) throw new Error(`Failed to download file: ${response.status}`);

                // Determine user-preferred filename (user choice > provided title)
                let fileName = this.pendingSaveAsName || saveAsData.title || '';
                this.pendingSaveAsName = null; // consume

                // If missing, ask now
                if (!fileName) fileName = await this.promptForFileName() || '';
                if (!fileName) { new Notice('Save As cancelled'); return; }

                // Ensure .docx extension
                if (!fileName.toLowerCase().endsWith('.'+activeExt)) fileName += '.'+activeExt;

                // If file exists, prompt until unique (or auto-increment if user cancels reprompt)
                let attempt = 0;
                const originalBase = fileName.replace(new RegExp(`\.${activeExt}$`,'i'),'');
                while (this.app.vault.getAbstractFileByPath(fileName)) {
                    attempt++;
                    const alt = `${originalBase} (${attempt}).${activeExt}`;
                    if (attempt === 1) { // Offer user a chance to specify different name only on first conflict
                        const retry = await this.promptForFileName();
                        if (retry) {
                            fileName = retry.toLowerCase().endsWith('.'+activeExt) ? retry : retry + '.'+activeExt;
                            continue; // re-check
                        }
                    }
                    fileName = alt;
                }

                const newFile = await this.app.vault.createBinary(fileName, response.arrayBuffer);
                new Notice(`Document saved as: ${fileName}`);
                await this.plugin.openOnlyOfficeFile(newFile); // Open the new file (generic)
                return;
            }
            // No event data: trigger OnlyOffice built-in downloadAs which will fire onDownloadAs
            if (this.webview) {
                if (!this.plugin.settings.useSystemSaveAs) {
                    if (forceSaveAs) {
                        const chosen = await this.promptForFileName();
                        if (!chosen) { new Notice('Save As cancelled'); return; }
                        this.pendingSaveAsName = chosen;
                    }
                    new Notice('Requesting document data...');
                } else {
                    new Notice('Preparing system download...');
                }
                const format = activeExt === 'pdf' ? 'docx' : activeExt; // OnlyOffice may not allow editing export of pdf; keep docx for unsupported download
                this.webview.executeJavaScript(`try { window.docEditor && window.docEditor.downloadAs && window.docEditor.downloadAs('${format}'); } catch(e){ console.error('downloadAs failed', e); }`);
            } else {
                new Notice('Save As not available (editor not ready)');
            }
        } catch (error) {
            console.error("Error during 'Save As':", error);
            new Notice(`Save As failed: ${error.message}`);
        }
    }

    // Fallback path for system Save As when OnlyOffice doesn't return a URL (new unsaved doc)
    private async fallbackSystemSaveAsFromVault() {
        try {
            if (!this.file) { new Notice('No file to save'); this.performingSystemSaveAsFallback = false; return; }
            // Read latest vault content (might be template copy or after autosave) as ArrayBuffer
            const bin = await this.app.vault.readBinary(this.file);
            // Synthesize a download filename based on current file stem + ' copy'
            const base = this.file.basename || 'Document';
            const ext = this.currentExt || this.file.extension || 'docx';
            const name = base.toLowerCase().endsWith('.'+ext) ? base : base + '.'+ext;
            const outName = name.endsWith('.'+ext) ? name : name + '.'+ext;
            const b64 = Buffer.from(bin).toString('base64');
            const js = `(() => { try {\n const b64='${b64}';\n const byteChars = atob(b64);\n const len = byteChars.length;\n const bytes = new Uint8Array(len);\n for (let i=0;i<len;i++) bytes[i]=byteChars.charCodeAt(i);\n const blob = new Blob([bytes], {type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});\n const a=document.createElement('a');\n a.href=URL.createObjectURL(blob);\n a.download=${JSON.stringify(outName)};\n a.style.display='none';\n document.body.appendChild(a);\n a.click();\n setTimeout(()=>{ try { URL.revokeObjectURL(a.href); a.remove(); } catch(_){} },4000);\n } catch(e){ console.error('fallbackSystemSaveAsFromVault failed', e); } })();`;
            this.webview?.executeJavaScript(js).catch(e=>console.error('fallback executeJavaScript failed', e));
            new Notice('Opening system Save dialog (fallback)');
        } catch (e) {
            console.error('OnlyOffice: fallbackSystemSaveAsFromVault error', e);
            new Notice('System Save As fallback failed');
        } finally {
            this.performingSystemSaveAsFallback = false;
        }
    }

    // New helper method to get content from the editor
    private getEditorContent(forceContentRetrieval = false): Promise<ArrayBuffer | null> {
    // Legacy path retained for possible future direct binary retrieval; currently unused.
    return Promise.resolve(null);
    }

    // Helper to prompt for a new file name
    private async promptForFileName(): Promise<string | null> {
        return new Promise((resolve) => {
            new SaveAsModal(this.app, (result) => {
                resolve(result);
            }).open();
        });
    }

    // Clean up when view is closed
    async onClose(): Promise<void> {
        // Remove the message listener when the view is closed
        if (this.boundMessageHandler) {
            window.removeEventListener('message', this.boundMessageHandler);
            this.boundMessageHandler = null;
        }
        // Clear the polling interval
        if (this.dirtyCheckInterval) {
            clearInterval(this.dirtyCheckInterval);
            this.dirtyCheckInterval = null;
        }
    }
}

class SaveAsModal extends Modal {
    private onSubmit: (result: string | null) => void;
    private input: HTMLInputElement;
    private submitted: boolean = false;

    constructor(app: App, onSubmit: (result: string | null) => void) {
        super(app);
        this.onSubmit = onSubmit;
    }

    onOpen() {
        const { contentEl } = this;
        contentEl.createEl('h2', { text: 'Save As' });
        contentEl.createEl('p', { text: 'Enter the new file name:' });
        this.input = contentEl.createEl('input', { type: 'text', placeholder: 'NewDocument.docx' });

        const saveButton = contentEl.createEl('button', { text: 'Save' });
        saveButton.addEventListener('click', () => {
            const value = this.input.value.trim();
            if (value) {
                this.submitted = true;
                this.onSubmit(value);
                this.close();
            }
        });

        const cancelButton = contentEl.createEl('button', { text: 'Cancel' });
        cancelButton.style.marginLeft = '10px';
        cancelButton.addEventListener('click', () => {
            this.close();
        });
    }

    onClose() {
        if (!this.submitted) {
            this.onSubmit(null);
        }
        this.contentEl.empty();
    }
}

/**
 * OnlyOffice Plugin class
 * Main plugin class for OnlyOffice integration
 */
export default class OnlyOfficePlugin extends Plugin implements IOnlyOfficePlugin {
    settings: OnlyOfficePluginSettings;
    httpServer: http.Server | null = null;
    callbackServer: http.Server | null = null; 
    localServerPort: number = 0; 
    callbackServerPort: number = 0; 
    keyFileMap: Record<string,string> = {};
    saveAsPromise: {
        resolve: (data: ArrayBuffer) => void,
        reject: (reason?: any) => void,
        keyPrefix: string
    } | null = null;

    async onload() {
        await this.loadSettings();
        await this.startInternalServer(); // Start internal HTTP server
        await this.startCallbackServer(); // Start callback server
        
        // Add ribbon icon for creating a new document
        this.addRibbonIcon('file-plus', 'New OnlyOffice Document', () => {
            this.openNewDocument();
        });
        
        // Register custom view type - FIXED for constructor compatibility
        this.registerView(
            VIEW_TYPE_ONLYOFFICE,
            (leaf) => new OnlyOfficeDocumentView(leaf, this)
        );
        
    // Try to register file extensions (docx + xlsx + pptx + pdf). If already registered, add context menu fallbacks.
        const officeExts = ['docx','xlsx','pptx','pdf'];
        const failed: string[] = [];
        for (const ext of officeExts) {
            try {
                this.registerExtensions([ext], VIEW_TYPE_ONLYOFFICE);
                console.log(`OnlyOffice: registered extension .${ext}`);
            } catch (e) {
                failed.push(ext);
                console.warn(`OnlyOffice: failed to register .${ext}:`, (e as any)?.message || e);
            }
        }
        if (failed.length) {
            console.log('OnlyOffice: installing context menu fallback for', failed.join(','));
            this.registerEvent(
                this.app.workspace.on('file-menu', (menu, file) => {
                    if (file instanceof TFile && officeExts.includes(file.extension)) {
                        menu.addItem((item) => {
                            item.setTitle('Open with OnlyOffice')
                                .setIcon('file-text')
                                .onClick(() => { this.openOnlyOfficeFile(file); });
                        });
                    }
                })
            );
        }
        
        // Add commands and settings tab
        this.addCommands();
        this.addSettingTab(new OnlyOfficeSettingTab(this.app, this));
    }

    // Add commands to the command palette
    private addCommands() {
        // Add a command to open a DOCX file in OnlyOffice
        this.addCommand({
            id: 'open-docx-in-onlyoffice',
            name: 'Open DOCX in OnlyOffice',
            checkCallback: (checking: boolean) => {
                // Check if a file is active and is a DOCX file
                const file = this.app.workspace.getActiveFile();
                if (file && file.extension === 'docx') {
                    if (!checking) {
                        this.openDocxFile(file);
                    }
                    return true;
                }
                return false;
            }
        });

        // Command to open XLSX
        this.addCommand({
            id: 'open-xlsx-in-onlyoffice',
            name: 'Open XLSX in OnlyOffice',
            checkCallback: (checking: boolean) => {
                const file = this.app.workspace.getActiveFile();
                if (file && file.extension === 'xlsx') {
                    if (!checking) {
                        this.openOnlyOfficeFile(file);
                    }
                    return true;
                }
                return false;
            }
        });

        // Command to open PPTX
        this.addCommand({
            id: 'open-pptx-in-onlyoffice',
            name: 'Open PPTX in OnlyOffice',
            checkCallback: (checking: boolean) => {
                const file = this.app.workspace.getActiveFile();
                if (file && file.extension === 'pptx') {
                    if (!checking) {
                        this.openOnlyOfficeFile(file);
                    }
                    return true;
                }
                return false;
            }
        });

        // Command to open PDF (viewer mode)
        this.addCommand({
            id: 'open-pdf-in-onlyoffice',
            name: 'Open PDF in OnlyOffice (viewer)',
            checkCallback: (checking: boolean) => {
                const file = this.app.workspace.getActiveFile();
                if (file && file.extension === 'pdf') {
                    if (!checking) {
                        this.openOnlyOfficeFile(file);
                    }
                    return true;
                }
                return false;
            }
        });

        // Debug command: list detected vs on-disk office files
        this.addCommand({
            id: 'onlyoffice-list-office-files',
            name: 'OnlyOffice: List Office Files (debug)',
            callback: async () => {
                try {
                    const vaultFiles = this.app.vault.getFiles().filter(f=>['docx','xlsx','pptx','pdf'].includes(f.extension));
                    console.log('OnlyOffice DEBUG: vault reports', vaultFiles.length, 'office files');
                    vaultFiles.forEach(f=>console.log('  VAULT:', f.path));
                    // Disk scan
                    if (!(this.app.vault.adapter instanceof FileSystemAdapter)) { console.log('OnlyOffice DEBUG: not FS adapter, cannot disk-scan'); return; }
                    const root = this.app.vault.adapter.getBasePath();
                    const found: string[] = [];
                    const exts = new Set(['.docx','.xlsx','.pptx','.pdf']);
                    const walk = (dir: string) => {
                        const entries = fs.readdirSync(dir,{withFileTypes:true});
                        for (const e of entries) {
                            if (e.name.startsWith('.')) continue; // skip hidden
                            const full = path.join(dir,e.name);
                            if (e.isDirectory()) { walk(full); } else {
                                const lower = e.name.toLowerCase();
                                for (const x of exts) { if (lower.endsWith(x)) { found.push(full); break; } }
                            }
                        }
                    };
                    walk(root);
                    console.log('OnlyOffice DEBUG: disk scan found', found.length, 'office files');
                    const rel = (p:string)=> p.replace(/\\/g,'/').substring(root.replace(/\\/g,'/').length+1);
                    const vaultSet = new Set(vaultFiles.map(f=>f.path));
                    const missingInVault = found.map(rel).filter(p=>!vaultSet.has(p));
                    if (missingInVault.length===0) console.log('OnlyOffice DEBUG: no discrepancies');
                    else {
                        console.warn('OnlyOffice DEBUG: files on disk NOT in vault index:', missingInVault);
                    }
                    new Notice('OnlyOffice: debug listing complete (see console)');
                } catch (e) {
                    console.error('OnlyOffice DEBUG: listing failed', e);
                    new Notice('OnlyOffice: listing failed (see console)');
                }
            }
        });

        // Add a command to open an XLSX file in OnlyOffice
        this.addCommand({
            id: 'open-xlsx-in-onlyoffice',
            name: 'Open XLSX in OnlyOffice',
            checkCallback: (checking: boolean) => {
                const file = this.app.workspace.getActiveFile();
                if (file && file.extension === 'xlsx') {
                    if (!checking) {
                        this.openOnlyOfficeFile(file);
                    }
                    return true;
                }
                return false;
            }
        });
        
        // Add a command to create a new OnlyOffice document
        this.addCommand({
            id: 'create-new-onlyoffice-document',
            name: 'Create New OnlyOffice Document',
            callback: () => {
                this.openNewDocument();
            }
        });
    }

    // Open a new blank document in the OnlyOffice view
    async openNewDocument() {
        try {
            // Optionally prompt for name before creating the file
            let requestedName: string | null = null;
            let chosenType: 'docx'|'xlsx'|'pptx' = 'docx';
            if (this.settings?.promptNameOnCreate !== false) {
                requestedName = await new Promise<string | null>((resolve) => {
                    const modal = new (class extends Modal {
                        private input!: HTMLInputElement; submitted=false; private select!: HTMLSelectElement; private wrapper!: HTMLElement; private okBtn!: HTMLButtonElement;
                        onOpen() {
                            const radius='10px';
                            const fieldHeight='40px';
                            const {contentEl} = this;
                            contentEl.empty();
                            contentEl.style.padding = '26px 32px 22px';
                            // Narrow the modal a bit and ensure no horizontal scroll
                            contentEl.style.minWidth = '600px';
                            contentEl.style.maxWidth = '600px';
                            contentEl.style.boxSizing = 'border-box';
                            contentEl.style.overflow = 'hidden';
                            contentEl.style.overflowX = 'hidden';
                            if (contentEl.parentElement) {
                                (contentEl.parentElement as HTMLElement).style.overflowX = 'hidden';
                            }
                            contentEl.createEl('h2', { text: 'New OnlyOffice Document', attr: { style: 'margin:0 0 16px 0; font-weight:600; font-size:22px; letter-spacing:.4px;' }});
                            this.wrapper = contentEl.createEl('div', { attr: { style: 'display:flex; flex-direction:column; gap:16px;' }});
                            // Fixed-width grid so we can align action buttons with right edge of dropdown.
                            // Expand fields to fill available inner content width (600 - 64 padding = 536).
                            const formGap = 18; const col1 = 320; const col2 = 118; const formWidth = col1 + col2 + formGap; // 536
                            const row = this.wrapper.createEl('div', { attr: { style: `display:grid; grid-template-columns: ${col1}px ${col2}px; gap:${formGap}px; align-items:end; width:${formWidth}px;` }});
                            const nameCol = row.createEl('div', { attr: { style: 'display:flex; flex-direction:column; gap:6px; min-width:0;' }});
                            nameCol.createEl('label', { text: 'Filename', attr: { style: 'font-size:11px; font-weight:600; letter-spacing:.5px; text-transform:uppercase; color:var(--text-muted);' }});
                            this.input = nameCol.createEl('input', { type: 'text', placeholder: 'MyDocument', attr: { style: `padding:8px 12px; font-size:14px; line-height:20px; height:${fieldHeight}; border:1px solid var(--background-modifier-border); border-radius:${radius}; width:100%; max-width:${col1}px; box-sizing:border-box;` }});
                            const typeCol = row.createEl('div', { attr: { style: 'display:flex; flex-direction:column; gap:6px;'} });
                            typeCol.createEl('label', { text: 'Type', attr: { style: 'font-size:11px; font-weight:600; letter-spacing:.5px; text-transform:uppercase; color:var(--text-muted);' }});
                            this.select = typeCol.createEl('select', { attr: { style: `padding:8px 12px; font-size:14px; line-height:20px; height:${fieldHeight}; border:1px solid var(--background-modifier-border); border-radius:${radius}; background:var(--background-primary); box-sizing:border-box; width:${col2}px;` }});
                            ['docx','xlsx','pptx'].forEach(t => { const o = document.createElement('option'); o.value = t; o.textContent = t.toUpperCase(); this.select.appendChild(o); });
                            const btnRow = this.wrapper.createEl('div', { attr: { style: `display:flex; justify-content:flex-end; gap:10px; margin-top:6px; width:${formWidth}px;` }});
                            const baseBtn = `padding:8px 20px; font-size:14px; font-weight:500; border:none; border-radius:${radius}; cursor:pointer; line-height:20px;`;
                            const cancelBtn = btnRow.createEl('button', { text: 'Cancel', attr: { style: baseBtn + 'background:var(--background-modifier-border); color:var(--text-normal);' }});
                            this.okBtn = btnRow.createEl('button', { text: 'Create', attr: { style: baseBtn + 'background:var(--interactive-accent); color:var(--text-on-accent,#fff);' }});
                            this.okBtn.addEventListener('click', () => { const v = this.input.value.trim(); if (v) { (chosenType as any) = this.select.value; this.submitted = true; this.close(); resolve(v); } });
                            cancelBtn.addEventListener('click', () => { this.close(); resolve(null); });
                            this.input.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); this.okBtn.click(); }});
                            this.input.focus();
                        }
                        onClose() { if (!this.submitted) resolve(null); this.contentEl.empty(); }
                    })(this.app);
                    modal.open();
                });
                if (requestedName === null) { new Notice('Creation cancelled'); return; }
                requestedName = requestedName.replace(/[\\/:*?"<>|]/g,'').trim();
                if (!requestedName) { new Notice('Invalid name'); return; }
                const extRegex = new RegExp(`\.${chosenType}$`,'i');
                if (!extRegex.test(requestedName)) requestedName += `.${chosenType}`;
                if (this.app.vault.getAbstractFileByPath(requestedName)) { // uniqueness
                    let base = requestedName.replace(new RegExp(`\.${chosenType}$`,'i'),'');
                    let i=1; let candidate = `${base} (${i}).${chosenType}`;
                    while (this.app.vault.getAbstractFileByPath(candidate)) { i++; candidate = `${base} (${i}).${chosenType}`; }
                    requestedName = candidate;
                }
            }
            // Resolve a template matching chosenType (Start.docx / Start.xlsx / Start.pptx)
            let templateData: ArrayBuffer | undefined;
            const adapter = this.app.vault.adapter as FileSystemAdapter;
            const base = adapter.getBasePath();
            const pluginDir = this.manifest.dir && path.isAbsolute(this.manifest.dir) ? this.manifest.dir : path.join(base, this.manifest.dir || '');
            const wantedExt = chosenType; // map same names
            const candidateNames = wantedExt === 'docx'
                ? ['Start.docx','Start.dotx']
                : wantedExt === 'xlsx'
                    ? ['Start.xlsx']
                    : ['Start.pptx'];
            const candidatePaths: string[] = [];
            for (const nm of candidateNames) {
                candidatePaths.push(path.join(pluginDir,'assets',nm));
                candidatePaths.push(path.join(base,nm));
            }
            for (const p of candidatePaths) {
                if (fs.existsSync(p)) {
                    const buf = await fs.promises.readFile(p);
                    const slice = buf.subarray(0); // Uint8Array
                    const copy = new Uint8Array(slice.length);
                    copy.set(slice);
                    templateData = copy.buffer; // Pure ArrayBuffer
                    console.log('OnlyOffice: using template', p, 'for new', wantedExt);
                    break;
                }
            }
            // If no template found for docx we embed minimal; for xlsx/pptx we abort to avoid mismatch
            if (!templateData && wantedExt === 'docx') {
                try {
                    const minimalB64 = 'UEsDBBQABgAIAAAAIQAAAAAAAAAAAAAAAAAJAAAAd29yZC9VVAkAA0oJZ2ZKCWdzdXgLAAEE9QEAAAQUAAAAAABQSwMEFAAIAAgAAAAhAAAAAAAAAAAAAAAAABwAAAB3b3JkL19yZWxzL2RvY3VtZW50LnhtbC5yZWxzVVQJAAOSAmdmkgJndXgLAAEE9QEAAAQUAAAAAABQSwMEFAAIAAgAAAAhAAAAAAAAAAAAAAAAADwAAAB3b3JkL2RvY3VtZW50LnhtbFVUCQADjQJnZ40CZ3V4CwABBPUBAAAQFAAAAAA8eG1sIHZlcnNpb249IjEuMCI+PHc6d29yZGRvYyB4bWxucz13PSJodHRwOi8vc2NoZW1hcy5vcGVub3hwb3J5Lm9yZy93b3JkcHJvYy8yMDA2L21haW4iPjx3OnNlY3Rpb24geG1sbnM6dyA9ICJodHRwOi8vc2NoZW1hcy5vcGVub3hwb3J5Lm9yZy93b3JkcHJvYy8yMDA2L21haW4iPjx3OnA+PC93OnA+PC93OnNlY3Rpb24+PC93OndvcmRkb2M+';
                    const buf = Buffer.from(minimalB64, 'base64');
                    const copy = new Uint8Array(buf.length); copy.set(buf); templateData = copy.buffer;
                    console.log('OnlyOffice: using embedded minimal DOCX fallback (no template found)');
                } catch(e) {
                    console.error('OnlyOffice: failed to build minimal DOCX fallback, creating empty file', e);
                    templateData = new ArrayBuffer(0);
                }
            } else if (!templateData) {
                new Notice(`No template found for ${wantedExt.toUpperCase()} (place Start.${wantedExt} in plugin assets). Creation cancelled.`);
                return;
            }
            // Determine final filename
            let finalName: string;
            if (requestedName) {
                finalName = requestedName;
            } else {
                let baseName = 'Untitled Document';
                let idx = 1;
                finalName = `${baseName}.${wantedExt}`;
                while (this.app.vault.getAbstractFileByPath(finalName)) { idx++; finalName = `${baseName} ${idx}.${wantedExt}`; }
            }
            const file = await this.app.vault.createBinary(finalName, templateData);
            console.log('OnlyOffice: created new document', finalName, 'size', templateData.byteLength, 'prompted?', !!requestedName);
            await this.openOnlyOfficeFile(file);
        } catch (e) {
            console.error('OnlyOffice: failed to create new document', e);
            new Notice('OnlyOffice: Failed to create new document');
    }
    }


    // Start an internal HTTP server to serve vault files and editor.html
    async startInternalServer() {
        if (!(this.app.vault.adapter instanceof FileSystemAdapter)) {
            new Notice("OnlyOffice plugin requires a file system adapter to run the local server.");
            return;
        }
        const vaultPath = this.app.vault.adapter.getBasePath();
        let pluginPath = this.manifest.dir;
        if (pluginPath && !path.isAbsolute(pluginPath)) {
            pluginPath = path.resolve(vaultPath, pluginPath);
        }
        const resolvedPluginPath = pluginPath && path.isAbsolute(pluginPath) ? pluginPath : __dirname;
        const editorHtmlPath = path.join(resolvedPluginPath, 'editor.html');

        // Debug paths to verify they're correct
        console.log("Vault path:", vaultPath);
        console.log("Plugin path:", pluginPath);
        console.log("Resolved plugin path:", resolvedPluginPath);
        console.log("Editor HTML path:", editorHtmlPath);
        console.log("Editor HTML exists:", fs.existsSync(editorHtmlPath));

        // Use configured port unless 0 (then start at 8081 dynamically)
        let port = (this.settings.htmlServerPort === 0 ? 8081 : (this.settings.htmlServerPort || 8081));
        const maxPort = port + 50;
        let serverStarted = false;

        while (port < maxPort && !serverStarted) {
            try {
                this.httpServer = http.createServer((req, res) => {
                    // --- LOG EVERY REQUEST ---
                    console.log("[HTTP] Request URL:", req.url);

                    // Add CORS headers to all responses
                    res.setHeader('Access-Control-Allow-Origin', '*');
                    res.setHeader('Access-Control-Allow-Methods', 'GET, HEAD, POST, OPTIONS');
                    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
                    
                    // Handle preflight OPTIONS requests
                    if (req.method === 'OPTIONS') {
                        res.writeHead(200);
                        res.end();
                        return;
                    }

                    try {
                        let reqPath = decodeURIComponent(req.url?.split('?')[0] || '/');

                        // Handle API endpoint for creating new documents
                        if (reqPath === '/api/createNew') {
                            console.log('[HTTP] Create New document request');
                            res.writeHead(200, { 'Content-Type': 'text/html' });
                            res.end(`<!DOCTYPE html><html><head><script>window.parent.postMessage({type:'onlyoffice-create-new'},'*');setTimeout(()=>window.close(),100);</script></head><body>Creating new document...</body></html>`);
                            return;
                        }
                        // Handle API endpoint for Save As functionality
                        if (reqPath === '/api/saveAs') {
                            console.log('[HTTP] Save As request');
                            const urlObj = new URL(req.url || '', `http://${req.headers.host || 'localhost'}`);
                            const url = urlObj.searchParams.get('url') || '';
                            const title = urlObj.searchParams.get('title') || '';
                            res.writeHead(200, { 'Content-Type': 'text/html' });
                            res.end(`<!DOCTYPE html><html><head><script>window.parent.postMessage({type:'onlyoffice-save-as',data:{url:'${url}',title:'${title}',format:'docx'}},'*');setTimeout(()=>window.close(),100);</script></head><body>Saving document as...</body></html>`);
                            return;
                        }

                        if (reqPath === '/' || reqPath === '/editor.html') {
                            console.log('Serving editor.html from:', editorHtmlPath);
                            if (!fs.existsSync(editorHtmlPath)) {
                                res.writeHead(404, { 'Content-Type': 'text/plain' });
                                res.end('editor.html not found');
                                return;
                            }
                            fs.readFile(editorHtmlPath, (err, data) => {
                                if (err) {
                                    res.writeHead(500, { 'Content-Type': 'text/plain' });
                                    res.end('Error reading editor.html');
                                    return;
                                }
                                res.writeHead(200, {
                                    'Content-Type': 'text/html; charset=UTF-8',
                                    'Content-Length': Buffer.byteLength(data),
                                    'Cache-Control': 'no-store, no-cache, must-revalidate, proxy-revalidate',
                                    'Pragma': 'no-cache',
                                    'Expires': '0',
                                    'Surrogate-Control': 'no-store',
                                });
                                res.end(data);
                            });
                        } else if (reqPath === '/embedded-editor.html') {
                            const embeddedPath = path.join(resolvedPluginPath, 'embedded-editor.html');
                            if (!fs.existsSync(embeddedPath)) {
                                res.writeHead(404, { 'Content-Type': 'text/plain' });
                                res.end('embedded-editor.html not found');
                                return;
                            }
                            fs.readFile(embeddedPath, (err, data) => {
                                if (err) {
                                    res.writeHead(500, { 'Content-Type': 'text/plain' });
                                    res.end('Error reading embedded-editor.html');
                                    return;
                                }
                                res.writeHead(200, { 'Content-Type': 'text/html; charset=UTF-8', 'Cache-Control': 'no-store' });
                                res.end(data);
                            });
                        } else {
                            const filePath = path.join(vaultPath, reqPath.replace(/^\//, ''));
                            console.log("Attempting to serve file:", filePath);
                            console.log("File exists:", fs.existsSync(filePath));
                            if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
                                const mimeType = mime.lookup(filePath) || 'application/octet-stream';
                                res.writeHead(200, {
                                    'Content-Type': String(mimeType),
                                    'Access-Control-Allow-Origin': '*',
                                    'Cache-Control': 'no-cache'
                                });
                                fs.createReadStream(filePath).pipe(res);
                            } else {
                                const pluginFilePath = path.join(resolvedPluginPath, reqPath.replace(/^\//, ''));
                                if (fs.existsSync(pluginFilePath) && fs.statSync(pluginFilePath).isFile()) {
                                    const mimeType = mime.lookup(pluginFilePath) || 'application/octet-stream';
                                    res.writeHead(200, {
                                        'Content-Type': String(mimeType),
                                        'Access-Control-Allow-Origin': '*',
                                        'Cache-Control': 'no-cache'
                                    });
                                    fs.createReadStream(pluginFilePath).pipe(res);
                                } else {
                                    res.writeHead(404, { 'Content-Type': 'text/plain' });
                                    res.end(`File not found: ${filePath}`);
                                }
                            }
                        }
                    } catch (e) {
                        console.error("Server error:", e);
                        res.writeHead(500, { 'Content-Type': 'text/plain' });
                        res.end('Internal server error: ' + e.message);
                    }
                });

                await new Promise<void>((resolve, reject) => {
                    this.httpServer!.once('error', (err: any) => reject(err));
                    this.httpServer!.listen(port, '0.0.0.0', () => resolve());
                });

                this.localServerPort = port;
                serverStarted = true;
                const os = require('os');
                const interfaces = os.networkInterfaces();
                let lanIps: string[] = [];
                for (const name of Object.keys(interfaces)) {
                    for (const iface of interfaces[name]!) {
                        if (iface.family === 'IPv4' && !iface.internal) {
                            lanIps.push(iface.address);
                        }
                    }
                }
                console.log(`OnlyOffice internal HTML server running at:`);
                console.log(`  Local:   http://127.0.0.1:${port}/`);
                lanIps.forEach(ip => console.log(`  LAN:     http://${ip}:${port}/`));
                try {
                    const testResponse = await requestUrl({
                        url: `http://127.0.0.1:${port}/editor.html`,
                        method: 'HEAD',
                        headers: { 'Cache-Control': 'no-cache' }
                    });
                    console.log("Server test result:", testResponse.status);
                } catch (err) {
                    console.error("Server test failed:", err);
                }
            } catch (err: any) {
                console.log(`Port ${port} is not available:`, err.message);
                port++;
            }
        }
        if (!serverStarted) {
            new Notice(`OnlyOffice plugin could not find an open port between 8081 and ${maxPort}.`);
        }
    }
    // Start an internal OnlyOffice callback server
    async startCallbackServer() {
    // Pick a configured port unless 0 (dynamic starting at 8082)
    let port = (this.settings.callbackServerPort === 0 ? 8082 : (this.settings.callbackServerPort || 8082));
    const maxPort = port + 1000; // Try many ports if the preferred one is busy
        let serverStarted = false;
        let lastError = null;

        while (port < maxPort && !serverStarted) {
            try {
                this.callbackServer = http.createServer(async (req, res) => {
                    // Add CORS headers to all responses
                    res.setHeader('Access-Control-Allow-Origin', '*');
                    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
                    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
                    
                    // Handle preflight OPTIONS requests
                    if (req.method === 'OPTIONS') {
                        res.writeHead(200);
                        res.end();
                        return;
                    }
                    
                    // Log all requests for debugging
                    console.log(`[Callback] ${req.method} ${req.url}`);
                    
                    // Normalize path (strip query) so /callback?token=... is accepted
                    const rawUrl = req.url || '/';
                    const pathOnly = rawUrl.split('?')[0];
                    if (req.method === 'POST' && pathOnly === '/callback') {
                        let body = '';
                        req.on('data', chunk => { body += chunk; });
                        req.on('end', async () => {
                            try {
                                console.log('[Callback] Received data:', body);
                                let data;
                                try {
                                    data = JSON.parse(body);
                                } catch (e) {
                                    console.error('[Callback] Failed to parse JSON:', e);
                                    // Try to handle non-JSON data
                                    if (body.includes('url=') && body.includes('&')) {
                                        // Handle form-encoded data
                                        const params = new URLSearchParams(body);
                                        data = Object.fromEntries(params);
                                    } else {
                                        res.writeHead(200, { 'Content-Type': 'application/json' });
                                        res.end(JSON.stringify({ error: 0 }));
                                        return;
                                    }
                                }
                                
                                // OnlyOffice callback: status 2 = must save, status 6 = must force save
                                // Also handle cases where status might be a string or missing
                                const status = parseInt(data.status) || 0;
                                
                                if (data && (status === 2 || status === 6 || status === 0) && data.url) {
                                    console.log('[Callback] Processing save request with URL:', data.url);
                                    
                                    // Extract key from data or URL
                                    let key = data.key;
                                    if (!key && data.url) {
                                        // Try to extract key from URL if not in data
                                        try {
                                            const urlObj = new URL(data.url);
                                            key = urlObj.searchParams.get('key');
                                        } catch (e) {
                                            console.error('[Callback] Error parsing URL:', e);
                                        }
                                    }
                                    
                                    if (key) {
                                        console.log('[Callback] Using key:', key);
                                        let filePath: string | null = null;
                                        // New mapping first
                                        if (this.keyFileMap && this.keyFileMap[key]) {
                                            filePath = this.keyFileMap[key];
                                            console.log('[Callback] Resolved file path from map:', filePath);
                                        } else {
                                            // Legacy pattern support
                                            const match = /^obsidian_(.+?)_\d+_[a-z0-9]+$/.exec(key);
                                            if (match) {
                                                filePath = decodeURIComponent(match[1]);
                                                console.log('[Callback] Extracted legacy file path:', filePath);
                                            }
                                        }
                                        if (filePath) {
                                            const file = this.app.vault.getAbstractFileByPath(filePath);
                                            if (file instanceof TFile) {
                                                // Download the new file from OnlyOffice and overwrite in vault
                                                console.log('[Callback] Downloading from URL:', data.url);
                                                try {
                                                    const response = await requestUrl({ 
                                                        url: data.url, 
                                                        method: 'GET',
                                                        headers: { 'Cache-Control': 'no-cache' }
                                                    });
                                                    if (response.status === 200 && response.arrayBuffer) {
                                                        await this.app.vault.modifyBinary(file, response.arrayBuffer);
                                                        new Notice('OnlyOffice: Document saved from callback');
                                                        console.log('[Callback] Document saved successfully');
                                                    } else {
                                                        console.error('[Callback] Failed to download file, status:', response.status);
                                                    }
                                                } catch (err) {
                                                    console.error('[Callback] Error downloading file:', err);
                                                }
                                            } else {
                                                console.error('[Callback] File not found in vault:', filePath);
                                            }
                                        } else {
                                            console.error('[Callback] Could not resolve file path for key:', key);
                                        }
                                    } else {
                                        console.error('[Callback] No key found in callback data');
                                    }
                                } else {
                                    console.log('[Callback] Received non-save status or missing URL:', data);
                                }
                                
                                // Always return success to OnlyOffice
                                res.writeHead(200, { 'Content-Type': 'application/json' });
                                res.end(JSON.stringify({ error: 0 }));
                            } catch (err) {
                                console.error('[Callback] Error processing callback:', err);
                                // Still return success to OnlyOffice to avoid errors
                                res.writeHead(200, { 'Content-Type': 'application/json' });
                                res.end(JSON.stringify({ error: 0 }));
                            }
                        });
                    } else {
                        // Health check or mismatched path: still return JSON shape OnlyOffice expects
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: 0, info: 'OnlyOffice callback server is running' }));
                    }
                });

                await new Promise<void>((resolve, reject) => {
                    const timeoutId = setTimeout(() => {
                        reject(new Error("Server start timed out"));
                    }, 3000);
                    
                    this.callbackServer!.once('error', (err: any) => {
                        clearTimeout(timeoutId);
                        reject(err);
                    });
                    
                    this.callbackServer!.listen(port, '0.0.0.0', () => {
                        clearTimeout(timeoutId);
                        resolve();
                    });
                });

                this.callbackServerPort = port;
                serverStarted = true;
                console.log(`OnlyOffice callback server running on port ${port}`);
                
                // Test if the server is actually accessible
                try {
                    const testResponse = await requestUrl({
                        url: `http://127.0.0.1:${port}/`,
                        method: 'HEAD',
                        headers: { 'Cache-Control': 'no-cache' }
                    });
                    console.log("Callback server test result:", testResponse.status);
                    new Notice(`OnlyOffice callback server running on port ${port} (status: ${testResponse.status})`);
                } catch (err) {
                    console.warn("Callback server test failed, but server appears to be running:", (err as Error).message);
                    new Notice(`OnlyOffice callback server started on port ${port}, but test failed: ${(err as Error).message}`);
                }
            } catch (err: any) {
                lastError = err;
                console.log(`Port ${port} is not available: ${err.message}`);
                
                // Close the server if it was created but failed to start
                if (this.callbackServer) {
                    try {
                        this.callbackServer.close();
                    } catch (e) {
                        // Ignore errors when closing
                    }
                    this.callbackServer = null;
                }
                
                port++;
            }
        }

        if (!serverStarted) {
            console.error(`OnlyOffice callback server could not find an open port between 8082 and ${maxPort}`, lastError);
            new Notice(`OnlyOffice plugin could not start the callback server. Some features may not work correctly.`);
            
            // Set a dummy port so the rest of the plugin can function
            this.callbackServerPort = 0;
        }
    }

    // Load settings from disk
    async loadSettings() {
        this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
    }

    // Save settings to disk
    async saveSettings() {
        await this.saveData(this.settings);
    }

    // Clean up when the plugin is unloaded
    async onunload() {
        if (this.httpServer) {
            this.httpServer.close();
            this.httpServer = null;
        }
        if (this.callbackServer) {
            this.callbackServer.close();
            this.callbackServer = null;
        }
    }

    // Open a DOCX file in the OnlyOffice view
    async openDocxFile(file: TFile): Promise<void> {
        // Use the standard Obsidian API to open the file in a new or existing leaf,
        // but force the OnlyOffice view type and set the file path in the view state.
        const leaf = this.app.workspace.getLeaf(true);
        await leaf.setViewState({
            type: VIEW_TYPE_ONLYOFFICE,
            state: { file: file.path }
        });
        this.app.workspace.setActiveLeaf(leaf, { focus: true });
    }

    // Generic opener for any supported OnlyOffice extension
    async openOnlyOfficeFile(file: TFile): Promise<void> {
        return this.openDocxFile(file); // same mechanics; view logic branches by extension
    }
}

/**
 * Settings tab for OnlyOffice plugin
 */
class OnlyOfficeSettingTab extends PluginSettingTab {
    plugin: OnlyOfficePlugin;

    constructor(app: App, plugin: OnlyOfficePlugin) {
        super(app, plugin);
        this.plugin = plugin;
    }

    display(): void {
        const { containerEl } = this;
        containerEl.empty();

        new Setting(containerEl)
            .setName('OnlyOffice Server Port')
            .setDesc('The port where your OnlyOffice Document Server is running.')
            .addText(text => text
                .setPlaceholder('8080')
                .setValue(this.plugin.settings.onlyOfficeServerPort.toString())
                .onChange(async (value) => {
                    this.plugin.settings.onlyOfficeServerPort = Number(value) || 8080;
                    await this.plugin.saveSettings();
                }));

        new Setting(containerEl)
            .setName('JWT Secret')
            .setDesc('The secret key for JWT token generation (leave empty to disable).')
            .addText(text => text
                .setPlaceholder('your-secret-key')
                .setValue(this.plugin.settings.jwtSecret || '')
                .onChange(async (value) => {
                    this.plugin.settings.jwtSecret = value;
                    await this.plugin.saveSettings();
                }));

        new Setting(containerEl)
            .setName('Local Server Address for Docker')
            .setDesc('The address of your machine that the OnlyOffice Docker container can reach (e.g., host.docker.internal, or your LAN IP).')
            .addText(text => (text
                .setPlaceholder('host.docker.internal')
                .setValue(this.plugin.settings.localServerAddress || '')
                .onChange(async (value) => {
                    this.plugin.settings.localServerAddress = value;
                    await this.plugin.saveSettings();
                })));

        new Setting(containerEl)
            .setName('Embedded HTML Server Port')
            .setDesc('The port for the plugin’s internal HTML server (default 8081). Restart plugin after changing.')
            .addText(text => text
                .setPlaceholder('8081')
                .setValue(String(this.plugin.settings.htmlServerPort || 8081))
                .onChange(async (value) => {
                    const v = Number(value) || 8081;
                    this.plugin.settings.htmlServerPort = v;
                    await this.plugin.saveSettings();
                }));

        new Setting(containerEl)
            .setName('Callback Server Port')
            .setDesc('The port for the OnlyOffice callback server (default 8082). Restart plugin after changing.')
            .addText(text => text
                .setPlaceholder('8082')
                .setValue(String(this.plugin.settings.callbackServerPort || 8082))
                .onChange(async (value) => {
                    const v = Number(value) || 8082;
                    this.plugin.settings.callbackServerPort = v;
                    await this.plugin.saveSettings();
                }));

        new Setting(containerEl)
            .setName('Append request token to URLs')
            .setDesc('Toggle adding ?token=... to document and callback URLs (disable if your server does not use request verification).')
            .addToggle(toggle => toggle
                .setValue(this.plugin.settings.useRequestToken ?? true)
                .onChange(async (value) => {
                    this.plugin.settings.useRequestToken = value;
                    await this.plugin.saveSettings();
                }));

        new Setting(containerEl)
            .setName('Use system Save As UI')
            .setDesc('Open the document copy in your default browser (download UI) instead of Obsidian modal when using Save As / Download Copy.')
            .addToggle(toggle => toggle
                .setValue(this.plugin.settings.useSystemSaveAs ?? false)
                .onChange(async (value) => {
                    this.plugin.settings.useSystemSaveAs = value;
                    await this.plugin.saveSettings();
                }));
    }
}
