# OnlyOffice Server Configuration Guide

## Required Server Settings for Obsidian Integration

For OnlyOffice to work properly within Obsidian, you need to configure your OnlyOffice Document Server with specific headers to allow iframe embedding:

### Nginx Configuration

Add these headers to your Nginx server block:

```nginx
# Allow embedding in iframes (required for Obsidian)
add_header X-Frame-Options "ALLOWALL";
add_header Content-Security-Policy "frame-ancestors 'self' app://obsidian.md";

# Enable CORS for Obsidian
add_header 'Access-Control-Allow-Origin' 'app://obsidian.md';
add_header 'Access-Control-Allow-Methods' 'GET, POST, OPTIONS, PUT, DELETE';
add_header 'Access-Control-Allow-Headers' 'DNT,X-CustomHeader,Keep-Alive,User-Agent,X-Requested-With,If-Modified-Since,Cache-Control,Content-Type,Authorization';
add_header 'Access-Control-Allow-Credentials' 'true';
```

### Docker Configuration

If using Docker, add these environment variables:

```bash
docker run -i -t -d -p 8080:80 \
-e JWT_ENABLED=true \
-e JWT_SECRET=your-jwt-secret \
-e WOPI_ENABLED=true \
-e ONLYOFFICE_NGINX_ACCESS_CONTROL_ALLOW_ORIGIN="app://obsidian.md" \
-e ONLYOFFICE_NGINX_ACCESS_CONTROL_ALLOW_METHODS="GET, POST, OPTIONS, PUT, DELETE" \
-e ONLYOFFICE_NGINX_ACCESS_CONTROL_ALLOW_HEADERS="Content-Type, Authorization" \
-e ONLYOFFICE_NGINX_ACCESS_CONTROL_ALLOW_CREDENTIALS="true" \
onlyoffice/documentserver
```

## Troubleshooting

If editing is still disabled (grayed-out tools):

1. Use "useTestToken" in the plugin settings to test with the known working token
2. Try the "Open in External Browser" button to confirm OnlyOffice works outside Obsidian
3. Check Obsidian's developer console (Ctrl+Shift+I) for any security or CORS errors
