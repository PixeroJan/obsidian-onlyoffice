#!/bin/bash
# Script to start OnlyOffice Document Server with settings optimized for Obsidian

# Check if Docker is installed
if ! command -v docker &> /dev/null
then
    echo "Docker command not found. Please install a Docker environment to continue."
    echo "We recommend Docker Desktop: https://www.docker.com/products/docker-desktop/"
    read -p "Press any key to exit..."
    exit 1
fi

CONTAINER_NAME="obsidian-onlyoffice"
JWT_SECRET="your-secret-key-please-change"
HOST_PORT=8080

echo "--- OnlyOffice Docker Manager ---"
echo ""
echo "Checking for container '$CONTAINER_NAME'..."

if [ "$(docker ps -a -q -f name=$CONTAINER_NAME)" ]; then
    echo "Container found."
    if [ "$(docker ps -q -f name=$CONTAINER_NAME)" ]; then
        echo "Container is already running."
    else
        echo "Container is stopped. Starting..."
        docker start $CONTAINER_NAME
    fi
else
    echo "Container not found. Creating and starting a new one..."
    echo ""
    echo "IMPORTANT: The server is being created with the following secret key:"
    echo "\"$JWT_SECRET\""
    echo "You must use this same key in the Obsidian plugin settings."
    echo ""
    
    docker run -d --name $CONTAINER_NAME -p $HOST_PORT:80 --restart always --add-host=host.docker.internal:host-gateway -e JWT_ENABLED=true -e JWT_SECRET=$JWT_SECRET onlyoffice/documentserver
fi

echo ""
echo "--- Operation Complete ---"
read -p "Press any key to continue..."
    fi
    
    echo ""
    echo "Creating and starting a new container..."
    docker run -d --name $CONTAINER_NAME -p $HOST_PORT:80 --restart always --add-host=host.docker.internal:host-gateway -e JWT_ENABLED=true -e JWT_SECRET=$JWT_SECRET onlyoffice/documentserver
fi

echo ""
echo "--- Operation Complete ---"
read -p "Press any key to continue..."
fi

echo ""
echo "--- Operation Complete ---"
read -p "Press any key to continue..."
