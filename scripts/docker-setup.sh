#!/bin/bash
# word-agent Docker setup script
# Builds the Docker image and optionally starts the container.
#
# Usage:
#   bash scripts/docker-setup.sh          # build only
#   bash scripts/docker-setup.sh --run    # build and run

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

cd "$PROJECT_DIR"

IMAGE_NAME="word-agent"
CONTAINER_NAME="word-agent-dev"

echo "Building word-agent Docker image..."
docker build -t "$IMAGE_NAME" .

echo "Image built successfully: $IMAGE_NAME"

if [ "$1" = "--run" ]; then
    echo "Starting container..."

    # Create documents directory if not exists
    mkdir -p documents

    docker run -it --rm \
        --name "$CONTAINER_NAME" \
        -v "$(pwd)/documents:/documents" \
        -e MCP_AUTHOR=Claude \
        "$IMAGE_NAME" bash

    echo "Container stopped."
else
    echo ""
    echo "To run the container:"
    echo "  docker run -it --rm -v \$(pwd)/documents:/documents $IMAGE_NAME bash"
    echo ""
    echo "Or use docker compose:"
    echo "  docker compose up -d"
fi
