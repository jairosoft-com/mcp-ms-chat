# MCP Server Docker Setup

This document explains how to build and run the MCP server in a Docker container.

## Prerequisites

- Docker installed on your system
- Node.js and npm (for local development without Docker)

## Building the Docker Image

```bash
docker build -t mcp-calendar-server .
```

## Running the Container

### Basic Usage

```bash
docker run -p 8787:8787 -e AUTH_TOKEN=your_auth_token mcp-calendar-server
```

### Environment Variables

- `AUTH_TOKEN`: (Required) Authentication token for the MCP server
- `NODE_ENV`: Set to `production` or `development` (default: `production`)

### Example with Environment File

1. Create a `.env` file:
   ```
   AUTH_TOKEN=your_auth_token_here
   NODE_ENV=development
   ```

2. Run the container:
   ```bash
   docker run -p 8787:8787 --env-file .env mcp-calendar-server
   ```

## Accessing the Server

- MCP Endpoint: `http://localhost:8787/mcp`
- Server-Sent Events: `http://localhost:8787/sse`

## Development

For local development without Docker:

1. Install dependencies:
   ```bash
   npm ci
   ```

2. Set environment variables:
   ```bash
   export AUTH_TOKEN=your_auth_token
   ```

3. Start the development server:
   ```bash
   npm run dev
   ```

## Notes

- The server runs in development mode by default when using `npm run dev`
- In production, ensure proper logging and monitoring are set up
- The container exposes port 8787 by default
