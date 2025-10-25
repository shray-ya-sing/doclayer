# DocLayer Remote MCP Server

Remote MCP server for DocLayer that works with ChatGPT, Claude API, and other cloud-based AI agents.

## Features

- **SSE Transport** - Server-Sent Events for remote connections
- **OpenAI Compatible** - Implements `search` and `fetch` tools for ChatGPT deep research
- **URL-Based** - Upload presentations via URL, no local file access needed
- **Cloud Ready** - Deploy to Railway, Render, or any Docker-compatible platform

## Architecture

```
ChatGPT/Claude → HTTPS/SSE → Remote MCP Server → DocLayer (Python) → .NET Runtime
```

## Tools

### Required for ChatGPT Deep Research

- **`search(query)`** - Search through presentations for relevant slides
- **`fetch(id)`** - Retrieve full slide content by ID

### Additional DocLayer Tools

- **`upload_presentation(file_url, presentation_id)`** - Upload a .pptx from URL
- **`create_presentation(presentation_id, title, ...)`** - Create new presentations
- **`render_slide(presentation_id, slide_number)`** - Render slide as base64 image

## Local Development

### Prerequisites

- Python 3.11+
- .NET 8.0 Runtime
- doclayer-py package

### Install

```bash
pip install -r requirements.txt
pip install ../../python-wrapper
```

### Run

```bash
python server.py
```

Server runs on `http://localhost:8000/sse`

## Deploy to Render

### Option 1: Via Dashboard

1. Create new Web Service on Render
2. Connect your GitHub repo
3. Set build command: `pip install -r src/doclayer-remote-mcp/requirements.txt && pip install python-wrapper`
4. Set start command: `python src/doclayer-remote-mcp/server.py`
5. Add environment variable:
   - `DOTNET_INSTALL_DIR=/opt/render/.dotnet`
6. Deploy

### Option 2: Using Docker

**Important**: Prepare the build context first:

```bash
cd src/doclayer-remote-mcp
# Copy python wrapper for Docker build
bash prepare-build.sh

# Build and run
docker build -t doclayer-mcp .
docker run -p 8000:8000 doclayer-mcp
```

**For Render**: Run `bash prepare-build.sh` locally, commit the `doclayer_python_local` folder, then deploy.

## Deploy to Railway

1. Install Railway CLI: `npm i -g railway`
2. Login: `railway login`
3. Create project: `railway init`
4. Deploy: `railway up`

Railway will auto-detect the Dockerfile and deploy.

## Connect to ChatGPT

Once deployed, you'll get a URL like: `https://your-app.onrender.com`

### Configure in ChatGPT Prompts

1. Go to [OpenAI Prompts](https://platform.openai.com/chat)
2. Create or edit a prompt
3. Add MCP tool:

```json
{
  "type": "mcp",
  "server_label": "doclayer",
  "server_url": "https://your-app.onrender.com/sse",
  "allowed_tools": ["search", "fetch", "upload_presentation", "create_presentation"],
  "require_approval": "never"
}
```

### Test via API

```bash
curl https://api.openai.com/v1/responses \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer $OPENAI_API_KEY" \
  -d '{
  "model": "o4-mini-deep-research",
  "input": [
    {
      "role": "user",
      "content": [{
        "type": "input_text",
        "text": "Create a presentation about AI trends"
      }]
    }
  ],
  "tools": [
    {
      "type": "mcp",
      "server_label": "doclayer",
      "server_url": "https://your-app.onrender.com/sse",
      "allowed_tools": ["search", "fetch", "create_presentation"],
      "require_approval": "never"
    }
  ]
}'
```

## Usage Example

1. **Upload a presentation:**
```
User: "Upload this presentation: https://example.com/sales.pptx with ID 'sales-2024'"
Agent: *calls upload_presentation*
```

2. **Search for content:**
```
User: "Find slides about Q4 revenue"
Agent: *calls search with query "Q4 revenue"*
```

3. **Get full content:**
```
User: "Show me the details of slide 3"
Agent: *calls fetch with id from search results*
```

4. **Create new presentation:**
```
User: "Create a presentation titled 'Annual Report 2024'"
Agent: *calls create_presentation*
```

## Storage

Currently uses in-memory storage for uploaded presentations. For production:

- Use Redis for shared state across instances
- Use S3/R2 for presentation file storage
- Add authentication/API keys

## Security

- Add API key authentication via headers
- Rate limiting
- Input validation
- HTTPS only in production

## Troubleshooting

### .NET Runtime Not Found

Ensure .NET 8.0 runtime is installed:
```bash
dotnet --list-runtimes
```

Should show: `Microsoft.NETCore.App 8.0.x`

### Port Already in Use

Change port in `server.py`:
```python
server.run(transport="sse", host="0.0.0.0", port=8080)
```

### Import Errors

Make sure doclayer-py is installed:
```bash
pip install ../../python-wrapper
```

## Links

- [FastMCP Documentation](https://github.com/jlowin/fastmcp)
- [OpenAI MCP Guide](https://platform.openai.com/docs/guides/mcp)
- [Model Context Protocol](https://modelcontextprotocol.io/)
