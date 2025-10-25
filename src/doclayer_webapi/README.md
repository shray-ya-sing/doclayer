# DocLayer Web API

REST API for PowerPoint presentation generation, extraction, editing, and rendering. Built with ASP.NET Core 8.0 and DocLayer.Core.

## Features

- **Create**: Generate presentations with title slides and custom themes
- **Upload/Download**: Process existing presentations via REST API
- **Extract**: Get slide count and extract content (shapes, text, pictures)
- **Edit**: Modify text content in presentations
- **Render**: Convert slides to JPEG images (single or batch)
- **Swagger**: Interactive API documentation at `/swagger`

## Getting Started

### Prerequisites

- .NET 8.0 SDK
- Windows, macOS, or Linux

### Run the API

```bash
cd src/doclayer_webapi/doclayer_webapi
dotnet run
```

The API will start at:
- HTTP: `http://localhost:5000`
- HTTPS: `https://localhost:5001`
- Swagger UI: `https://localhost:5001/swagger`

### Build for Production

```bash
dotnet publish -c Release -o ./publish
```

## API Endpoints

### Creation Endpoints

#### Create Title Slide
```http
POST /api/presentation/create-title-slide
Content-Type: application/json

{
  "title": "My Presentation",
  "subtitle": "Created via API",
  "footnote": "Source: DocLayer API"
}
```

**Response:** PowerPoint file download

#### Create with Custom Theme
```http
POST /api/presentation/create-with-theme
Content-Type: application/json

{
  "title": "Themed Presentation",
  "subtitle": "With custom styling",
  "footnote": "Source: DocLayer",
  "fontName": "Arial",
  "accentColors": ["FF5733", "33FF57", "3357FF", "F3FF33"]
}
```

**Response:** PowerPoint file download

### Upload/Download Endpoints

#### Upload Presentation
```http
POST /api/presentation/upload
Content-Type: multipart/form-data

file: [presentation.pptx]
```

**Response:**
```json
{
  "fileId": "tmp1A2B3C.pptx",
  "originalName": "presentation.pptx"
}
```

#### Download Presentation
```http
GET /api/presentation/{fileId}/download
```

**Response:** PowerPoint file download

### Query Endpoints

#### Get Slide Count
```http
GET /api/presentation/{fileId}/slide-count
```

**Response:**
```json
{
  "count": 5
}
```

#### Extract Slide Content
```http
GET /api/presentation/{fileId}/slides/{slideNumber}
```

**Response:**
```json
{
  "shapes": [
    {
      "name": "Title 1",
      "text": "Slide Title",
      "position": { "x": 914400, "y": 914400 },
      "size": { "width": 10058400, "height": 914400 }
    }
  ],
  "pictures": []
}
```

#### Extract All Slides
```http
GET /api/presentation/{fileId}/slides
```

**Response:**
```json
{
  "1": {
    "shapes": [...],
    "pictures": [...]
  },
  "2": {
    "shapes": [...],
    "pictures": [...]
  }
}
```

### Edit Endpoints

#### Edit Slide Text
```http
PUT /api/presentation/{fileId}/edit-text
Content-Type: application/json

{
  "slideNumber": 1,
  "elementName": "Title 1",
  "newText": "Updated Title"
}
```

**Response:**
```json
{
  "success": true,
  "message": "Text updated successfully"
}
```

### Rendering Endpoints

#### Render Single Slide
```http
GET /api/presentation/{fileId}/slides/{slideNumber}/render
```

**Response:** JPEG image file

#### Render All Slides
```http
GET /api/presentation/{fileId}/slides/render-all
```

**Response:** ZIP file containing all slide images

### Management Endpoints

#### Delete Presentation
```http
DELETE /api/presentation/{fileId}
```

**Response:**
```json
{
  "success": true,
  "message": "File deleted successfully"
}
```

## Usage Examples

### cURL Examples

**Create Title Slide:**
```bash
curl -X POST "https://localhost:5001/api/presentation/create-title-slide" \
  -H "Content-Type: application/json" \
  -d '{"title":"My Presentation","subtitle":"Created via API"}' \
  -o presentation.pptx
```

**Upload and Extract:**
```bash
# Upload
FILE_ID=$(curl -X POST "https://localhost:5001/api/presentation/upload" \
  -F "file=@presentation.pptx" | jq -r '.fileId')

# Get slide count
curl "https://localhost:5001/api/presentation/$FILE_ID/slide-count"

# Extract slide 1
curl "https://localhost:5001/api/presentation/$FILE_ID/slides/1"

# Render slide 1
curl "https://localhost:5001/api/presentation/$FILE_ID/slides/1/render" \
  -o slide1.jpg
```

### JavaScript/TypeScript Example

```typescript
// Create presentation
const response = await fetch('https://localhost:5001/api/presentation/create-title-slide', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    title: 'My Presentation',
    subtitle: 'Created via API'
  })
});
const blob = await response.blob();
// Download or process blob

// Upload and process
const formData = new FormData();
formData.append('file', pptxFile);

const uploadRes = await fetch('https://localhost:5001/api/presentation/upload', {
  method: 'POST',
  body: formData
});
const { fileId } = await uploadRes.json();

// Extract content
const contentRes = await fetch(`https://localhost:5001/api/presentation/${fileId}/slides/1`);
const content = await contentRes.json();
console.log(content.shapes);

// Edit text
await fetch(`https://localhost:5001/api/presentation/${fileId}/edit-text`, {
  method: 'PUT',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    slideNumber: 1,
    elementName: 'Title 1',
    newText: 'Updated Title'
  })
});

// Download modified file
const downloadRes = await fetch(`https://localhost:5001/api/presentation/${fileId}/download`);
const modifiedBlob = await downloadRes.blob();
```

### Python Example

```python
import requests

# Create presentation
response = requests.post(
    'https://localhost:5001/api/presentation/create-title-slide',
    json={
        'title': 'My Presentation',
        'subtitle': 'Created via API'
    },
    verify=False  # Only for development
)
with open('presentation.pptx', 'wb') as f:
    f.write(response.content)

# Upload and process
with open('presentation.pptx', 'rb') as f:
    upload_response = requests.post(
        'https://localhost:5001/api/presentation/upload',
        files={'file': f},
        verify=False
    )
file_id = upload_response.json()['fileId']

# Extract content
content = requests.get(
    f'https://localhost:5001/api/presentation/{file_id}/slides/1',
    verify=False
).json()
print(f"Shapes: {len(content['shapes'])}")

# Edit text
requests.put(
    f'https://localhost:5001/api/presentation/{file_id}/edit-text',
    json={
        'slideNumber': 1,
        'elementName': 'Title 1',
        'newText': 'Updated Title'
    },
    verify=False
)

# Download modified
response = requests.get(
    f'https://localhost:5001/api/presentation/{file_id}/download',
    verify=False
)
with open('modified.pptx', 'wb') as f:
    f.write(response.content)
```

## Configuration

### appsettings.json

```json
{
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  },
  "AllowedHosts": "*"
}
```

### CORS Configuration (Optional)

To allow cross-origin requests, add to `Program.cs`:

```csharp
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

// ...

app.UseCors();
```

## Architecture

```
┌─────────────────┐
│   REST Client   │
│ (HTTP Requests) │
└────────┬────────┘
         │
         v
┌─────────────────────────┐
│  PresentationController │
│   (ASP.NET Core API)    │
└────────┬────────────────┘
         │
         v
┌─────────────────────────┐
│   DocLayer.Core         │
│   PresentationBuilder   │
└────────┬────────────────┘
         │
         v
┌─────────────────────────┐
│   OpenXML SDK           │
│   Syncfusion Rendering  │
└─────────────────────────┘
         │
         v
┌─────────────────────────┐
│   PowerPoint Files      │
│   JPEG Images           │
└─────────────────────────┘
```

## Error Handling

The API returns standard HTTP status codes:

- **200 OK**: Successful operation
- **400 Bad Request**: Invalid input or operation failed
- **404 Not Found**: File not found
- **500 Internal Server Error**: Server error

Error responses include a descriptive message:
```json
{
  "error": "Must provide exactly 4 accent colors"
}
```

## File Storage

Uploaded files are stored temporarily in the system temp directory. Files should be deleted via the DELETE endpoint when no longer needed to prevent disk space issues.

**Best Practice:** Implement a background service to clean up files older than 24 hours.

## Deployment

### Docker (Optional)

```dockerfile
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
WORKDIR /app
EXPOSE 80
EXPOSE 443

FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src
COPY ["doclayer_webapi.csproj", "./"]
RUN dotnet restore
COPY . .
RUN dotnet build -c Release -o /app/build

FROM build AS publish
RUN dotnet publish -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "doclayer_webapi.dll"]
```

### Azure App Service

```bash
# Publish and deploy
dotnet publish -c Release
cd bin/Release/net8.0/publish
zip -r deploy.zip .
az webapp deployment source config-zip --resource-group <group> --name <app-name> --src deploy.zip
```

## Performance Considerations

- **File Size**: Large presentations may take longer to process
- **Rendering**: Slide rendering is CPU-intensive; consider rate limiting
- **Storage**: Temp files should be cleaned up regularly
- **Concurrency**: The API is stateless and can scale horizontally

## Security Considerations

- **HTTPS**: Always use HTTPS in production
- **Authentication**: Add authentication/authorization as needed
- **File Validation**: The API validates file types (.pptx only)
- **Input Validation**: All inputs are validated
- **Rate Limiting**: Consider implementing rate limiting for rendering operations

## License

MIT License

## Support

For issues and questions, please refer to the main DocLayer repository.
