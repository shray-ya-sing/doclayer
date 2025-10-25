# DocLayer TypeScript/JavaScript Wrapper

TypeScript/JavaScript bindings for the DocLayer PowerPoint library. Generate, extract, edit, and render PowerPoint presentations from Node.js. This package uses the Python wrapper as a bridge to access the full DocLayer.Core functionality.

## Installation

```bash
npm install @doclayer/ts
```

## Requirements

- Node.js >= 16.0.0
- Python 3.8+ with `pythonnet` package installed
- `doclayer-py` Python package

## Usage

### Basic Title Slide

```typescript
import { createTitleSlide } from '@doclayer/ts';

// Create a presentation with a title slide
const buffer = await createTitleSlide(
  'presentation.pptx',
  'Welcome to DocLayer',
  'PowerPoint Generation Made Easy',
  'Source: DocLayer'
);

console.log(`Created presentation: ${buffer.length} bytes`);
```

### Presentation with Custom Theme

```typescript
import { createPresentationWithTheme } from '@doclayer/ts';

// Create a presentation with custom font and colors
const buffer = await createPresentationWithTheme(
  'custom_theme.pptx',
  'Custom Theme Demo',
  {
    subtitle: 'Arial font with brand colors',
    footnote: 'Source: My Company',
    fontName: 'Arial',
    accentColors: ['FF5733', '33FF57', '3357FF', 'F3FF33'] // Must be exactly 4 colors
  }
);

console.log(`Created themed presentation: ${buffer.length} bytes`);
```

### Using the Client

```typescript
import { DocLayerClient } from '@doclayer/ts';

const client = new DocLayerClient({
  pythonPath: 'python', // Optional: path to Python executable
  pythonWrapperPath: '../python-wrapper' // Optional: path to doclayer_python package
});

// Check environment
const env = await client.checkEnvironment();
console.log('Python available:', env.pythonAvailable);
console.log('DocLayer available:', env.doclayerAvailable);

// Create presentations
const titleSlideBuffer = await client.createTitleSlide('output.pptx', {
  title: 'My Presentation',
  subtitle: 'Created with TypeScript',
  footnote: 'Source: My Data'
});

const themedBuffer = await client.createPresentationWithTheme('themed.pptx', {
  title: 'Themed Presentation',
  subtitle: 'With custom styling',
  theme: {
    fontName: 'Calibri',
    accentColors: ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000']
  }
});
```

### Extract and Edit Presentations

```typescript
import { DocLayerClient } from '@doclayer/ts';

const client = new DocLayerClient();

// Get slide count
const count = await client.getSlideCount('presentation.pptx');
console.log(`Presentation has ${count} slides`);

// Extract content from a specific slide
const content = await client.extractSlideContent('presentation.pptx', 1);
content.shapes.forEach(shape => {
  console.log(`Shape: ${shape.name}`);
  console.log(`Text: ${shape.text}`);
  console.log(`Position:`, shape.position);
  console.log(`Size:`, shape.size);
});

// Extract all slides
const allContent = await client.extractAllSlides('presentation.pptx');
for (const [slideNum, slideContent] of Object.entries(allContent)) {
  console.log(`Slide ${slideNum}: ${slideContent.shapes.length} shapes, ${slideContent.pictures.length} pictures`);
}

// Edit text on a slide
await client.editSlideText(
  'presentation.pptx',
  1,
  'Title 1',
  'Updated Title Text'
);

// Render slide to JPEG image
const imagePath = await client.renderSlideToImage('presentation.pptx', 1);
console.log(`Slide rendered to: ${imagePath}`);

// Render all slides to images
const imagePaths = await client.renderAllSlides('presentation.pptx');
imagePaths.forEach((path, i) => {
  console.log(`Slide ${i + 1} rendered to: ${path}`);
});
```

## Architecture

The TypeScript wrapper uses a Python bridge architecture:

```
TypeScript/Node.js → child_process → Python → doclayer_python → DocLayer.Core (C#) → PowerPoint Files
```

This approach provides:
- ✅ Cross-platform compatibility
- ✅ Full feature parity with Python wrapper  
- ✅ Easy to maintain (reuses Python implementation)
- ✅ No complex .NET to Node.js interop

## API Reference

### createTitleSlide(filepath, title, subtitle?, footnote?)

Create a presentation with a title slide.

**Parameters:**
- `filepath` (string): Output file path
- `title` (string): Main title text
- `subtitle` (string, optional): Subtitle text
- `footnote` (string, optional): Footnote text

**Returns:** `Promise<Buffer>` - The generated presentation file as a buffer

### createPresentationWithTheme(filepath, title, options?)

Create a presentation with custom theme.

**Parameters:**
- `filepath` (string): Output file path
- `title` (string): Main title text
- `options` (object, optional):
  - `subtitle` (string): Subtitle text
  - `footnote` (string): Footnote text
  - `fontName` (string): Font typeface name (e.g., "Arial", "Calibri")
  - `accentColors` (array): Array of exactly 4 hex color codes

**Returns:** `Promise<Buffer>` - The generated presentation file as a buffer

### DocLayerClient

Main client class for DocLayer operations.

**Constructor options:**
- `pythonPath` (string): Path to Python executable (default: "python")
- `pythonWrapperPath` (string): Path to doclayer_python package (for source installations)
- `tempDir` (string): Temporary directory for intermediate files

**Methods:**

*Creation:*
- `createTitleSlide(filepath, options)` - Create title slide presentation
- `createPresentationWithTheme(filepath, options)` - Create themed presentation

*Query:*
- `getSlideCount(filepath)` - Get number of slides in presentation
- `extractSlideContent(filepath, slideNumber)` - Extract shapes and pictures from a slide
- `extractAllSlides(filepath)` - Extract content from all slides

*Edit:*
- `editSlideText(filepath, slideNumber, elementName, newText)` - Modify text of a shape

*Rendering:*
- `renderSlideToImage(filepath, slideNumber)` - Render slide to JPEG, returns image path
- `renderAllSlides(filepath)` - Render all slides to JPEG images, returns array of paths

*Utility:*
- `checkEnvironment()` - Check if Python and DocLayer are available

## API Types

```typescript
interface Position {
  x: number;
  y: number;
}

interface Size {
  width: number;
  height: number;
}

interface ShapeInfo {
  name: string;
  text: string;
  position?: Position;
  size?: Size;
}

interface PictureInfo {
  name: string;
  position?: Position;
  size?: Size;
}

interface SlideContent {
  shapes: ShapeInfo[];
  pictures: PictureInfo[];
}
```

## Testing

```bash
npm install
npm run build
npm test
```

## License

MIT License
