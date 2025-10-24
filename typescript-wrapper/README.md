# DocLayer TypeScript/JavaScript Wrapper

TypeScript/JavaScript bindings for the DocLayer PowerPoint generation library. This package uses the Python wrapper as a bridge to generate PowerPoint files from Node.js.

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
- `createTitleSlide(filepath, options)` - Create title slide presentation
- `createPresentationWithTheme(filepath, options)` - Create themed presentation
- `checkEnvironment()` - Check if Python and DocLayer are available

## Testing

```bash
npm install
npm run build
npm test
```

## License

MIT License
