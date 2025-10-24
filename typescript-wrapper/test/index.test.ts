/**
 * Test suite for DocLayer TypeScript wrapper
 */

import { DocLayerClient, createTitleSlide, createPresentationWithTheme } from '../src/index';
import * as path from 'path';
import * as fs from 'fs';

const TEST_OUTPUT_DIR = path.join(__dirname, 'test_outputs');

// Ensure test output directory exists
if (!fs.existsSync(TEST_OUTPUT_DIR)) {
  fs.mkdirSync(TEST_OUTPUT_DIR, { recursive: true });
}

describe('DocLayer TypeScript Wrapper', () => {
  let client: DocLayerClient;
  let envCheck: any;

  beforeAll(async () => {
    client = new DocLayerClient();
    envCheck = await client.checkEnvironment();
  });

  test('Environment check - Python available', () => {
    expect(envCheck.pythonAvailable).toBe(true);
  });

  test('Environment check - DocLayer Python available', () => {
    expect(envCheck.doclayerAvailable).toBe(true);
    if (!envCheck.doclayerAvailable) {
      console.log('⚠️  DocLayer Python package not available. Install it to run remaining tests.');
    }
  });

  test('Create title slide presentation', async () => {
    if (!envCheck.doclayerAvailable) {
      console.log('⚠️  Skipping test - DocLayer not available');
      return;
    }

    const outputPath = path.join(TEST_OUTPUT_DIR, 'ts_test_title_slide.pptx');
    
    const buffer = await createTitleSlide(
      outputPath,
      'Welcome to DocLayer TypeScript',
      'PowerPoint Generation from Node.js via Python',
      'Source: DocLayer TypeScript Wrapper'
    );

    expect(buffer).toBeInstanceOf(Buffer);
    expect(buffer.length).toBeGreaterThan(0);
    expect(fs.existsSync(outputPath)).toBe(true);
    
    console.log(`✓ Created presentation: ${outputPath} (${buffer.length} bytes)`);
  }, 30000);

  test('Create presentation with custom theme', async () => {
    if (!envCheck.doclayerAvailable) {
      console.log('⚠️  Skipping test - DocLayer not available');
      return;
    }

    const outputPath = path.join(TEST_OUTPUT_DIR, 'ts_test_theme.pptx');
    
    const buffer = await createPresentationWithTheme(
      outputPath,
      'Custom Theme Test',
      {
        subtitle: 'Arial font with custom colors from TypeScript',
        footnote: 'Source: DocLayer TypeScript Theme Test',
        fontName: 'Arial',
        accentColors: ['FF5733', '33FF57', '3357FF', 'F3FF33']
      }
    );

    expect(buffer).toBeInstanceOf(Buffer);
    expect(buffer.length).toBeGreaterThan(0);
    expect(fs.existsSync(outputPath)).toBe(true);
    
    console.log(`✓ Created themed presentation: ${outputPath} (${buffer.length} bytes)`);
  }, 30000);
});
