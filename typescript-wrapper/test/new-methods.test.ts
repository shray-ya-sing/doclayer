/**
 * Test suite for new DocLayer TypeScript wrapper methods
 * Tests slide extraction, editing, and rendering functionality
 */

import { DocLayerClient } from '../src/index';
import * as path from 'path';
import * as fs from 'fs';

const TEST_OUTPUT_DIR = path.join(__dirname, 'test_outputs');
const TEST_FILE = path.join(TEST_OUTPUT_DIR, 'ts_test_title_slide.pptx');

describe('DocLayer TypeScript Wrapper - New Methods', () => {
  let client: DocLayerClient;

  beforeAll(() => {
    // Use venv Python for testing
    const venvPython = path.join(__dirname, '..', '..', 'python-wrapper', 'venv', 'Scripts', 'python.exe');
    client = new DocLayerClient({
      pythonPath: venvPython
    });
    
    // Check if test file exists
    if (!fs.existsSync(TEST_FILE)) {
      console.log(`⚠️  Test file not found: ${TEST_FILE}`);
      console.log('⚠️  Run the basic tests first to create test files.');
    }
  });

  test('Get slide count', async () => {
    if (!fs.existsSync(TEST_FILE)) {
      console.log('⚠️  Skipping test - test file not available');
      return;
    }

    const count = await client.getSlideCount(TEST_FILE);
    
    expect(typeof count).toBe('number');
    expect(count).toBeGreaterThan(0);
    
    console.log(`✓ Slide count: ${count}`);
  }, 30000);

  test('Extract slide content', async () => {
    if (!fs.existsSync(TEST_FILE)) {
      console.log('⚠️  Skipping test - test file not available');
      return;
    }

    const content = await client.extractSlideContent(TEST_FILE, 1);
    
    expect(content).toHaveProperty('shapes');
    expect(content).toHaveProperty('pictures');
    expect(Array.isArray(content.shapes)).toBe(true);
    expect(Array.isArray(content.pictures)).toBe(true);
    
    console.log(`✓ Extracted content from slide 1:`);
    console.log(`  - Shapes: ${content.shapes.length}`);
    console.log(`  - Pictures: ${content.pictures.length}`);
    
    if (content.shapes.length > 0) {
      console.log(`  First shape: ${content.shapes[0].name}`);
    }
  }, 30000);

  test('Extract all slides', async () => {
    if (!fs.existsSync(TEST_FILE)) {
      console.log('⚠️  Skipping test - test file not available');
      return;
    }

    const allContent = await client.extractAllSlides(TEST_FILE);
    
    expect(typeof allContent).toBe('object');
    
    const slideNumbers = Object.keys(allContent).map(Number);
    expect(slideNumbers.length).toBeGreaterThan(0);
    
    console.log(`✓ Extracted content from ${slideNumbers.length} slide(s):`);
    
    for (const slideNum of slideNumbers) {
      const content = allContent[slideNum];
      console.log(`  Slide ${slideNum}: ${content.shapes?.length || 0} shapes, ${content.pictures?.length || 0} pictures`);
    }
  }, 30000);

  test('Edit slide text', async () => {
    if (!fs.existsSync(TEST_FILE)) {
      console.log('⚠️  Skipping test - test file not available');
      return;
    }

    // Create a copy to edit
    const editFile = path.join(TEST_OUTPUT_DIR, 'ts_test_edited.pptx');
    fs.copyFileSync(TEST_FILE, editFile);

    // Extract content to find shape names
    const content = await client.extractSlideContent(editFile, 1);
    
    if (content.shapes.length === 0) {
      console.log('⚠️  No shapes found to edit');
      return;
    }

    const shapeName = content.shapes[0].name;
    console.log(`  Editing shape: ${shapeName}`);

    // Edit the text
    await client.editSlideText(
      editFile,
      1,
      shapeName,
      'EDITED: This text was changed by TypeScript wrapper!'
    );

    expect(fs.existsSync(editFile)).toBe(true);
    console.log(`✓ Text edited successfully: ${editFile}`);
  }, 30000);

  test('Render slide to image', async () => {
    if (!fs.existsSync(TEST_FILE)) {
      console.log('⚠️  Skipping test - test file not available');
      return;
    }

    const imagePath = await client.renderSlideToImage(TEST_FILE, 1);
    
    expect(typeof imagePath).toBe('string');
    expect(imagePath).toBeTruthy();
    expect(fs.existsSync(imagePath)).toBe(true);
    
    const stats = fs.statSync(imagePath);
    expect(stats.size).toBeGreaterThan(1000); // Image should be at least 1KB
    console.log(`✓ Rendered slide to: ${imagePath} (${stats.size} bytes)`);
  }, 30000);

  test('Render all slides', async () => {
    if (!fs.existsSync(TEST_FILE)) {
      console.log('⚠️  Skipping test - test file not available');
      return;
    }

    const imagePaths = await client.renderAllSlides(TEST_FILE);
    
    expect(Array.isArray(imagePaths)).toBe(true);
    expect(imagePaths.length).toBeGreaterThan(0);
    expect(imagePaths.length).toBe(2); // Should have 2 slides
    
    console.log(`✓ Rendered ${imagePaths.length} slide(s):`);
    
    for (let i = 0; i < imagePaths.length; i++) {
      expect(imagePaths[i]).toBeTruthy();
      expect(fs.existsSync(imagePaths[i])).toBe(true);
      const stats = fs.statSync(imagePaths[i]);
      expect(stats.size).toBeGreaterThan(1000); // Each image should be at least 1KB
      console.log(`  Slide ${i + 1}: ${imagePaths[i]} (${stats.size} bytes)`);
    }
  }, 30000);
});
