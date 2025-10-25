/**
 * DocLayer TypeScript Client Library
 * Provides TypeScript bindings for PowerPoint generation via Python bridge
 */

import { spawn } from 'child_process';
import { promises as fs } from 'fs';
import * as path from 'path';
import * as os from 'os';

export interface TitleSlideOptions {
  title: string;
  subtitle?: string;
  footnote?: string;
}

export interface ThemeOptions {
  fontName?: string;
  accentColors?: [string, string, string, string]; // Must be exactly 4 colors
}

export interface PresentationWithThemeOptions extends TitleSlideOptions {
  theme?: ThemeOptions;
}

export interface Position {
  x: number;
  y: number;
}

export interface Size {
  width: number;
  height: number;
}

export interface ShapeInfo {
  name: string;
  text: string;
  position?: Position;
  size?: Size;
}

export interface PictureInfo {
  name: string;
  position?: Position;
  size?: Size;
}

export interface SlideContent {
  shapes: ShapeInfo[];
  pictures: PictureInfo[];
}

export class DocLayerError extends Error {
  constructor(message: string, public code?: string) {
    super(message);
    this.name = 'DocLayerError';
  }
}

/**
 * Main DocLayer client for TypeScript/JavaScript applications
 * Uses Python bridge to generate PowerPoint files
 */
export class DocLayerClient {
  private pythonPath: string;
  private pythonWrapperPath: string;
  private tempDir: string;

  constructor(options: {
    pythonPath?: string;
    pythonWrapperPath?: string;
    tempDir?: string;
  } = {}) {
    // Default to system Python
    this.pythonPath = options.pythonPath || 'python';
    
    // Default to python-wrapper directory in parent (for source installations)
    this.pythonWrapperPath = options.pythonWrapperPath || 
      path.join(__dirname, '..', '..', 'python-wrapper');
    
    this.tempDir = options.tempDir || os.tmpdir();
  }

  /**
   * Create a presentation with a title slide
   */
  async createTitleSlide(
    filepath: string,
    options: TitleSlideOptions
  ): Promise<Buffer> {
    const script = this._generatePythonScript('create_title_slide', {
      filepath,
      ...options
    });

    return await this._executePythonScript(script, filepath);
  }

  /**
   * Create a presentation with custom theme
   */
  async createPresentationWithTheme(
    filepath: string,
    options: PresentationWithThemeOptions
  ): Promise<Buffer> {
    // Validate accent colors if provided
    if (options.theme?.accentColors) {
      if (options.theme.accentColors.length !== 4) {
        throw new DocLayerError('Must provide exactly 4 accent colors');
      }
    }

    const script = this._generatePythonScript('create_presentation_with_theme', {
      filepath,
      title: options.title,
      subtitle: options.subtitle,
      footnote: options.footnote,
      font_name: options.theme?.fontName,
      accent_colors: options.theme?.accentColors
    });

    return await this._executePythonScript(script, filepath);
  }

  /**
   * Get the number of slides in a presentation
   */
  async getSlideCount(filepath: string): Promise<number> {
    const script = `
import sys
sys.path.insert(0, r"${this.pythonWrapperPath}")
from doclayer_python import DocLayerClient

try:
    client = DocLayerClient()
    count = client.get_slide_count(r"${filepath}")
    print(count)
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
    sys.exit(1)
`;

    return new Promise((resolve, reject) => {
      const python = spawn(this.pythonPath, ['-c', script]);
      let stdout = '';
      let stderr = '';

      python.stdout.on('data', (data) => stdout += data.toString());
      python.stderr.on('data', (data) => stderr += data.toString());

      python.on('close', (code) => {
        if (code === 0) {
          resolve(parseInt(stdout.trim()));
        } else {
          reject(new DocLayerError(`Failed to get slide count: ${stderr}`, 'PYTHON_ERROR'));
        }
      });
    });
  }

  /**
   * Extract content from a specific slide
   */
  async extractSlideContent(filepath: string, slideNumber: number): Promise<SlideContent> {
    const script = `
import sys
import json
sys.path.insert(0, r"${this.pythonWrapperPath}")
from doclayer_python import DocLayerClient

try:
    client = DocLayerClient()
    content = client.extract_slide_content(r"${filepath}", ${slideNumber})
    print(json.dumps(content))
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
    sys.exit(1)
`;

    return new Promise((resolve, reject) => {
      const python = spawn(this.pythonPath, ['-c', script]);
      let stdout = '';
      let stderr = '';

      python.stdout.on('data', (data) => stdout += data.toString());
      python.stderr.on('data', (data) => stderr += data.toString());

      python.on('close', (code) => {
        if (code === 0) {
          try {
            resolve(JSON.parse(stdout.trim()));
          } catch (e) {
            reject(new DocLayerError('Failed to parse slide content', 'PARSE_ERROR'));
          }
        } else {
          reject(new DocLayerError(`Failed to extract slide content: ${stderr}`, 'PYTHON_ERROR'));
        }
      });
    });
  }

  /**
   * Extract content from all slides
   */
  async extractAllSlides(filepath: string): Promise<Record<number, SlideContent>> {
    const script = `
import sys
import json
sys.path.insert(0, r"${this.pythonWrapperPath}")
from doclayer_python import DocLayerClient

try:
    client = DocLayerClient()
    all_content = client.extract_all_slides(r"${filepath}")
    print(json.dumps(all_content))
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
    sys.exit(1)
`;

    return new Promise((resolve, reject) => {
      const python = spawn(this.pythonPath, ['-c', script]);
      let stdout = '';
      let stderr = '';

      python.stdout.on('data', (data) => stdout += data.toString());
      python.stderr.on('data', (data) => stderr += data.toString());

      python.on('close', (code) => {
        if (code === 0) {
          try {
            resolve(JSON.parse(stdout.trim()));
          } catch (e) {
            reject(new DocLayerError('Failed to parse slides content', 'PARSE_ERROR'));
          }
        } else {
          reject(new DocLayerError(`Failed to extract all slides: ${stderr}`, 'PYTHON_ERROR'));
        }
      });
    });
  }

  /**
   * Edit the text of a shape on a slide
   */
  async editSlideText(
    filepath: string,
    slideNumber: number,
    elementName: string,
    newText: string
  ): Promise<void> {
    const script = `
import sys
sys.path.insert(0, r"${this.pythonWrapperPath}")
from doclayer_python import DocLayerClient

try:
    client = DocLayerClient()
    client.edit_slide_text(r"${filepath}", ${slideNumber}, r"""${elementName}""", r"""${newText}""")
    print("SUCCESS")
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
    sys.exit(1)
`;

    return new Promise((resolve, reject) => {
      const python = spawn(this.pythonPath, ['-c', script]);
      let stderr = '';

      python.stderr.on('data', (data) => stderr += data.toString());

      python.on('close', (code) => {
        if (code === 0) {
          resolve();
        } else {
          reject(new DocLayerError(`Failed to edit slide text: ${stderr}`, 'PYTHON_ERROR'));
        }
      });
    });
  }

  /**
   * Render a slide to a JPEG image
   */
  async renderSlideToImage(filepath: string, slideNumber: number): Promise<string> {
    const script = `
import sys
sys.path.insert(0, r"${this.pythonWrapperPath}")
from doclayer_python import DocLayerClient

try:
    client = DocLayerClient()
    image_path = client.render_slide_to_image(r"${filepath}", ${slideNumber})
    print(image_path)
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
    sys.exit(1)
`;

    return new Promise((resolve, reject) => {
      const python = spawn(this.pythonPath, ['-c', script]);
      let stdout = '';
      let stderr = '';

      python.stdout.on('data', (data) => stdout += data.toString());
      python.stderr.on('data', (data) => stderr += data.toString());

      python.on('close', (code) => {
        if (code === 0) {
          resolve(stdout.trim());
        } else {
          reject(new DocLayerError(`Failed to render slide: ${stderr}`, 'PYTHON_ERROR'));
        }
      });
    });
  }

  /**
   * Render all slides to JPEG images
   */
  async renderAllSlides(filepath: string): Promise<string[]> {
    const script = `
import sys
import json
sys.path.insert(0, r"${this.pythonWrapperPath}")
from doclayer_python import DocLayerClient

try:
    client = DocLayerClient()
    image_paths = client.render_all_slides(r"${filepath}")
    print(json.dumps(image_paths))
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
    sys.exit(1)
`;

    return new Promise((resolve, reject) => {
      const python = spawn(this.pythonPath, ['-c', script]);
      let stdout = '';
      let stderr = '';

      python.stdout.on('data', (data) => stdout += data.toString());
      python.stderr.on('data', (data) => stderr += data.toString());

      python.on('close', (code) => {
        if (code === 0) {
          try {
            resolve(JSON.parse(stdout.trim()));
          } catch (e) {
            reject(new DocLayerError('Failed to parse image paths', 'PARSE_ERROR'));
          }
        } else {
          reject(new DocLayerError(`Failed to render slides: ${stderr}`, 'PYTHON_ERROR'));
        }
      });
    });
  }

  /**
   * Generate Python script to call doclayer_python package
   */
  private _generatePythonScript(method: string, params: any): string {
    const paramsJson = JSON.stringify(params);

    return `
import sys
import json
sys.path.insert(0, r"${this.pythonWrapperPath}")

from doclayer_python import ${method}, DocLayerError

try:
    params = json.loads(r'''${paramsJson}''')
    
    # Remove None values
    params = {k: v for k, v in params.items() if v is not None}
    
    # Call the function
    result = ${method}(**params)
    
    print("SUCCESS")
except DocLayerError as e:
    print(f"DOCLAYER_ERROR: {e}", file=sys.stderr)
    sys.exit(1)
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
    sys.exit(2)
`;
  }

  /**
   * Execute Python script and return result
   */
  private async _executePythonScript(
    script: string,
    outputPath: string
  ): Promise<Buffer> {
    return new Promise((resolve, reject) => {
      const python = spawn(this.pythonPath, ['-c', script]);
      
      let stdout = '';
      let stderr = '';

      python.stdout.on('data', (data) => {
        stdout += data.toString();
      });

      python.stderr.on('data', (data) => {
        stderr += data.toString();
      });

      python.on('close', async (code) => {
        if (code === 0 && stdout.includes('SUCCESS')) {
          try {
            // Read the generated file
            const buffer = await fs.readFile(outputPath);
            resolve(buffer);
          } catch (error) {
            reject(new DocLayerError(
              `Failed to read output file: ${error}`,
              'FILE_READ_ERROR'
            ));
          }
        } else {
          const errorMsg = stderr.trim();
          
          if (errorMsg.includes('DOCLAYER_ERROR:')) {
            reject(new DocLayerError(
              errorMsg.replace('DOCLAYER_ERROR:', '').trim(),
              'DOCLAYER_ERROR'
            ));
          } else if (errorMsg.includes('ERROR:')) {
            reject(new DocLayerError(
              errorMsg.replace('ERROR:', '').trim(),
              'PYTHON_ERROR'
            ));
          } else {
            reject(new DocLayerError(
              `Python process exited with code ${code}: ${errorMsg}`,
              'EXECUTION_ERROR'
            ));
          }
        }
      });

      python.on('error', (error) => {
        reject(new DocLayerError(
          `Failed to spawn Python process: ${error.message}`,
          'SPAWN_ERROR'
        ));
      });
    });
  }

  /**
   * Check if Python and doclayer_python are available
   */
  async checkEnvironment(): Promise<{
    pythonAvailable: boolean;
    pythonVersion?: string;
    doclayerAvailable: boolean;
    error?: string;
  }> {
    try {
      const versionCheck = spawn(this.pythonPath, ['--version']);
      
      return new Promise((resolve) => {
        let output = '';
        
        versionCheck.stdout.on('data', (data) => output += data.toString());
        versionCheck.stderr.on('data', (data) => output += data.toString());
        
        versionCheck.on('close', async (code) => {
          if (code !== 0) {
            resolve({
              pythonAvailable: false,
              doclayerAvailable: false,
              error: 'Python not found'
            });
            return;
          }

          const pythonVersion = output.trim();
          
          // Check if doclayer_python is available
          const importCheck = spawn(this.pythonPath, [
            '-c',
            `import sys; sys.path.insert(0, "${this.pythonWrapperPath.replace(/\\/g, '/')}"); import doclayer_python; print("OK")`
          ]);

          let importOutput = '';
          importCheck.stdout.on('data', (data) => importOutput += data.toString());
          
          importCheck.on('close', (importCode) => {
            resolve({
              pythonAvailable: true,
              pythonVersion,
              doclayerAvailable: importCode === 0 && importOutput.includes('OK')
            });
          });
        });
      });
    } catch (error: any) {
      return {
        pythonAvailable: false,
        doclayerAvailable: false,
        error: error.message
      };
    }
  }
}

/**
 * Convenience functions for quick usage
 */

/**
 * Create a presentation with a title slide
 */
export async function createTitleSlide(
  filepath: string,
  title: string,
  subtitle?: string,
  footnote?: string
): Promise<Buffer> {
  const client = new DocLayerClient();
  return await client.createTitleSlide(filepath, { title, subtitle, footnote });
}

/**
 * Create a presentation with custom theme
 */
export async function createPresentationWithTheme(
  filepath: string,
  title: string,
  options?: {
    subtitle?: string;
    footnote?: string;
    fontName?: string;
    accentColors?: [string, string, string, string];
  }
): Promise<Buffer> {
  const client = new DocLayerClient();
  return await client.createPresentationWithTheme(filepath, {
    title,
    subtitle: options?.subtitle,
    footnote: options?.footnote,
    theme: {
      fontName: options?.fontName,
      accentColors: options?.accentColors
    }
  });
}

// Export main class
export default DocLayerClient;
