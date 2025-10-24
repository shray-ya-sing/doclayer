/**
 * DocLayer TypeScript Client Library
 * Provides TypeScript bindings for PowerPoint generation via REST bridge
 */

export interface SlideData {
  title?: string;
  content?: string;
  shapes?: ShapeData[];
  table?: TableData;
  footnote?: string;
  pageNumber?: string;
}

export interface ShapeData {
  type: 'rectangle' | 'circle' | 'triangle' | 'textbox' | 'arrow';
  x: number;
  y: number;
  width: number;
  height: number;
  text?: string;
  style?: ShapeStyle;
}

export interface ShapeStyle {
  fillColor?: string;
  borderColor?: string;
  borderWidth?: number;
  fontSize?: number;
}

export interface TableData {
  rows: number;
  cols: number;
  data?: string[][];
  style?: TableStyle;
}

export interface TableStyle {
  headerStyle?: ShapeStyle;
  cellStyle?: ShapeStyle;
}

export interface PresentationOptions {
  format: '4:3' | '16:9';
  theme?: string;
}

export class DocLayerError extends Error {
  constructor(message: string, public code?: string) {
    super(message);
    this.name = 'DocLayerError';
  }
}

/**
 * Main DocLayer client for TypeScript/JavaScript applications
 */
export class DocLayerClient {
  private bridgeUrl: string;
  private tempDir: string;

  constructor(options: {
    bridgeUrl?: string;
    tempDir?: string;
  } = {}) {
    // For Node.js environments, we'll use a local bridge server
    // For browser environments, this would be a WebAssembly module
    this.bridgeUrl = options.bridgeUrl || 'http://localhost:5000';
    this.tempDir = options.tempDir || './temp';
  }

  /**
   * Create a new presentation builder
   */
  createPresentation(options: PresentationOptions = { format: '16:9' }): PresentationBuilder {
    return new PresentationBuilder(this, options);
  }

  /**
   * Quick method to create a simple presentation
   */
  async createBasicPresentation(
    title: string,
    slides: SlideData[],
    options: PresentationOptions = { format: '16:9' }
  ): Promise<Uint8Array> {
    const builder = this.createPresentation(options);
    
    slides.forEach((slideData, index) => {
      const slide = builder.addSlide();
      
      if (slideData.title) slide.addTitle(slideData.title);
      if (slideData.content) slide.addTextbox(slideData.content);
      if (slideData.table) slide.addTable(slideData.table.rows, slideData.table.cols);
      if (slideData.shapes) {
        slideData.shapes.forEach(shape => slide.addShape(shape));
      }
      if (slideData.footnote) slide.addFootnote(slideData.footnote);
      if (slideData.pageNumber) slide.addPageNumber(slideData.pageNumber);
    });

    return await builder.build();
  }

  /**
   * Internal method to communicate with C# bridge
   */
  async _callBridge(method: string, payload: any): Promise<any> {
    try {
      // In a real implementation, this would either:
      // 1. Call a local REST bridge server running your C# code
      // 2. Use WebAssembly compilation of your C# code
      // 3. Use Node.js addon for direct .NET interop
      
      if (typeof window !== 'undefined') {
        // Browser environment - use WebAssembly or REST API
        return await this._callRestBridge(method, payload);
      } else {
        // Node.js environment - use local bridge server
        return await this._callLocalBridge(method, payload);
      }
    } catch (error) {
      throw new DocLayerError(`Bridge communication failed: ${error}`);
    }
  }

  private async _callRestBridge(method: string, payload: any): Promise<any> {
    const response = await fetch(`${this.bridgeUrl}/${method}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    return response.arrayBuffer();
  }

  private async _callLocalBridge(method: string, payload: any): Promise<any> {
    // For Node.js, you could use child_process to call a local bridge
    const { spawn } = await import('child_process');
    
    return new Promise((resolve, reject) => {
      const bridge = spawn('dotnet', ['run', '--project', './bridge', '--', method]);
      
      let output = '';
      bridge.stdout.on('data', (data) => output += data);
      bridge.stderr.on('data', (data) => reject(new Error(data.toString())));
      bridge.on('close', (code) => {
        if (code === 0) resolve(JSON.parse(output));
        else reject(new Error(`Bridge exited with code ${code}`));
      });

      bridge.stdin.write(JSON.stringify(payload));
      bridge.stdin.end();
    });
  }
}

/**
 * Fluent builder for PowerPoint presentations
 */
export class PresentationBuilder {
  private slides: SlideBuilder[] = [];

  constructor(
    private client: DocLayerClient,
    private options: PresentationOptions
  ) {}

  /**
   * Add a new slide to the presentation
   */
  addSlide(): SlideBuilder {
    const slide = new SlideBuilder(this.client);
    this.slides.push(slide);
    return slide;
  }

  /**
   * Set presentation to widescreen format
   */
  setWidescreen(): PresentationBuilder {
    this.options.format = '16:9';
    return this;
  }

  /**
   * Set presentation to standard format
   */
  setStandard(): PresentationBuilder {
    this.options.format = '4:3';
    return this;
  }

  /**
   * Build the final presentation
   */
  async build(): Promise<Uint8Array> {
    const payload = {
      options: this.options,
      slides: this.slides.map(slide => slide.getData())
    };

    const result = await this.client._callBridge('build-presentation', payload);
    return new Uint8Array(result);
  }
}

/**
 * Builder for individual slides
 */
export class SlideBuilder {
  private data: SlideData = {};

  constructor(private client: DocLayerClient) {}

  /**
   * Add title to the slide
   */
  addTitle(text: string): SlideBuilder {
    this.data.title = text;
    return this;
  }

  /**
   * Add text box to the slide
   */
  addTextbox(text: string): SlideBuilder {
    this.data.content = text;
    return this;
  }

  /**
   * Add table to the slide
   */
  addTable(rows: number, cols: number, data?: string[][]): SlideBuilder {
    this.data.table = { rows, cols, data };
    return this;
  }

  /**
   * Add shape to the slide
   */
  addShape(shape: ShapeData): SlideBuilder {
    if (!this.data.shapes) this.data.shapes = [];
    this.data.shapes.push(shape);
    return this;
  }

  /**
   * Add rectangle shape
   */
  addRectangle(x: number, y: number, width: number, height: number): SlideBuilder {
    return this.addShape({ type: 'rectangle', x, y, width, height });
  }

  /**
   * Add circle shape
   */
  addCircle(x: number, y: number, width: number, height: number): SlideBuilder {
    return this.addShape({ type: 'circle', x, y, width, height });
  }

  /**
   * Add triangle shape
   */
  addTriangle(x: number, y: number, width: number, height: number): SlideBuilder {
    return this.addShape({ type: 'triangle', x, y, width, height });
  }

  /**
   * Add positioned text box
   */
  addPositionedTextbox(text: string, x: number, y: number, width: number = 2, height: number = 1): SlideBuilder {
    return this.addShape({ type: 'textbox', text, x, y, width, height });
  }

  /**
   * Add footnote to the slide
   */
  addFootnote(text: string = "Source:"): SlideBuilder {
    this.data.footnote = text;
    return this;
  }

  /**
   * Add page number to the slide
   */
  addPageNumber(pageNum: string): SlideBuilder {
    this.data.pageNumber = pageNum;
    return this;
  }

  /**
   * Get slide data (internal)
   */
  getData(): SlideData {
    return { ...this.data };
  }
}

/**
 * Utility functions for common presentation patterns
 */
export class DocLayerUtils {
  /**
   * Create a dashboard-style presentation
   */
  static async createDashboard(
    client: DocLayerClient,
    title: string,
    metrics: Array<{ label: string; value: string; trend?: 'up' | 'down' | 'stable' }>
  ): Promise<Uint8Array> {
    const builder = client.createPresentation({ format: '16:9' });
    const slide = builder.addSlide();

    slide.addTitle(title);
    
    // Add metric boxes
    metrics.forEach((metric, index) => {
      const x = 1 + (index % 3) * 3;
      const y = 2 + Math.floor(index / 3) * 2;
      
      slide.addRectangle(x, y, 2, 1);
      slide.addPositionedTextbox(`${metric.label}\n${metric.value}`, x, y, 2, 1);
    });

    slide.addFootnote("Dashboard generated by DocLayer");
    slide.addPageNumber("1");

    return await builder.build();
  }

  /**
   * Create a process flow presentation
   */
  static async createProcessFlow(
    client: DocLayerClient,
    title: string,
    steps: string[]
  ): Promise<Uint8Array> {
    const builder = client.createPresentation({ format: '16:9' });
    const slide = builder.addSlide();

    slide.addTitle(title);

    // Add process steps
    steps.forEach((step, index) => {
      const x = 1 + index * 2;
      const y = 3;
      
      slide.addRectangle(x, y, 1.5, 1);
      slide.addPositionedTextbox(step, x, y - 0.5, 1.5, 0.5);
      
      // Add arrow (except for last step)
      if (index < steps.length - 1) {
        slide.addShape({ type: 'arrow', x: x + 1.5, y: y + 0.25, width: 0.5, height: 0.5 });
      }
    });

    slide.addFootnote("Process flow generated by DocLayer");
    slide.addPageNumber("1");

    return await builder.build();
  }
}

// Export main classes and utilities
export default DocLayerClient;

// Re-export for convenience
export { DocLayerClient, PresentationBuilder, SlideBuilder, DocLayerUtils };