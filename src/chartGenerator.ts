import PptxGenJS from 'pptxgenjs';
import { ChartRequest, ChartData, ChartOptions } from './types';

const DEFAULT_COLORS = [
  '0088FE', '00C49F', 'FFBB28', 'FF8042', 
  '8884D8', '82CA9D', 'FFC658', 'FF6B9D'
];

export class ChartGenerator {
  private pptx: PptxGenJS;

  constructor() {
    this.pptx = new PptxGenJS();
  }

  generate(request: ChartRequest): Buffer {
    const slide = this.pptx.addSlide();
    const { type, data, options = {} } = request;

    // Add title if provided
    if (options.title) {
      slide.addText(options.title, {
        x: 0.5,
        y: 0.5,
        w: 9,
        h: 0.5,
        fontSize: 24,
        bold: true,
        color: '363636'
      });
    }

    const chartY = options.title ? 1.5 : 0.5;
    const colors = options.colors || DEFAULT_COLORS;

    switch (type) {
      case 'bar':
        this.addBarChart(slide, data, options, chartY, colors);
        break;
      case 'line':
        this.addLineChart(slide, data, options, chartY, colors);
        break;
      case 'pie':
        this.addPieChart(slide, data, options, chartY, colors);
        break;
      case 'area':
        this.addAreaChart(slide, data, options, chartY, colors);
        break;
      case 'scatter':
        this.addScatterChart(slide, data, options, chartY, colors);
        break;
      default:
        throw new Error(`Unsupported chart type: ${type}`);
    }

    return this.pptx.write({ outputType: 'nodebuffer' }) as Promise<Buffer> as any;
  }

  private addBarChart(
    slide: any,
    data: ChartData,
    options: ChartOptions,
    y: number,
    colors: string[]
  ) {
    const chartData = this.formatChartData(data, colors);
    
    slide.addChart(this.pptx.ChartType.bar, chartData, {
      x: 1,
      y,
      w: 8,
      h: 4.5,
      showLegend: options.legend !== false,
      showTitle: false,
      catAxisLabelFontSize: 11,
      valAxisLabelFontSize: 11,
    });
  }

  private addLineChart(
    slide: any,
    data: ChartData,
    options: ChartOptions,
    y: number,
    colors: string[]
  ) {
    const chartData = this.formatChartData(data, colors);
    
    slide.addChart(this.pptx.ChartType.line, chartData, {
      x: 1,
      y,
      w: 8,
      h: 4.5,
      showLegend: options.legend !== false,
      showTitle: false,
      catAxisLabelFontSize: 11,
      valAxisLabelFontSize: 11,
      lineSmooth: true,
    });
  }

  private addAreaChart(
    slide: any,
    data: ChartData,
    options: ChartOptions,
    y: number,
    colors: string[]
  ) {
    const chartData = this.formatChartData(data, colors);
    
    slide.addChart(this.pptx.ChartType.area, chartData, {
      x: 1,
      y,
      w: 8,
      h: 4.5,
      showLegend: options.legend !== false,
      showTitle: false,
      catAxisLabelFontSize: 11,
      valAxisLabelFontSize: 11,
    });
  }

  private addPieChart(
    slide: any,
    data: ChartData,
    options: ChartOptions,
    y: number,
    colors: string[]
  ) {
    // For pie charts, use first dataset only
    const dataset = data.datasets[0];
    const chartData = data.labels.map((label, idx) => ({
      name: label,
      labels: [label],
      values: [dataset.data[idx]],
    }));

    slide.addChart(this.pptx.ChartType.pie, chartData, {
      x: 1.5,
      y,
      w: 7,
      h: 4.5,
      showLegend: options.legend !== false,
      showTitle: false,
      dataLabelFontSize: 11,
      chartColors: colors,
    });
  }

  private addScatterChart(
    slide: any,
    data: ChartData,
    options: ChartOptions,
    y: number,
    colors: string[]
  ) {
    const chartData = this.formatChartData(data, colors);
    
    slide.addChart(this.pptx.ChartType.scatter, chartData, {
      x: 1,
      y,
      w: 8,
      h: 4.5,
      showLegend: options.legend !== false,
      showTitle: false,
      catAxisLabelFontSize: 11,
      valAxisLabelFontSize: 11,
    });
  }

  private formatChartData(data: ChartData, colors: string[]) {
    return data.datasets.map((dataset, idx) => ({
      name: dataset.label,
      labels: data.labels,
      values: dataset.data,
      color: dataset.borderColor?.replace('#', '') || colors[idx % colors.length],
    }));
  }
}

