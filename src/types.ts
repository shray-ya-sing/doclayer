export type ChartType = 'bar' | 'line' | 'pie' | 'scatter' | 'area';

export interface Dataset {
  label: string;
  data: number[];
  backgroundColor?: string | string[];
  borderColor?: string;
  borderWidth?: number;
}

export interface ChartData {
  labels: string[];
  datasets: Dataset[];
}

export interface ChartOptions {
  title?: string;
  legend?: boolean;
  colors?: string[];
  width?: number;
  height?: number;
}

export interface ChartRequest {
  type: ChartType;
  data: ChartData;
  options?: ChartOptions;
}

export interface PptxResponse {
  success: boolean;
  message?: string;
  error?: string;
}

