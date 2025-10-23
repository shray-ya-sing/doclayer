import express, { Request, Response } from 'express';
import cors from 'cors';
import { ChartGenerator } from './chartGenerator';
import { ChartRequest } from './types';

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

app.post('/api/chart-to-pptx', async (req: Request, res: Response) => {
  try {
    const chartRequest: ChartRequest = req.body;

    // Validation
    if (!chartRequest.type || !chartRequest.data) {
      return res.status(400).json({
        success: false,
        error: 'Missing required fields: type and data',
      });
    }

    if (!chartRequest.data.labels || !chartRequest.data.datasets) {
      return res.status(400).json({
        success: false,
        error: 'data must include labels and datasets',
      });
    }

    const generator = new ChartGenerator();
    const pptxBuffer = await generator.generate(chartRequest);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename=chart.pptx');
    res.send(pptxBuffer);
  } catch (error: any) {
    console.error('Error generating PPTX:', error);
    res.status(500).json({
      success: false,
      error: error.message || 'Failed to generate PPTX',
    });
  }
});

app.get('/health', (req: Request, res: Response) => {
  res.json({ status: 'ok' });
});

app.listen(PORT, () => {
  console.log(`DocLayer API running on http://localhost:${PORT}`);
  console.log(`POST /api/chart-to-pptx - Generate PPTX from chart data`);
});

