# DocLayer

Simple API to convert chart data to PowerPoint presentations. Built for developers who need to automate report generation.

## Features

- 📊 Support for bar, line, pie, area, and scatter charts
- 🎨 Customizable colors and styling
- 📝 Native PowerPoint charts (editable, not images)
- 🚀 Chart.js-compatible API
- ⚡ Fast and lightweight

## Installation

```bash
npm install
```

## Usage

### Start the server

```bash
npm run dev
```

### Generate a chart

```bash
curl -X POST http://localhost:3000/api/chart-to-pptx \
  -H "Content-Type: application/json" \
  -d '{
    "type": "bar",
    "data": {
      "labels": ["Q1", "Q2", "Q3", "Q4"],
      "datasets": [{
        "label": "Revenue",
        "data": [45000, 52000, 61000, 58000]
      }]
    },
    "options": {
      "title": "Quarterly Revenue"
    }
  }' \
  --output chart.pptx
```

## API Reference

### `POST /api/chart-to-pptx`

**Request Body:**

```typescript
{
  type: 'bar' | 'line' | 'pie' | 'scatter' | 'area',
  data: {
    labels: string[],
    datasets: [{
      label: string,
      data: number[],
      backgroundColor?: string | string[],
      borderColor?: string,
      borderWidth?: number
    }]
  },
  options?: {
    title?: string,
    legend?: boolean,
    colors?: string[],
    width?: number,
    height?: number
  }
}
```

**Response:** PPTX file (binary)

## Examples

See `examples/sample-requests.json` for more examples.

### JavaScript/Node.js

```javascript
const axios = require('axios');
const fs = require('fs');

async function generateChart() {
  const response = await axios.post('http://localhost:3000/api/chart-to-pptx', {
    type: 'line',
    data: {
      labels: ['Jan', 'Feb', 'Mar'],
      datasets: [{
        label: 'Sales',
        data: [12000, 19000, 15000]
      }]
    },
    options: {
      title: 'Monthly Sales'
    }
  }, {
    responseType: 'arraybuffer'
  });

  fs.writeFileSync('output.pptx', response.data);
}
```

### Python

```python
import requests

response = requests.post('http://localhost:3000/api/chart-to-pptx', json={
    'type': 'bar',
    'data': {
        'labels': ['Q1', 'Q2', 'Q3'],
        'datasets': [{
            'label': 'Revenue',
            'data': [45000, 52000, 61000]
        }]
    },
    'options': {
        'title': 'Quarterly Report'
    }
})

with open('chart.pptx', 'wb') as f:
    f.write(response.content)
```

## License

MIT

