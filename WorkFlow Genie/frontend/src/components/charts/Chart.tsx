/**
 * Chart component for data visualization
 * Supports multiple chart types: bar, line, pie, scatter, area, column
 */

import React, { useEffect, useRef } from 'react';
import type { Chart as ChartType } from '../../types/index';
import { useSpreadsheetStore } from '../../stores/spreadsheetStore';

interface ChartProps {
  chart: ChartType;
  onClose?: () => void;
  onResize?: (size: { width: number; height: number }) => void;
}

export const Chart: React.FC<ChartProps> = ({ chart, onClose }) => {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const { currentWorksheet } = useSpreadsheetStore();

  useEffect(() => {
    if (!canvasRef.current || !currentWorksheet) return;

    const canvas = canvasRef.current;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    // Clear canvas
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    // Extract data from range
    const data = extractChartData(currentWorksheet.cells, chart.dataRange);

    // Draw chart based on type
    switch (chart.type) {
      case 'bar':
        drawBarChart(ctx, data, canvas.width, canvas.height, chart);
        break;
      case 'column':
        drawColumnChart(ctx, data, canvas.width, canvas.height, chart);
        break;
      case 'line':
        drawLineChart(ctx, data, canvas.width, canvas.height, chart);
        break;
      case 'pie':
        drawPieChart(ctx, data, canvas.width, canvas.height, chart);
        break;
      case 'scatter':
        drawScatterChart(ctx, data, canvas.width, canvas.height, chart);
        break;
      case 'area':
        drawAreaChart(ctx, data, canvas.width, canvas.height, chart);
        break;
    }
  }, [chart, currentWorksheet]);

  return (
    <div
      className="absolute bg-white border-2 border-gray-300 rounded-lg shadow-lg p-4"
      style={{
        left: chart.position.col * 100,
        top: chart.position.row * 24,
        width: chart.size.width,
        height: chart.size.height,
        zIndex: 100,
      }}
    >
      {/* Chart header */}
      <div className="flex items-center justify-between mb-2">
        <h3 className="font-semibold text-sm">{chart.title || 'Chart'}</h3>
        <div className="flex gap-1">
          <button
            onClick={onClose}
            className="p-1 hover:bg-gray-100 rounded text-gray-500 hover:text-gray-700"
          >
            ✕
          </button>
        </div>
      </div>

      {/* Chart canvas */}
      <canvas
        ref={canvasRef}
        width={chart.size.width - 32}
        height={chart.size.height - 64}
        className="w-full"
      />
    </div>
  );
};

/**
 * Extract data from cells for charting
 */
const extractChartData = (
  cells: Record<string, any>,
  range: any
): { labels: string[]; values: number[] } => {
  const labels: string[] = [];
  const values: number[] = [];

  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let col = range.startCol; col <= range.endCol; col++) {
      const key = `${row},${col}`;
      const cell = cells[key];

      if (col === range.startCol) {
        // First column is labels
        labels.push(cell?.value?.toString() || '');
      } else {
        // Other columns are values
        const value = parseFloat(cell?.value);
        if (!isNaN(value)) {
          values.push(value);
        }
      }
    }
  }

  return { labels, values };
};

/**
 * Draw bar chart
 */
const drawBarChart = (
  ctx: CanvasRenderingContext2D,
  data: { labels: string[]; values: number[] },
  width: number,
  height: number,
  _chart: ChartType
) => {
  const padding = 40;
  const barHeight = (height - padding * 2) / data.values.length;
  const maxValue = Math.max(...data.values);

  // Draw axes
  ctx.strokeStyle = '#333';
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(padding, padding);
  ctx.lineTo(padding, height - padding);
  ctx.lineTo(width - padding, height - padding);
  ctx.stroke();

  // Draw bars
  data.values.forEach((value, i) => {
    const barWidth = ((value / maxValue) * (width - padding * 2));
    const y = padding + i * barHeight + 5;

    // Bar
    ctx.fillStyle = `hsl(${(i * 360) / data.values.length}, 70%, 60%)`;
    ctx.fillRect(padding, y, barWidth, barHeight - 10);

    // Label
    ctx.fillStyle = '#333';
    ctx.font = '12px Arial';
    ctx.textAlign = 'right';
    ctx.fillText(data.labels[i] || '', padding - 5, y + barHeight / 2);

    // Value
    ctx.textAlign = 'left';
    ctx.fillText(value.toString(), padding + barWidth + 5, y + barHeight / 2);
  });
};

/**
 * Draw column chart
 */
const drawColumnChart = (
  ctx: CanvasRenderingContext2D,
  data: { labels: string[]; values: number[] },
  width: number,
  height: number,
  _chart: ChartType
) => {
  const padding = 40;
  const barWidth = (width - padding * 2) / data.values.length;
  const maxValue = Math.max(...data.values);

  // Draw axes
  ctx.strokeStyle = '#333';
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(padding, padding);
  ctx.lineTo(padding, height - padding);
  ctx.lineTo(width - padding, height - padding);
  ctx.stroke();

  // Draw columns
  data.values.forEach((value, i) => {
    const barHeight = ((value / maxValue) * (height - padding * 2));
    const x = padding + i * barWidth + 5;

    // Column
    ctx.fillStyle = `hsl(${(i * 360) / data.values.length}, 70%, 60%)`;
    ctx.fillRect(x, height - padding - barHeight, barWidth - 10, barHeight);

    // Label
    ctx.fillStyle = '#333';
    ctx.font = '12px Arial';
    ctx.textAlign = 'center';
    ctx.fillText(data.labels[i] || '', x + barWidth / 2, height - padding + 20);
  });
};

/**
 * Draw line chart
 */
const drawLineChart = (
  ctx: CanvasRenderingContext2D,
  data: { labels: string[]; values: number[] },
  width: number,
  height: number,
  _chart: ChartType
) => {
  const padding = 40;
  const stepX = (width - padding * 2) / (data.values.length - 1 || 1);
  const maxValue = Math.max(...data.values);
  const minValue = Math.min(...data.values);
  const range = maxValue - minValue || 1;

  // Draw axes
  ctx.strokeStyle = '#333';
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(padding, padding);
  ctx.lineTo(padding, height - padding);
  ctx.lineTo(width - padding, height - padding);
  ctx.stroke();

  // Draw line
  ctx.strokeStyle = '#107c41';
  ctx.lineWidth = 3;
  ctx.beginPath();

  data.values.forEach((value, i) => {
    const x = padding + i * stepX;
    const y = height - padding - ((value - minValue) / range) * (height - padding * 2);

    if (i === 0) {
      ctx.moveTo(x, y);
    } else {
      ctx.lineTo(x, y);
    }
  });

  ctx.stroke();

  // Draw points
  data.values.forEach((value, i) => {
    const x = padding + i * stepX;
    const y = height - padding - ((value - minValue) / range) * (height - padding * 2);

    ctx.fillStyle = '#107c41';
    ctx.beginPath();
    ctx.arc(x, y, 4, 0, Math.PI * 2);
    ctx.fill();

    // Label
    ctx.fillStyle = '#333';
    ctx.font = '12px Arial';
    ctx.textAlign = 'center';
    ctx.fillText(data.labels[i] || '', x, height - padding + 20);
  });
};

/**
 * Draw pie chart
 */
const drawPieChart = (
  ctx: CanvasRenderingContext2D,
  data: { labels: string[]; values: number[] },
  width: number,
  height: number,
  _chart: ChartType
) => {
  const centerX = width / 2;
  const centerY = height / 2;
  const radius = Math.min(width, height) / 2 - 40;
  const total = data.values.reduce((sum, v) => sum + v, 0);

  let currentAngle = -Math.PI / 2;

  data.values.forEach((value, i) => {
    const sliceAngle = (value / total) * Math.PI * 2;

    // Draw slice
    ctx.fillStyle = `hsl(${(i * 360) / data.values.length}, 70%, 60%)`;
    ctx.beginPath();
    ctx.moveTo(centerX, centerY);
    ctx.arc(centerX, centerY, radius, currentAngle, currentAngle + sliceAngle);
    ctx.closePath();
    ctx.fill();

    // Draw label
    const labelAngle = currentAngle + sliceAngle / 2;
    const labelX = centerX + Math.cos(labelAngle) * (radius * 0.7);
    const labelY = centerY + Math.sin(labelAngle) * (radius * 0.7);

    ctx.fillStyle = '#fff';
    ctx.font = 'bold 12px Arial';
    ctx.textAlign = 'center';
    ctx.fillText(`${((value / total) * 100).toFixed(1)}%`, labelX, labelY);

    currentAngle += sliceAngle;
  });

  // Draw legend
  const legendX = width - 100;
  let legendY = 20;

  data.labels.forEach((label, i) => {
    ctx.fillStyle = `hsl(${(i * 360) / data.values.length}, 70%, 60%)`;
    ctx.fillRect(legendX, legendY, 15, 15);

    ctx.fillStyle = '#333';
    ctx.font = '12px Arial';
    ctx.textAlign = 'left';
    ctx.fillText(label, legendX + 20, legendY + 12);

    legendY += 20;
  });
};

/**
 * Draw scatter chart
 */
const drawScatterChart = (
  ctx: CanvasRenderingContext2D,
  data: { labels: string[]; values: number[] },
  width: number,
  height: number,
  _chart: ChartType
) => {
  const padding = 40;
  const maxValue = Math.max(...data.values);

  // Draw axes
  ctx.strokeStyle = '#333';
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(padding, padding);
  ctx.lineTo(padding, height - padding);
  ctx.lineTo(width - padding, height - padding);
  ctx.stroke();

  // Draw points
  data.values.forEach((value, i) => {
    const x = padding + (i / data.values.length) * (width - padding * 2);
    const y = height - padding - (value / maxValue) * (height - padding * 2);

    ctx.fillStyle = `hsl(${(i * 360) / data.values.length}, 70%, 60%)`;
    ctx.beginPath();
    ctx.arc(x, y, 6, 0, Math.PI * 2);
    ctx.fill();
  });
};

/**
 * Draw area chart
 */
const drawAreaChart = (
  ctx: CanvasRenderingContext2D,
  data: { labels: string[]; values: number[] },
  width: number,
  height: number,
  _chart: ChartType
) => {
  const padding = 40;
  const stepX = (width - padding * 2) / (data.values.length - 1 || 1);
  const maxValue = Math.max(...data.values);

  // Draw axes
  ctx.strokeStyle = '#333';
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(padding, padding);
  ctx.lineTo(padding, height - padding);
  ctx.lineTo(width - padding, height - padding);
  ctx.stroke();

  // Draw area
  ctx.fillStyle = 'rgba(16, 124, 65, 0.3)';
  ctx.beginPath();
  ctx.moveTo(padding, height - padding);

  data.values.forEach((value, i) => {
    const x = padding + i * stepX;
    const y = height - padding - (value / maxValue) * (height - padding * 2);
    ctx.lineTo(x, y);
  });

  ctx.lineTo(padding + (data.values.length - 1) * stepX, height - padding);
  ctx.closePath();
  ctx.fill();

  // Draw line on top
  ctx.strokeStyle = '#107c41';
  ctx.lineWidth = 3;
  ctx.beginPath();

  data.values.forEach((value, i) => {
    const x = padding + i * stepX;
    const y = height - padding - (value / maxValue) * (height - padding * 2);

    if (i === 0) {
      ctx.moveTo(x, y);
    } else {
      ctx.lineTo(x, y);
    }
  });

  ctx.stroke();
};

export default Chart;
