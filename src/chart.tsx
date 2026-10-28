import React, { useState, useEffect } from 'react';
import { createRoot } from 'react-dom/client';
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Cell,
  LabelList
} from 'recharts';

export interface ChartDataCounts {
  mmtaz: number;
  jayidJiddan: number;
  jayid: number;
  mutawassit: number;
  daif: number;
  rasib: number;
  total: number;
}

declare global {
  interface Window {
    updateAnalysisChart?: (counts: ChartDataCounts) => void;
  }
}

const GRADE_CONFIG = [
  { key: 'mmtaz', name: 'ممتاز', range: '90 - 100%', color: '#10b981', bg: '#ecfdf5' },
  { key: 'jayidJiddan', name: 'جيد جداً', range: '80 - 89%', color: '#0284c7', bg: '#f0f9ff' },
  { key: 'jayid', name: 'جيد', range: '65 - 79%', color: '#d97706', bg: '#fffbeb' },
  { key: 'mutawassit', name: 'متوسط', range: '50 - 64%', color: '#ea580c', bg: '#fff7ed' },
  { key: 'daif', name: 'ضعيف', range: '40 - 49%', color: '#e11d48', bg: '#fff1f2' },
  { key: 'rasib', name: 'راسب', range: 'أقل من 40%', color: '#991b1b', bg: '#fef2f2' }
];

const CustomTooltip = ({ active, payload }: any) => {
  if (active && payload && payload.length) {
    const data = payload[0].payload;
    return (
      <div
        style={{
          background: '#ffffff',
          border: `2px solid ${data.color}`,
          borderRadius: '10px',
          padding: '10px 14px',
          boxShadow: '0 4px 14px rgba(0,0,0,0.12)',
          direction: 'rtl',
          textAlign: 'right',
          fontFamily: "'Cairo', Tahoma, sans-serif",
          fontSize: '13px'
        }}
      >
        <div style={{ fontWeight: 'bold', color: data.color, marginBottom: '4px', fontSize: '14px' }}>
          {data.name} ({data.range})
        </div>
        <div style={{ color: '#1e293b', fontWeight: 'bold' }}>
          عدد الطلاب: <span style={{ fontSize: '15px', color: data.color }}>{data.count}</span> طالب
        </div>
        <div style={{ color: '#64748b', fontSize: '12px', marginTop: '3px' }}>
          النسبة من الفصل: {data.percentage}%
        </div>
      </div>
    );
  }
  return null;
};

export const AnalysisBarChart: React.FC = () => {
  const [counts, setCounts] = useState<ChartDataCounts>({
    mmtaz: 0,
    jayidJiddan: 0,
    jayid: 0,
    mutawassit: 0,
    daif: 0,
    rasib: 0,
    total: 0
  });

  useEffect(() => {
    window.updateAnalysisChart = (newCounts: ChartDataCounts) => {
      setCounts({ ...newCounts });
    };

    // Pull initial counts from DOM if already present
    try {
      const getNum = (id: string) => {
        const el = document.getElementById(id);
        return el ? (parseInt(el.innerText, 10) || 0) : 0;
      };
      const mmtaz = getNum('count-mmtaz');
      const jayidJiddan = getNum('count-jayid-jiddan');
      const jayid = getNum('count-jayid');
      const mutawassit = getNum('count-mutawassit');
      const daif = getNum('count-daif');
      const rasib = getNum('kpi-val-rasib');
      const total = mmtaz + jayidJiddan + jayid + mutawassit + daif + rasib;

      setCounts({
        mmtaz,
        jayidJiddan,
        jayid,
        mutawassit,
        daif,
        rasib,
        total
      });
    } catch (e) {
      // Ignore
    }
  }, []);

  const totalStudents = counts.total > 0
    ? counts.total
    : (counts.mmtaz + counts.jayidJiddan + counts.jayid + counts.mutawassit + counts.daif + counts.rasib);

  const chartData = GRADE_CONFIG.map(cfg => {
    const val = (counts as any)[cfg.key] || 0;
    const pct = totalStudents > 0 ? ((val / totalStudents) * 100).toFixed(1) : '0';
    return {
      name: cfg.name,
      range: cfg.range,
      count: val,
      percentage: pct,
      color: cfg.color
    };
  });

  const maxCount = Math.max(...chartData.map(d => d.count), 5);

  return (
    <div
      style={{
        width: '100%',
        height: '100%',
        fontFamily: "'Cairo', Tahoma, sans-serif",
        direction: 'rtl'
      }}
    >
      <ResponsiveContainer width="100%" height={300}>
        <BarChart
          data={chartData}
          margin={{ top: 25, right: 20, left: 20, bottom: 10 }}
        >
          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
          <XAxis
            dataKey="name"
            tick={{ fill: '#334155', fontWeight: 'bold', fontSize: 13 }}
            axisLine={{ stroke: '#cbd5e1' }}
            tickLine={false}
          />
          <YAxis
            allowDecimals={false}
            domain={[0, Math.ceil(maxCount * 1.15)]}
            tick={{ fill: '#64748b', fontSize: 12 }}
            axisLine={{ stroke: '#cbd5e1' }}
            tickLine={false}
            label={{ value: 'عدد الطلاب', angle: -90, position: 'insideLeft', offset: -10, fill: '#64748b', fontSize: 12 }}
          />
          <Tooltip content={<CustomTooltip />} />
          <Bar
            dataKey="count"
            radius={[6, 6, 0, 0]}
            animationDuration={600}
          >
            <LabelList
              dataKey="count"
              position="top"
              style={{
                fill: '#1e293b',
                fontWeight: 'bold',
                fontSize: 13,
                fontFamily: "'Cairo', Tahoma, sans-serif"
              }}
            />
            {chartData.map((entry, index) => (
              <Cell key={`cell-${index}`} fill={entry.color} />
            ))}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
};

// Auto-mount when DOM is ready
function mountChartRoot() {
  const container = document.getElementById('analysisChartRoot');
  if (container) {
    const root = createRoot(container);
    root.render(<AnalysisBarChart />);
  }
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', mountChartRoot);
} else {
  mountChartRoot();
}
