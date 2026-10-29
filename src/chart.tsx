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
  LabelList,
  PieChart,
  Pie,
  Legend
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

// Tooltip مخصص للمخطط الدائري والشريطي
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
          النسبة المئوية من الفصل: <b style={{ color: data.color }}>{data.percentage}%</b>
        </div>
      </div>
    );
  }
  return null;
};

// تسميات النسب المئوية داخل شرائح المخطط الدائري
const renderCustomizedPieLabel = ({
  cx,
  cy,
  midAngle,
  innerRadius,
  outerRadius,
  percent
}: any) => {
  if (!percent || percent < 0.05) return null;
  const RADIAN = Math.PI / 180;
  const radius = innerRadius + (outerRadius - innerRadius) * 0.58;
  const x = cx + radius * Math.cos(-midAngle * RADIAN);
  const y = cy + radius * Math.sin(-midAngle * RADIAN);

  return (
    <text
      x={x}
      y={y}
      fill="#ffffff"
      textAnchor="middle"
      dominantBaseline="central"
      style={{
        fontFamily: "'Cairo', Tahoma, sans-serif",
        fontSize: '12px',
        fontWeight: 'bold',
        filter: 'drop-shadow(0 1px 2px rgba(0,0,0,0.65))'
      }}
    >
      {`${(percent * 100).toFixed(1)}%`}
    </text>
  );
};

export const AnalysisCharts: React.FC = () => {
  const [counts, setCounts] = useState<ChartDataCounts>({
    mmtaz: 0,
    jayidJiddan: 0,
    jayid: 0,
    mutawassit: 0,
    daif: 0,
    rasib: 0,
    total: 0
  });

  // نمط العرض: 'both' (المخططان معاً)، 'pie' (المخطط الدائري)، 'bar' (المخطط الشريطي)
  const [viewMode, setViewMode] = useState<'both' | 'pie' | 'bar'>('both');

  useEffect(() => {
    window.updateAnalysisChart = (newCounts: ChartDataCounts) => {
      setCounts({ ...newCounts });
    };

    // قراءة البيانات الابتدائية من الجدول عند تحميل الصفحة
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
      // تجاهل الخطأ في حالة عدم توفر العناصر فورياً
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
      value: val,
      percentage: pct,
      color: cfg.color
    };
  });

  // شرائح المخطط الدائري التي تحتوي على طلاب فقط لمنع الشرائح الفارغة
  const pieData = chartData.filter(d => d.count > 0);
  const maxCount = Math.max(...chartData.map(d => d.count), 5);

  // نسبة النجاح (الحاصلين على 50% فأعلى: ممتاز، جيد جداً، جيد، متوسط)
  const passCount = counts.mmtaz + counts.jayidJiddan + counts.jayid + counts.mutawassit;
  const passRate = totalStudents > 0 ? ((passCount / totalStudents) * 100).toFixed(1) : '0.0';

  return (
    <div
      style={{
        width: '100%',
        fontFamily: "'Cairo', Tahoma, sans-serif",
        direction: 'rtl'
      }}
    >
      {/* شريط رأس الرسوم البيانية وأزرار التبديل */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          flexWrap: 'wrap',
          gap: '12px',
          marginBottom: '16px',
          borderBottom: '1.5px solid #eef2f7',
          paddingBottom: '12px'
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <span style={{ fontSize: '24px' }}>📊</span>
          <div>
            <h3 style={{ margin: 0, fontSize: '16.5px', color: '#1a5276', fontWeight: 800 }}>
              الرسوم البيانية لتحليل نتائج وتقديرات الاختبار (Recharts)
            </h3>
            <p style={{ margin: '3px 0 0', fontSize: '12px', color: '#64748b' }}>
              المخطط الدائري (Pie Chart) يوضح النسب المئوية، والمخطط الشريطي (Bar Chart) يوضح أعداد الطلاب
            </p>
          </div>
        </div>

        {/* أزرار اختيار نمط العرض */}
        <div
          style={{
            display: 'inline-flex',
            alignItems: 'center',
            background: '#f1f5f9',
            borderRadius: '10px',
            padding: '3px',
            gap: '3px'
          }}
        >
          <button
            type="button"
            onClick={() => setViewMode('both')}
            style={{
              border: 'none',
              borderRadius: '8px',
              padding: '6px 12px',
              fontSize: '12px',
              fontWeight: 700,
              cursor: 'pointer',
              fontFamily: 'inherit',
              transition: 'all 0.2s',
              background: viewMode === 'both' ? '#1a5276' : 'transparent',
              color: viewMode === 'both' ? '#ffffff' : '#475569',
              boxShadow: viewMode === 'both' ? '0 2px 6px rgba(26,82,118,0.25)' : 'none'
            }}
          >
            🔀 المخططان معاً
          </button>
          <button
            type="button"
            onClick={() => setViewMode('pie')}
            style={{
              border: 'none',
              borderRadius: '8px',
              padding: '6px 12px',
              fontSize: '12px',
              fontWeight: 700,
              cursor: 'pointer',
              fontFamily: 'inherit',
              transition: 'all 0.2s',
              background: viewMode === 'pie' ? '#1a5276' : 'transparent',
              color: viewMode === 'pie' ? '#ffffff' : '#475569',
              boxShadow: viewMode === 'pie' ? '0 2px 6px rgba(26,82,118,0.25)' : 'none'
            }}
          >
            🥧 المخطط الدائري (النسب المئوية)
          </button>
          <button
            type="button"
            onClick={() => setViewMode('bar')}
            style={{
              border: 'none',
              borderRadius: '8px',
              padding: '6px 12px',
              fontSize: '12px',
              fontWeight: 700,
              cursor: 'pointer',
              fontFamily: 'inherit',
              transition: 'all 0.2s',
              background: viewMode === 'bar' ? '#1a5276' : 'transparent',
              color: viewMode === 'bar' ? '#ffffff' : '#475569',
              boxShadow: viewMode === 'bar' ? '0 2px 6px rgba(26,82,118,0.25)' : 'none'
            }}
          >
            📊 المخطط الشريطي (الأعداد)
          </button>
        </div>
      </div>

      {/* شريط الإحصاءات السريعة فوق المخططات */}
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fit, minmax(130px, 1fr))',
          gap: '8px',
          marginBottom: '18px'
        }}
      >
        <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '10px', padding: '8px 12px', textAlign: 'center' }}>
          <div style={{ fontSize: '11px', color: '#64748b', fontWeight: 600 }}>إجمالي الطلاب المرصودين</div>
          <div style={{ fontSize: '18px', fontWeight: 800, color: '#1a5276' }}>{totalStudents}</div>
        </div>
        <div style={{ background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: '10px', padding: '8px 12px', textAlign: 'center' }}>
          <div style={{ fontSize: '11px', color: '#166534', fontWeight: 600 }}>نسبة النجاح العامة</div>
          <div style={{ fontSize: '18px', fontWeight: 800, color: '#15803d' }}>{passRate}%</div>
        </div>
        <div style={{ background: '#ecfdf5', border: '1px solid #a7f3d0', borderRadius: '10px', padding: '8px 12px', textAlign: 'center' }}>
          <div style={{ fontSize: '11px', color: '#047857', fontWeight: 600 }}>الطلاب المتفوقون (ممتاز)</div>
          <div style={{ fontSize: '18px', fontWeight: 800, color: '#059669' }}>{counts.mmtaz}</div>
        </div>
        <div style={{ background: '#fff1f2', border: '1px solid #fecdd3', borderRadius: '10px', padding: '8px 12px', textAlign: 'center' }}>
          <div style={{ fontSize: '11px', color: '#9f1239', fontWeight: 600 }}>المحتاجون للمتابعة</div>
          <div style={{ fontSize: '18px', fontWeight: 800, color: '#e11d48' }}>{counts.daif + counts.rasib}</div>
        </div>
      </div>

      {/* منطقة عرض الرسوم البيانية */}
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: viewMode === 'both' ? 'repeat(auto-fit, minmax(320px, 1fr))' : '1fr',
          gap: '20px',
          alignItems: 'stretch'
        }}
      >
        {/* 1. المخطط الدائري (Pie Chart): النسبة المئوية لتوزيع التقديرات */}
        {(viewMode === 'both' || viewMode === 'pie') && (
          <div
            style={{
              background: '#ffffff',
              border: '1.5px solid #d4e6f1',
              borderRadius: '14px',
              padding: '16px 14px',
              boxShadow: '0 2px 8px rgba(0,0,0,0.03)',
              display: 'flex',
              flexDirection: 'column'
            }}
          >
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '8px' }}>
              <div style={{ fontWeight: 800, fontSize: '14px', color: '#1a5276', display: 'flex', alignItems: 'center', gap: '6px' }}>
                <span>🥧</span>
                <span>المخطط الدائري: النسبة المئوية لكل تقدير (Pie Chart)</span>
              </div>
              <span style={{ fontSize: '11px', color: '#15803d', background: '#dcfce7', padding: '2px 8px', borderRadius: '12px', fontWeight: 700 }}>
                توزيع نسبي %
              </span>
            </div>

            {totalStudents === 0 ? (
              <div
                style={{
                  height: '280px',
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  justifyContent: 'center',
                  color: '#94a3b8',
                  fontSize: '13px',
                  gap: '8px'
                }}
              >
                <span style={{ fontSize: '32px' }}>🥧</span>
                <span>بانتظار رصد درجات الطلاب لعرض المخطط الدائري</span>
              </div>
            ) : (
              <>
                <div style={{ width: '100%', height: '260px' }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Tooltip content={<CustomTooltip />} />
                      <Pie
                        data={pieData}
                        dataKey="value"
                        nameKey="name"
                        cx="50%"
                        cy="50%"
                        outerRadius={95}
                        innerRadius={36}
                        paddingAngle={3}
                        labelLine={false}
                        label={renderCustomizedPieLabel}
                        animationDuration={700}
                      >
                        {pieData.map((entry, index) => (
                          <Cell
                            key={`cell-${index}`}
                            fill={entry.color}
                            stroke="#ffffff"
                            strokeWidth={2}
                          />
                        ))}
                      </Pie>
                    </PieChart>
                  </ResponsiveContainer>
                </div>

                {/* مفتاح المخطط الدائري (Legend) ببيانات النسب المئوية */}
                <div
                  style={{
                    display: 'flex',
                    flexWrap: 'wrap',
                    justifyContent: 'center',
                    gap: '6px 12px',
                    marginTop: '8px',
                    paddingTop: '10px',
                    borderTop: '1px dashed #e2e8f0',
                    fontSize: '11.5px'
                  }}
                >
                  {chartData.map((item, idx) => (
                    <div
                      key={idx}
                      style={{
                        display: 'inline-flex',
                        alignItems: 'center',
                        gap: '5px',
                        opacity: item.count > 0 ? 1 : 0.5
                      }}
                    >
                      <span
                        style={{
                          width: '10px',
                          height: '10px',
                          borderRadius: '50%',
                          backgroundColor: item.color,
                          display: 'inline-block'
                        }}
                      />
                      <span style={{ fontWeight: 700, color: '#334155' }}>{item.name}:</span>
                      <span style={{ fontWeight: 800, color: item.color }}>{item.percentage}%</span>
                      <span style={{ color: '#94a3b8', fontSize: '10.5px' }}>({item.count})</span>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>
        )}

        {/* 2. المخطط الشريطي (Bar Chart): أعداد الطلاب لكل تقدير */}
        {(viewMode === 'both' || viewMode === 'bar') && (
          <div
            style={{
              background: '#ffffff',
              border: '1.5px solid #d4e6f1',
              borderRadius: '14px',
              padding: '16px 14px',
              boxShadow: '0 2px 8px rgba(0,0,0,0.03)',
              display: 'flex',
              flexDirection: 'column'
            }}
          >
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '8px' }}>
              <div style={{ fontWeight: 800, fontSize: '14px', color: '#1a5276', display: 'flex', alignItems: 'center', gap: '6px' }}>
                <span>📊</span>
                <span>المخطط الشريطي: أعداد الطلاب لكل فئة (Bar Chart)</span>
              </div>
              <span style={{ fontSize: '11px', color: '#0369a1', background: '#e0f2fe', padding: '2px 8px', borderRadius: '12px', fontWeight: 700 }}>
                تكرار عددي 👤
              </span>
            </div>

            <div style={{ width: '100%', height: '280px' }}>
              <ResponsiveContainer width="100%" height="100%">
                <BarChart
                  data={chartData}
                  margin={{ top: 25, right: 10, left: 10, bottom: 5 }}
                >
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                  <XAxis
                    dataKey="name"
                    tick={{ fill: '#334155', fontWeight: 'bold', fontSize: 12 }}
                    axisLine={{ stroke: '#cbd5e1' }}
                    tickLine={false}
                  />
                  <YAxis
                    allowDecimals={false}
                    domain={[0, Math.ceil(maxCount * 1.15)]}
                    tick={{ fill: '#64748b', fontSize: 11 }}
                    axisLine={{ stroke: '#cbd5e1' }}
                    tickLine={false}
                    width={32}
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
                        fontSize: 12,
                        fontFamily: "'Cairo', Tahoma, sans-serif"
                      }}
                    />
                    {chartData.map((entry, index) => (
                      <Cell key={`cell-bar-${index}`} fill={entry.color} />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>

            <div style={{ textAlign: 'center', fontSize: '11px', color: '#64748b', marginTop: '6px' }}>
              كل شريط يمثل عدد الطلاب في فئة التقدير مع نطاق الدرجات المخصص
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

// التركيب التلقائي عند جاهزية الـ DOM
function mountChartRoot() {
  const container = document.getElementById('analysisChartRoot');
  if (container) {
    const root = createRoot(container);
    root.render(<AnalysisCharts />);
  }
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', mountChartRoot);
} else {
  mountChartRoot();
}
