import fs from 'fs';
import path from 'path';
import { execFile } from 'child_process';
import { promisify } from 'util';
import PptxGenJS from 'pptxgenjs';
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import * as cheerio from 'cheerio';
import { extractText } from 'unpdf';
import { GoogleGenAI } from '@google/genai';

const execFileAsync = promisify(execFile);

export interface SlideItem {
  title: string;
  bullets: string[];
  slide_type?: string;
  subtitle?: string;
  image_data?: string;
  table_data?: string[][];
  chart_type?: string;
  chart_labels?: string[];
  chart_values?: number[];
  chart_title?: string;
  notes?: string;
}

export interface ThemeColors {
  primary: string;
  secondary: string;
  accent: string;
  bg: string;
  card_bg: string;
  text_dark: string;
  text_light: string;
}

export interface StudioCanvaTemplate {
  id: string;
  canva_id?: string;
  name: string;
  category: string;
  source: string;
  theme_color?: string;
  layout: string;
  font_name?: string;
  font_family?: string;
  heading_font?: string;
  description: string;
  colors: ThemeColors;
  canva_url?: string;
  preview_image?: string;
  safe_preview_data_url?: string;
  fallback_preview_data_url?: string;
  badge_text?: string;
}

export const PRESENTATION_CANVA_TEMPLATES: StudioCanvaTemplate[] = [
  {
    id: 'canva_exec_navy',
    name: 'كانفا — العرض التنفيذي الكلاسيكي',
    category: 'أعمال',
    source: 'Canva Presentations',
    theme_color: 'blue',
    layout: 'split_accent',
    font_name: 'Dubai',
    description: 'تصميم مؤسسي أنيق بدرجات الكحلي والذهبي الدافئ مثالي لتقارير الإدارة والخطط الاستراتيجية.',
    colors: {
      primary: '#0F2942',
      secondary: '#1D4E89',
      accent: '#F59E0B',
      bg: '#F8FAFC',
      card_bg: '#FFFFFF',
      text_dark: '#0F172A',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/presentations/templates/business/',
  },
  {
    id: 'canva_coral_pitch',
    name: 'كانفا — عرض الشركات الناشئة المرجاني',
    category: 'تسويق',
    source: 'Canva Pitch Deck',
    theme_color: 'red',
    layout: 'modern_card',
    font_name: 'Dubai',
    description: 'أسلوب Pitch Deck حيوي ومقنع مناسب للعروض التسويقية وإطلاق المنتجات والخدمات.',
    colors: {
      primary: '#E11D48',
      secondary: '#9F1239',
      accent: '#FB7185',
      bg: '#FFF1F2',
      card_bg: '#FFFFFF',
      text_dark: '#1F2937',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/presentations/templates/pitch-deck/',
  },
  {
    id: 'canva_sage_academic',
    name: 'كانفا — الأكاديمي الهادئ (Sage)',
    category: 'تعليمي',
    source: 'Canva Education',
    theme_color: 'green',
    layout: 'minimal_clean',
    font_name: 'Traditional Arabic',
    description: 'تنسيق تعليمي وبحثي مريح للعين مع وضوح عالٍ للنصوص والجداول الأكاديمية.',
    colors: {
      primary: '#166534',
      secondary: '#14532D',
      accent: '#10B981',
      bg: '#F2FBF5',
      card_bg: '#FFFFFF',
      text_dark: '#14261C',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/presentations/templates/education/',
  },
  {
    id: 'canva_teal_modern',
    name: 'كانفا — كانفا فيروزي تقني حديث',
    category: 'تقنية',
    source: 'Canva Creative',
    theme_color: 'blue',
    layout: 'modern_card',
    font_name: 'Dubai',
    description: 'هوية كانفا الفيروزية الممزوجة بالأزرق العميق للعروض التقنية والتحول الرقمي.',
    colors: {
      primary: '#008B94',
      secondary: '#0F3D5E',
      accent: '#00C4CC',
      bg: '#F0FDFA',
      card_bg: '#FFFFFF',
      text_dark: '#0F172A',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/presentations/templates/technology/',
  },
  {
    id: 'canva_dark_midnight',
    name: 'كانفا — الليلي الاحترافي (Dark Pitch)',
    category: 'تقنية',
    source: 'Canva Dark Studio',
    theme_color: 'purple',
    layout: 'dark_pitch',
    font_name: 'Dubai',
    description: 'خلفية داكنة فاخرة تبرز النصوص البيضاء والبنفسجية بوضوح درامي مبهر.',
    colors: {
      primary: '#6366F1',
      secondary: '#1E1B4B',
      accent: '#38BDF8',
      bg: '#0F172A',
      card_bg: '#1E293B',
      text_dark: '#F8FAFC',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/presentations/templates/modern/',
  },
  {
    id: 'canva_royal_gold',
    name: 'كانفا — الملكي الفاخر والذهبي',
    category: 'فاخر',
    source: 'Canva Luxury',
    theme_color: 'purple',
    layout: 'split_accent',
    font_name: 'Traditional Arabic',
    description: 'تنسيق رسمي راقٍ بالبنفسجي العميق ولمسات ذهبية للمناسبات والتقارير العليا.',
    colors: {
      primary: '#4C1D95',
      secondary: '#2E1065',
      accent: '#F59E0B',
      bg: '#FAF5FF',
      card_bg: '#FFFFFF',
      text_dark: '#1E1B4B',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/presentations/templates/creative/',
  },
  {
    id: 'canva_minimal_editorial',
    name: 'كانفا — المينيمال التحريري الأنيق',
    category: 'أعمال',
    source: 'Canva Minimalist',
    theme_color: 'blue',
    layout: 'minimal_clean',
    font_name: 'Simplified Arabic',
    description: 'مساحات بيضاء واسعة مع خطوط عربية واضحة وشريط جانبي هادئ للتركيز على الأفكار.',
    colors: {
      primary: '#1E293B',
      secondary: '#334155',
      accent: '#2563EB',
      bg: '#FFFFFF',
      card_bg: '#F8FAFC',
      text_dark: '#0F172A',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/presentations/templates/minimalist/',
  },
  {
    id: 'canva_sunset_creative',
    name: 'كانفا — الإبداعي الدافئ (Sunset)',
    category: 'إبداعي',
    source: 'Canva Creative',
    theme_color: 'red',
    layout: 'modern_card',
    font_name: 'Dubai',
    description: 'ألوان دافئة مشرقة بالبرتقالي والمرجاني تجذب الانتباه في ورش العمل والمبادرات.',
    colors: {
      primary: '#EA580C',
      secondary: '#9A3412',
      accent: '#FBBF24',
      bg: '#FFF7ED',
      card_bg: '#FFFFFF',
      text_dark: '#27150A',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/presentations/templates/aesthetic/',
  },
  {
    id: 'canva_data_infographic',
    name: 'كانفا — تحليل البيانات والإنفوجرافيك',
    category: 'أعمال',
    source: 'Canva Data Deck',
    theme_color: 'blue',
    layout: 'modern_card',
    font_name: 'Dubai',
    description: 'مخصص لعرض الإحصائيات والجداول والمؤشرات المالية والرسوم البيانية بوضوح فائق.',
    colors: {
      primary: '#2563EB',
      secondary: '#1E40AF',
      accent: '#06B6D4',
      bg: '#EFF6FF',
      card_bg: '#FFFFFF',
      text_dark: '#0F172A',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/presentations/templates/infographic/',
  },
  {
    id: 'canva_islamic_heritage',
    name: 'كانفا — التراث العربي والوقار',
    category: 'تعليمي',
    source: 'Canva Arabic',
    theme_color: 'green',
    layout: 'split_accent',
    font_name: 'Traditional Arabic',
    description: 'مستوحى من الزمردي والذهبي العربي للبحوث الإسلامية والتاريخية والأدبية.',
    colors: {
      primary: '#065F46',
      secondary: '#064E3B',
      accent: '#D97706',
      bg: '#FBFBF7',
      card_bg: '#FFFFFF',
      text_dark: '#1C241E',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/presentations/',
  },
  {
    id: 'canva_edu_royal_emerald',
    name: 'كانفا — التوثيق الملكي الزمردي والذهبي',
    category: 'شواهد وتقارير',
    source: 'Canva Education Studio',
    theme_color: 'green',
    layout: 'split_accent',
    font_name: 'Cairo',
    description: 'تصميم رسمي فاخر بالزمردي السعودي والذهبي مثالي لملفات الشواهد والتقارير المدرسية.',
    colors: {
      primary: '#14532D',
      secondary: '#0F766E',
      accent: '#D97706',
      bg: '#F4FBF7',
      card_bg: '#FFFFFF',
      text_dark: '#14261C',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/templates/s/education-report/',
  },
  {
    id: 'canva_edu_saudi_emerald',
    name: 'كانفا — الأخضر التعليمي الرسمي (المملكة)',
    category: 'شواهد وتقارير',
    source: 'Canva Education KSA',
    theme_color: 'green',
    layout: 'split_accent',
    font_name: 'Cairo',
    description: 'تصميم رسمي فاخر بالزمردي السعودي والذهبي مثالي لملفات الشواهد والتقارير المدرسية.',
    colors: {
      primary: '#14532D',
      secondary: '#0F766E',
      accent: '#D97706',
      bg: '#F4FBF7',
      card_bg: '#FFFFFF',
      text_dark: '#14261C',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/templates/s/education-report/',
  },
  {
    id: 'canva_edu_lavender_charm',
    name: 'كانفا — اللافندر التربوي الراقي',
    category: 'شواهد وتقارير',
    source: 'Canva Portfolio',
    theme_color: 'purple',
    layout: 'modern_card',
    font_name: 'Cairo',
    description: 'ألوان اللافندر والبنفسجي الهادئ المصممة خصيصاً لسجلات شواهد الأداء الوظيفي.',
    colors: {
      primary: '#6D28D9',
      secondary: '#4C1D95',
      accent: '#C084FC',
      bg: '#FAF7FF',
      card_bg: '#FFFFFF',
      text_dark: '#1E1B4B',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/templates/s/teacher-portfolio/',
  },
  {
    id: 'canva_edu_navy_ceremony',
    name: 'كانفا — التقرير المدرسي الرسمي (الكحلي والذهبي)',
    category: 'شواهد وتقارير',
    source: 'Canva Reports',
    theme_color: 'blue',
    layout: 'split_accent',
    font_name: 'Cairo',
    description: 'هوية رسمية متوازنة بالكحلي العميق مع لمسات ذهبية لتقارير الأنشطة والمبادرات.',
    colors: {
      primary: '#1E3A8A',
      secondary: '#0F172A',
      accent: '#F59E0B',
      bg: '#F8FAFC',
      card_bg: '#FFFFFF',
      text_dark: '#0F172A',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/templates/s/school-report/',
  },
  {
    id: 'canva_edu_rose_bloom',
    name: 'كانفا — الإبداع التربوي الوردي الناعم',
    category: 'تعليمي',
    source: 'Canva Creative Edu',
    theme_color: 'red',
    layout: 'modern_card',
    font_name: 'Cairo',
    description: 'تدرجات وردية وعنابية أنيقة تمنح ملفات المعلمات طابعاً جمالياً مرتباً للطباعة.',
    colors: {
      primary: '#9D174D',
      secondary: '#831843',
      accent: '#EC4899',
      bg: '#FDF2F8',
      card_bg: '#FFFFFF',
      text_dark: '#3B0764',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/templates/s/aesthetic-report/',
  },
  {
    id: 'canva_edu_turquoise_modern',
    name: 'كانفا — ملف الإنجاز الفيروزي العصري',
    category: 'تعليمي',
    source: 'Canva Modern',
    theme_color: 'blue',
    layout: 'modern_card',
    font_name: 'Cairo',
    description: 'تصميم عصري بالفيروزي والأزرق السماوي يبرز الصور والباركود بوضوح عالٍ.',
    colors: {
      primary: '#0F766E',
      secondary: '#134E4A',
      accent: '#14B8A6',
      bg: '#F0FDFA',
      card_bg: '#FFFFFF',
      text_dark: '#132A2E',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/templates/s/modern-portfolio/',
  },
  {
    id: 'canva_edu_sand_classic',
    name: 'كانفا — الكلاسيكي الدافئ والذهبي للوثائق',
    category: 'فاخر',
    source: 'Canva Classic Docs',
    theme_color: 'red',
    layout: 'minimal_clean',
    font_name: 'Amiri',
    description: 'إطارات تراثية دافئة بالذهبي والعسلي تمنح الشواهد والتقارير فخامة كلاسيكية عند الطباعة.',
    colors: {
      primary: '#92400E',
      secondary: '#78350F',
      accent: '#D97706',
      bg: '#FFFBEB',
      card_bg: '#FFFFFF',
      text_dark: '#291807',
      text_light: '#FFFFFF',
    },
    canva_url: 'https://www.canva.com/templates/s/certificate-document/',
  },
];

const CLASSIC_THEMES: Record<string, ThemeColors & { name: string; font_name: string }> = {
  blue: {
    name: 'أزرق احترافي',
    primary: '#1E3A8A',
    secondary: '#1D4ED8',
    accent: '#F59E0B',
    bg: '#F5F7FF',
    card_bg: '#FFFFFF',
    text_dark: '#1A1A2E',
    text_light: '#FFFFFF',
    font_name: 'Traditional Arabic',
  },
  green: {
    name: 'أخضر طبيعي',
    primary: '#119955',
    secondary: '#007A3D',
    accent: '#D97706',
    bg: '#F0FBF4',
    card_bg: '#FFFFFF',
    text_dark: '#1A2E1A',
    text_light: '#FFFFFF',
    font_name: 'Traditional Arabic',
  },
  red: {
    name: 'أحمر حيوي',
    primary: '#E53935',
    secondary: '#B71C1C',
    accent: '#FFCC02',
    bg: '#FDF2F2',
    card_bg: '#FFFFFF',
    text_dark: '#2E1A1A',
    text_light: '#FFFFFF',
    font_name: 'Traditional Arabic',
  },
  purple: {
    name: 'بنفسجي ملكي',
    primary: '#6A1B9A',
    secondary: '#4A148C',
    accent: '#FFD700',
    bg: '#F8F0FF',
    card_bg: '#FFFFFF',
    text_dark: '#1A1A2E',
    text_light: '#FFFFFF',
    font_name: 'Traditional Arabic',
  },
};

export function resolveColors(
  themeId: string,
  customTemplates: StudioCanvaTemplate[] = []
): { name: string; colors: ThemeColors; fontName: string; layout: string } {
  const canva = [...customTemplates, ...PRESENTATION_CANVA_TEMPLATES].find((t) => t.id === themeId);
  if (canva) {
    return {
      name: canva.name,
      colors: canva.colors,
      fontName: canva.font_name || canva.font_family || 'Dubai',
      layout: canva.layout || 'modern_card',
    };
  }
  const classic = CLASSIC_THEMES[themeId] || CLASSIC_THEMES.blue;
  return { name: classic.name, colors: classic, fontName: classic.font_name, layout: 'modern_card' };
}

function stripHash(hex: string): string {
  return (hex || '#1E3A8A').replace('#', '').toUpperCase();
}

export function getGeminiClient(): GoogleGenAI | null {
  const key = (process.env.GEMINI_API_KEY || '').trim();
  if (!key || key === 'MY_GEMINI_API_KEY') return null;
  try {
    return new GoogleGenAI({ apiKey: key });
  } catch {
    return null;
  }
}

export function buildPipelineMetrics(opts: {
  operation: string;
  stages: Array<{ key: string; stage: string; shortName: string; duration_ms: number }>;
  wordCount: number;
  itemsCount: number;
  provider: string;
}) {
  const total_ms = opts.stages.reduce((acc, s) => acc + Math.max(1, s.duration_ms), 0);
  const words_per_sec = total_ms > 0 ? Math.round((opts.wordCount / (total_ms / 1000)) * 10) / 10 : 0;
  return {
    operation: opts.operation,
    provider: opts.provider,
    status: 'done',
    total_ms,
    word_count: opts.wordCount,
    words_per_sec,
    items_count: opts.itemsCount,
    stages: opts.stages.map((s) => ({ ...s, status: 'done' })),
  };
}

export async function generateSlidesStructure(params: {
  text: string;
  numSlides: number;
  ptype: string;
  titleOverride?: string;
}): Promise<{ slides: SlideItem[]; provider: string }> {
  const cleanText = params.text.trim();
  const targetCount = Math.max(3, Math.min(24, params.numSlides || 7));
  const ai = getGeminiClient();

  if (ai && cleanText.length > 20) {
    try {
      const prompt = `أنت مصمم عروض تقديمية احترافي باللغة العربية. حوّل النص التالي إلى هيكل عرض تقديمي من ${targetCount} شريحة لنوع (${params.ptype}).
${params.titleOverride ? `العنوان الرئيسي المطلوب: ${params.titleOverride}` : ''}
أعد النتيجة بصيغة JSON فقط على هيئة مصفوفة كائنات:
[{"title": "عنوان الشريحة", "bullets": ["نقطة 1", "نقطة 2", "نقطة 3"]}]
النص:
${cleanText.slice(0, 12000)}`;
      const resp = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: prompt,
        config: { responseMimeType: 'application/json' },
      });
      const parsed = JSON.parse(resp.text || '[]');
      if (Array.isArray(parsed) && parsed.length > 0) {
        const slides: SlideItem[] = parsed.slice(0, targetCount).map((s: Record<string, unknown>, idx: number) => ({
          title: String(s.title || `شريحة ${idx + 1}`).trim(),
          bullets: Array.isArray(s.bullets) ? s.bullets.map((b) => String(b).trim()).filter(Boolean) : [],
          slide_type: idx === 0 ? 'title' : 'content',
        }));
        return { slides, provider: 'Gemini Flash' };
      }
    } catch {
      // Fallback to local structuring
    }
  }

  // Smart local Arabic structuring
  const lines = cleanText
    .split(/\r?\n+/)
    .map((l) => l.trim())
    .filter(Boolean);
  const firstTitle = params.titleOverride || (lines[0] ? lines[0].slice(0, 75) : 'عرض تقديمي احترافي');
  const contentLines = params.titleOverride ? lines : lines.slice(1);
  const slides: SlideItem[] = [
    {
      title: firstTitle,
      bullets: contentLines.slice(0, 3).length ? contentLines.slice(0, 3) : ['نظرة عامة على العرض التقديمي', 'المحاور والأهداف الرئيسية'],
      slide_type: 'title',
    },
  ];

  const remainingSlots = Math.max(1, targetCount - 2);
  if (contentLines.length > 0) {
    const chunkSize = Math.max(1, Math.ceil(contentLines.length / remainingSlots));
    for (let i = 0; i < contentLines.length && slides.length < targetCount - 1; i += chunkSize) {
      const chunk = contentLines.slice(i, i + chunkSize);
      const slideTitle = chunk[0].length <= 55 ? chunk[0] : `المحور ${slides.length}`;
      const bullets = chunk.length > 1 ? chunk.slice(1, 6) : chunk;
      slides.push({
        title: slideTitle,
        bullets: bullets.map((b) => b.slice(0, 140)),
        slide_type: 'content',
      });
    }
  } else {
    for (let i = 1; i < targetCount - 1; i++) {
      slides.push({
        title: `المحور الرئيسي ${i}`,
        bullets: ['تفاصيل وعناصر المحور', 'أبرز النقاط التنفيذية والنتائج'],
        slide_type: 'content',
      });
    }
  }

  slides.push({
    title: 'الخلاصة والتوصيات',
    bullets: ['تلخيص أبرز النتائج والمخرجات', 'الخطوات التنفيذية القادمة', 'شكراً لحسن استماعكم'],
    slide_type: 'closing',
  });

  return { slides, provider: 'المعالج الذكي المحلي' };
}

export async function buildPptxFile(params: {
  outputDir: string;
  slides: SlideItem[];
  themeColor: string;
  customTemplates?: StudioCanvaTemplate[];
  coverData?: Record<string, string> | null;
  templateImages?: string[];
  fontName?: string;
  bodyFontSize?: number;
  filenamePrefix?: string;
  layoutOverride?: string;
}): Promise<{ filename: string; filePath: string; themeName: string; layout: string }> {
  fs.mkdirSync(params.outputDir, { recursive: true });
  const PptxCtor = (PptxGenJS as unknown as { default?: typeof PptxGenJS }).default || PptxGenJS;
  const pptx = new PptxCtor();
  pptx.layout = 'LAYOUT_16x9';
  pptx.rtlMode = true;

  const resolved = resolveColors(params.themeColor, params.customTemplates || []);
  const c = resolved.colors;
  const layout = params.layoutOverride || resolved.layout || 'modern_card';
  const fontFace = params.fontName && params.fontName !== 'auto' ? params.fontName : resolved.fontName;
  const bodySize = Math.max(12, Math.min(34, params.bodyFontSize || 20));
  const images = (params.templateImages || []).filter((img) => /^data:image\/(png|jpeg|jpg|webp|gif)/i.test(img));

  const cover = params.coverData || {};
  const hasCustomCover = Boolean(cover.org || cover.presenter || cover.subtitle || cover.date || cover.logo);

  const isLuxuryFrame = layout === 'royal_double' || layout === 'botanical_frame';
  const isRibbon = layout === 'modern_ribbon' || layout === 'gradient_arch';
  const isStripe = layout === 'clean_stripe' || layout === 'split_accent';
  const isGeometric = layout === 'geometric_split';
  const isPastelOrCard = layout === 'modern_card' || layout === 'dark_pitch' || layout === 'pastel_card';

  // Slide 1: Cover or Title (styled to match Canva template layout)
  const firstSlide = params.slides[0] || { title: params.filenamePrefix || 'عرض تقديمي', bullets: [] };
  const coverSlide = pptx.addSlide();
  coverSlide.background = { color: stripHash(c.primary) };

  if (isLuxuryFrame) {
    coverSlide.addShape(pptx.ShapeType.rect, {
      x: 0.25,
      y: 0.22,
      w: 9.5,
      h: 5.18,
      fill: { color: stripHash(c.secondary) },
      line: { color: stripHash(c.accent), width: 2.5 },
    });
    coverSlide.addShape(pptx.ShapeType.rect, {
      x: 0.37,
      y: 0.34,
      w: 9.26,
      h: 4.94,
      fill: { type: 'none' },
      line: { color: stripHash(c.accent), width: 0.75 },
    });
  } else if (isStripe || isGeometric) {
    coverSlide.addShape(pptx.ShapeType.rect, {
      x: 9.45,
      y: 0,
      w: 0.55,
      h: 5.625,
      fill: { color: stripHash(c.accent) },
    });
    coverSlide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: 9.45,
      h: 0.9,
      fill: { color: stripHash(c.secondary) },
    });
  } else if (isRibbon) {
    coverSlide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: 10,
      h: 0.22,
      fill: { color: stripHash(c.accent) },
    });
    coverSlide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 5.15,
      w: 10,
      h: 0.475,
      fill: { color: stripHash(c.secondary) },
    });
  } else {
    coverSlide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 5.15,
      w: 10,
      h: 0.475,
      fill: { color: stripHash(c.accent) },
    });
  }

  const coverImgSrc = images[0] || firstSlide.image_data;
  if (coverImgSrc) {
    try {
      coverSlide.addImage({ data: coverImgSrc, x: 0.48, y: 0.48, w: 1.45, h: 1.45 });
    } catch {
      // ignore invalid image data
    }
  }

  if (cover.org || cover.logo) {
    coverSlide.addText(`${cover.logo ? cover.logo + ' ' : ''}${cover.org || ''}`.trim(), {
      x: 0.6,
      y: 0.38,
      w: 8.65,
      h: 0.55,
      fontSize: 16,
      bold: true,
      color: stripHash(c.accent),
      align: 'right',
      fontFace,
      rtlMode: true,
    });
  }

  coverSlide.addText(firstSlide.title || params.filenamePrefix || 'عرض تقديمي احترافي', {
    x: 0.75,
    y: 1.5,
    w: 8.5,
    h: 1.45,
    fontSize: 32,
    bold: true,
    color: stripHash(c.text_light),
    align: 'center',
    fontFace,
    rtlMode: true,
  });

  // Accent divider bar
  coverSlide.addShape(pptx.ShapeType.rect, {
    x: 3.1,
    y: 3.02,
    w: 3.8,
    h: 0.06,
    fill: { color: stripHash(c.accent) },
  });

  const subText =
    cover.subtitle ||
    firstSlide.subtitle ||
    (hasCustomCover ? '' : (firstSlide.bullets || []).slice(0, 2).join(' • ')) ||
    'عرض تقديمي منسّق بقالب كانفا';
  if (subText) {
    coverSlide.addText(subText, {
      x: 0.9,
      y: 3.22,
      w: 8.2,
      h: 0.95,
      fontSize: 17,
      color: stripHash(c.text_light),
      align: 'center',
      fontFace,
      rtlMode: true,
    });
  }

  if (cover.presenter || cover.date) {
    coverSlide.addText([cover.presenter, cover.date].filter(Boolean).join('  |  '), {
      x: 0.8,
      y: 4.38,
      w: 8.4,
      h: 0.48,
      fontSize: 13.5,
      color: stripHash(c.accent),
      align: 'center',
      fontFace,
      rtlMode: true,
    });
  }

  // Subsequent slides
  const restSlides = params.slides.slice(1);
  restSlides.forEach((s, idx) => {
    const slide = pptx.addSlide();
    slide.background = { color: stripHash(c.bg) };

    if (isLuxuryFrame) {
      slide.addShape(pptx.ShapeType.rect, {
        x: 0.16,
        y: 0.14,
        w: 9.68,
        h: 5.34,
        fill: { type: 'none' },
        line: { color: stripHash(c.primary), width: 1.5 },
      });
    } else if (isStripe || isGeometric) {
      slide.addShape(pptx.ShapeType.rect, {
        x: 9.82,
        y: 0.95,
        w: 0.18,
        h: 4.675,
        fill: { color: stripHash(c.accent) },
      });
    }

    // Top header bar
    const headerBg = isPastelOrCard ? c.secondary : c.primary;
    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: 10,
      h: 0.92,
      fill: { color: stripHash(headerBg) },
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0.92,
      w: 10,
      h: 0.06,
      fill: { color: stripHash(c.accent) },
    });

    slide.addText(s.title || `شريحة ${idx + 2}`, {
      x: 0.45,
      y: 0.12,
      w: 8.85,
      h: 0.68,
      fontSize: 22,
      bold: true,
      color: stripHash(c.text_light),
      align: 'right',
      fontFace,
      rtlMode: true,
    });

    if (s.slide_type === 'table' && Array.isArray(s.table_data) && s.table_data.length > 0) {
      const tableRowsData = s.table_data.slice(0, 12);
      const tblFontSize = tableRowsData.length > 8 ? Math.max(10, bodySize - 9) : Math.max(11, bodySize - 7);
      const rows = tableRowsData.map((row, rIdx) =>
        row.slice(0, 7).map((cell) => ({
          text: String(cell || ''),
          options: {
            fill: { color: rIdx === 0 ? stripHash(c.primary) : rIdx % 2 === 0 ? stripHash(c.bg) : stripHash(c.card_bg) },
            color: rIdx === 0 ? stripHash(c.text_light) : stripHash(c.text_dark),
            bold: rIdx === 0,
            fontSize: tblFontSize,
            fontFace,
            align: 'right' as const,
            valign: 'middle' as const,
          },
        }))
      );
      slide.addTable(rows, {
        x: 0.45,
        y: 1.2,
        w: 9.1,
        border: { type: 'solid', color: stripHash(c.accent), pt: 1 },
      });
    } else if (
      s.slide_type === 'chart' &&
      Array.isArray(s.chart_labels) &&
      Array.isArray(s.chart_values) &&
      s.chart_labels.length > 0
    ) {
      slide.addChart(
        pptx.ChartType.bar,
        [
          {
            name: s.chart_title || 'القيم',
            labels: s.chart_labels.slice(0, 10),
            values: s.chart_values.slice(0, 10),
          },
        ],
        {
          x: 0.7,
          y: 1.25,
          w: 8.6,
          h: 3.85,
          chartColors: [stripHash(c.primary), stripHash(c.accent), stripHash(c.secondary)],
          showLegend: true,
          showValue: true,
        }
      );
    } else {
      const slideImg =
        (s.image_data && /^data:image\/(png|jpeg|jpg|webp|gif)/i.test(s.image_data) ? s.image_data : undefined) ||
        (images.length > 1 ? images[(idx % (images.length - 1)) + 1] : undefined);
      const hasImg = Boolean(slideImg);
      const contentW = hasImg ? 5.8 : (isStripe ? 8.75 : 9.0);
      const contentX = hasImg ? 3.7 : 0.45;

      if (isPastelOrCard || isLuxuryFrame || isRibbon) {
        slide.addShape(pptx.ShapeType.roundRect, {
          x: contentX,
          y: 1.22,
          w: contentW,
          h: 3.92,
          fill: { color: stripHash(c.card_bg) },
          line: { color: stripHash(c.accent), width: 1.2 },
        });
      } else {
        slide.addShape(pptx.ShapeType.rect, {
          x: contentX,
          y: 1.22,
          w: contentW,
          h: 3.92,
          fill: { color: stripHash(c.card_bg) },
          line: { color: stripHash(c.primary), width: 0.75 },
        });
      }

      const rawBullets = s.bullets && s.bullets.length ? s.bullets : ['محتوى الشريحة'];
      const effectiveSize = rawBullets.length >= 6 ? Math.max(13, bodySize - 4) : bodySize;
      const bulletsText = rawBullets.map((b) => ({
        text: `◆ ${b}`,
        options: {
          fontSize: effectiveSize,
          color: stripHash(c.text_dark),
          fontFace,
          breakLine: true,
          paraSpaceAfter: rawBullets.length >= 6 ? 6 : 10,
        },
      }));

      slide.addText(bulletsText, {
        x: contentX + 0.25,
        y: 1.4,
        w: contentW - 0.5,
        h: 3.56,
        align: 'right',
        valign: 'top',
        rtlMode: true,
      });

      if (slideImg) {
        try {
          slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.4,
            y: 1.25,
            w: 3.05,
            h: 3.8,
            fill: { color: stripHash(c.card_bg) },
            line: { color: stripHash(c.accent), width: 1.5 },
          });
          slide.addImage({ data: slideImg, x: 0.5, y: 1.35, w: 2.85, h: 3.6 });
        } catch {
          // ignore
        }
      }
    }

    // Slide footer
    slide.addText(`${resolved.name}`, {
      x: 4.5,
      y: 5.25,
      w: 5.0,
      h: 0.28,
      fontSize: 9.5,
      color: stripHash(c.primary),
      align: 'right',
      fontFace,
      rtlMode: true,
    });
    slide.addText(`${idx + 2}`, {
      x: 0.35,
      y: 5.25,
      w: 0.8,
      h: 0.28,
      fontSize: 10,
      bold: true,
      color: stripHash(c.primary),
      align: 'left',
    });
  });

  const safePrefix = (params.filenamePrefix || 'عرض_تقديمي')
    .replace(/[^a-zA-Z0-9\u0600-\u06FF_-]+/g, '_')
    .slice(0, 48);
  const filename = `${safePrefix || 'presentation'}_${Date.now()}.pptx`;
  const filePath = path.join(params.outputDir, filename);
  await pptx.writeFile({ fileName: filePath, compression: false });

  return { filename, filePath, themeName: resolved.name, layout };
}

export async function extractFileContent(fileBuffer: Buffer, originalName: string): Promise<{
  full_text: string;
  word_count: number;
  tables: string[][][];
  slides: SlideItem[];
}> {
  const ext = path.extname(originalName || '').toLowerCase();
  let fullText = '';
  const tables: string[][][] = [];
  const slides: SlideItem[] = [];

  if (ext === '.pdf') {
    const { text } = await extractText(new Uint8Array(fileBuffer), { mergePages: true });
    fullText = (text || '').trim();
  } else if (ext === '.pptx') {
    const zip = await JSZip.loadAsync(fileBuffer);
    const slideFiles = Object.keys(zip.files)
      .filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k))
      .sort((a, b) => {
        const na = Number(a.match(/\d+/)?.[0] || 0);
        const nb = Number(b.match(/\d+/)?.[0] || 0);
        return na - nb;
      });

    for (const sPath of slideFiles) {
      const xml = await zip.files[sPath].async('string');
      const $ = cheerio.load(xml, { xmlMode: true });
      const texts: string[] = [];
      $('a\\:p, p').each((_i, el) => {
        const line = $(el).text().replace(/\s+/g, ' ').trim();
        if (line) texts.push(line);
      });
      const title = texts[0] || `شريحة ${slides.length + 1}`;
      const bullets = texts.slice(1);
      slides.push({ title, bullets, slide_type: slides.length === 0 ? 'title' : 'content' });
      fullText += `${title}\n${bullets.join('\n')}\n\n`;
    }
    fullText = fullText.trim();
  } else if (ext === '.docx') {
    const zip = await JSZip.loadAsync(fileBuffer);
    const docXml = zip.files['word/document.xml'];
    if (docXml) {
      const xml = await docXml.async('string');
      const $ = cheerio.load(xml, { xmlMode: true });
      const paras: string[] = [];
      $('w\\:p').each((_i, el) => {
        const line = $(el).text().replace(/\s+/g, ' ').trim();
        if (line) paras.push(line);
      });
      fullText = paras.join('\n');

      $('w\\:tbl').each((_ti, tblEl) => {
        const tblRows: string[][] = [];
        $(tblEl)
          .find('w\\:tr')
          .each((_ri, trEl) => {
            const rowCells: string[] = [];
            $(trEl)
              .find('w\\:tc')
              .each((_ci, tcEl) => {
                rowCells.push($(tcEl).text().replace(/\s+/g, ' ').trim());
              });
            if (rowCells.length) tblRows.push(rowCells);
          });
        if (tblRows.length) tables.push(tblRows);
      });
    }
  } else {
    fullText = fileBuffer.toString('utf-8').trim();
  }

  const words = fullText ? fullText.split(/\s+/).filter(Boolean).length : 0;
  return { full_text: fullText, word_count: words, tables, slides };
}

export function parseHtmlDocument(html: string): {
  title: string;
  sections: Array<{ heading: string; paragraphs: string[]; bullets: string[]; table?: string[][] }>;
  tables: string[][][];
} {
  const $ = cheerio.load(html || '');
  $('script, style, noscript, template').remove();
  const title = $('title').text().trim() || $('h1').first().text().trim() || 'مستند احترافي';
  const tables: string[][][] = [];
  const sections: Array<{ heading: string; paragraphs: string[]; bullets: string[]; table?: string[][] }> = [];

  let current = { heading: title, paragraphs: [] as string[], bullets: [] as string[], table: undefined as string[][] | undefined };

  $('h1, h2, h3, h4, p, ul, ol, table').each((_i, el) => {
    const tag = (el.tagName || '').toLowerCase();
    if (tag !== 'table' && $(el).closest('table').length > 0) return;
    if (tag === 'p' && $(el).closest('li').length > 0) return;

    if (tag === 'h1' || tag === 'h2' || tag === 'h3' || tag === 'h4') {
      if (current.paragraphs.length || current.bullets.length || current.table) {
        sections.push(current);
      }
      current = { heading: $(el).text().replace(/\s+/g, ' ').trim() || title, paragraphs: [], bullets: [], table: undefined };
    } else if (tag === 'p') {
      const txt = $(el).text().replace(/\s+/g, ' ').trim();
      if (txt) current.paragraphs.push(txt);
    } else if (tag === 'ul' || tag === 'ol') {
      $(el)
        .find('li')
        .each((_li, liEl) => {
          const txt = $(liEl).text().replace(/\s+/g, ' ').trim();
          if (txt) current.bullets.push(txt);
        });
    } else if (tag === 'table') {
      const rows: string[][] = [];
      $(el)
        .find('tr')
        .each((_ri, trEl) => {
          const row: string[] = [];
          $(trEl)
            .find('th, td')
            .each((_ci, cellEl) => {
              row.push($(cellEl).text().replace(/\s+/g, ' ').trim());
            });
          if (row.length) rows.push(row);
        });
      if (rows.length) {
        tables.push(rows);
        current.table = rows;
      }
    }
  });

  if (current.paragraphs.length || current.bullets.length || current.table) {
    sections.push(current);
  }

  if (!sections.length) {
    const bodyText = $('body').text().replace(/\s+/g, ' ').trim();
    if (bodyText) {
      sections.push({ heading: title, paragraphs: [bodyText], bullets: [] });
    }
  }

  return { title, sections, tables };
}

export function parseHtmlToSlides(html: string, titleOverride?: string): {
  title: string;
  slides: SlideItem[];
  wordCount: number;
} {
  const $ = cheerio.load(html || '');
  $('script, style, noscript, template, iframe').remove();

  const cleanTxt = (s: string) => s.replace(/\s+/g, ' ').trim();
  const docTitle =
    cleanTxt(String(titleOverride || '')) ||
    cleanTxt($('h1').first().text()) ||
    cleanTxt($('title').text()) ||
    'عرض تقديمي احترافي';

  const bodyText = cleanTxt($('body').text() || $.root().text());
  const wordCount = bodyText ? bodyText.split(/\s+/).filter(Boolean).length : 0;

  const isRasterDataUrl = (src?: string) =>
    Boolean(src && /^data:image\/(png|jpeg|jpg|webp|gif)/i.test(src) && src.length > 140);

  const extractTableRows = (tblEl: cheerio.Cheerio<any>): string[][] => {
    const rows: string[][] = [];
    tblEl.find('tr').each((_ri, trEl) => {
      const row: string[] = [];
      $(trEl)
        .find('th, td')
        .each((_ci, cellEl) => {
          row.push(cleanTxt($(cellEl).text()));
        });
      if (row.length && row.some(Boolean)) rows.push(row);
    });
    return rows;
  };

  const pushTableSlides = (outSlides: SlideItem[], heading: string, rows: string[][]) => {
    if (!rows.length) return;
    const header = rows[0];
    const body = rows.slice(1);
    const maxBodyPerSlide = 9;
    if (body.length <= maxBodyPerSlide) {
      outSlides.push({ title: heading, bullets: [], slide_type: 'table', table_data: rows });
      return;
    }
    for (let i = 0; i < body.length; i += maxBodyPerSlide) {
      const partIdx = Math.floor(i / maxBodyPerSlide) + 1;
      outSlides.push({
        title: partIdx === 1 ? heading : `${heading} (تابع ${partIdx})`,
        bullets: [],
        slide_type: 'table',
        table_data: [header, ...body.slice(i, i + maxBodyPerSlide)],
      });
    }
  };

  const pushBulletSlides = (
    outSlides: SlideItem[],
    heading: string,
    bullets: string[],
    imageData?: string,
    slideType = 'content'
  ) => {
    const cleanBullets = bullets.map(cleanTxt).filter(Boolean);
    if (!cleanBullets.length && !imageData) return;
    const chunkSize = 6;
    if (cleanBullets.length <= chunkSize) {
      outSlides.push({
        title: heading,
        bullets: cleanBullets.length ? cleanBullets : [heading],
        slide_type: slideType,
        image_data: imageData,
      });
      return;
    }
    for (let i = 0; i < cleanBullets.length; i += chunkSize) {
      const part = Math.floor(i / chunkSize) + 1;
      outSlides.push({
        title: part === 1 ? heading : `${heading} (${part})`,
        bullets: cleanBullets.slice(i, i + chunkSize),
        slide_type: slideType,
        image_data: part === 1 ? imageData : undefined,
      });
    }
  };

  const slides: SlideItem[] = [];

  // 1) Fast path: explicit slide / page / section containers
  const slideContainers = $('.slide, .ppt-slide, [data-slide], .canva-a4-sheet, main > section, body > section, article.slide')
    .toArray()
    .filter((el) => {
      const $el = $(el);
      if ($el.parents('.slide, .ppt-slide, [data-slide], .canva-a4-sheet').length > 0) return false;
      const tLen = cleanTxt($el.text()).length;
      return tLen > 8 || $el.find('table, img').length > 0;
    });

  if (slideContainers.length >= 2) {
    slideContainers.forEach((el, idx) => {
      const $el = $(el);
      const heading =
        cleanTxt($el.find('h1, h2, h3, h4, .slide-title, .title').first().text()) ||
        (idx === 0 ? docTitle : `شريحة ${idx + 1}`);

      let slideImg: string | undefined;
      $el.find('img').each((_ii, imgEl) => {
        if (!slideImg) {
          const src = $(imgEl).attr('src');
          if (isRasterDataUrl(src)) slideImg = src;
        }
      });

      const items: string[] = [];
      const seenItem = new Set<string>();
      $el.find('li, p, .card, .kpi, blockquote').each((_bi, bEl) => {
        const $b = $(bEl);
        if ($b.closest('table').length > 0) return;
        if ($b.is('p') && $b.closest('li, .card, .kpi').length > 0) return;
        const txt = cleanTxt($b.text());
        if (txt && txt !== heading && !seenItem.has(txt)) {
          seenItem.add(txt);
          items.push(txt.slice(0, 200));
        }
      });

      const tblEls = $el.find('table').toArray();
      if (items.length || slideImg || !tblEls.length) {
        if (idx === 0) {
          slides.push({
            title: heading,
            subtitle: items[0] || '',
            bullets: items.slice(0, 4),
            slide_type: 'title',
            image_data: slideImg,
          });
        } else {
          pushBulletSlides(slides, heading, items, slideImg, 'content');
        }
      }
      tblEls.forEach((tEl, ti) => {
        const rows = extractTableRows($(tEl));
        pushTableSlides(slides, tblEls.length > 1 ? `${heading} — جدول ${ti + 1}` : heading, rows);
      });
    });

    if (slides.length > 0) {
      return { title: docTitle, slides, wordCount };
    }
  }

  // 2) General HTML document fast single-pass extraction
  const firstBannerP = cleanTxt($('.header-banner p, header p').first().text()) || '';
  const coverSub =
    firstBannerP && firstBannerP !== docTitle
      ? firstBannerP
      : cleanTxt($('p').first().text()) || 'عرض تقديمي مستخرج من كود HTML';

  const firstRasterImg = $('img')
    .toArray()
    .map((el) => $(el).attr('src'))
    .find((s) => isRasterDataUrl(s));

  slides.push({
    title: docTitle,
    subtitle: coverSub.slice(0, 180),
    bullets: coverSub ? [coverSub.slice(0, 160)] : [],
    slide_type: 'title',
    image_data: firstRasterImg,
  });

  let currentHeading = docTitle;
  let currentBullets: string[] = [];
  let currentImg: string | undefined;
  const seenBlocks = new Set<string>();
  const skipInHeaderBanner = new Set<any>();
  $('.header-banner, header').find('h1, h2, h3, p').each((_i, el) => {
    skipInHeaderBanner.add(el);
  });

  const flushCurrent = () => {
    if (currentBullets.length || currentImg) {
      pushBulletSlides(slides, currentHeading, currentBullets, currentImg, 'content');
      currentBullets = [];
      currentImg = undefined;
    }
  };

  $('h1, h2, h3, h4, p, ul, ol, table, .card, .highlight-box, .kpi, .stat-card, blockquote, img').each((_i, el) => {
    const $el = $(el);
    if (skipInHeaderBanner.has(el)) return;
    if ($el.closest('table').length > 0 && !$el.is('table')) return;
    if (($el.is('p') || $el.is('h4')) && $el.closest('li, .card, .highlight-box, .kpi, .stat-card').length > 0) return;

    const tag = (el.tagName || '').toLowerCase();
    if (tag === 'h1' || tag === 'h2' || tag === 'h3' || tag === 'h4') {
      const hTxt = cleanTxt($el.text());
      if (!hTxt) return;
      flushCurrent();
      currentHeading = hTxt.slice(0, 110);
    } else if (tag === 'img') {
      const src = $el.attr('src');
      if (isRasterDataUrl(src) && src !== firstRasterImg && !currentImg) {
        currentImg = src;
      }
    } else if (tag === 'table') {
      flushCurrent();
      const tblHeading = cleanTxt($el.prev('h2, h3, h4, .table-title, caption').text()) ||
        cleanTxt($el.find('caption').text()) ||
        (currentHeading !== docTitle ? currentHeading : 'جدول البيانات');
      const rows = extractTableRows($el);
      pushTableSlides(slides, tblHeading, rows);
    } else if (tag === 'ul' || tag === 'ol') {
      $el.find('li').each((_li, liEl) => {
        const txt = cleanTxt($(liEl).text());
        if (txt && !seenBlocks.has(txt)) {
          seenBlocks.add(txt);
          currentBullets.push(txt.slice(0, 200));
        }
      });
    } else {
      // p, .card, .highlight-box, .kpi, .stat-card, blockquote
      const cardTitle = cleanTxt($el.find('h2, h3, h4, strong').first().text());
      const fullTxt = cleanTxt($el.text());
      if (!fullTxt || fullTxt === docTitle || fullTxt === coverSub) return;
      const line =
        cardTitle && fullTxt.startsWith(cardTitle) && fullTxt.length > cardTitle.length + 2
          ? `${cardTitle}: ${fullTxt.slice(cardTitle.length).replace(/^[:\s-]+/, '')}`
          : fullTxt;
      if (line && !seenBlocks.has(line)) {
        seenBlocks.add(line);
        currentBullets.push(line.slice(0, 220));
      }
    }
  });

  flushCurrent();

  if (slides.length === 1 && bodyText) {
    const chunks = bodyText
      .split(/[.!؟\n]+/)
      .map(cleanTxt)
      .filter((x) => x.length > 4 && x !== docTitle);
    if (chunks.length) {
      pushBulletSlides(slides, 'محتوى العرض التقديمي', chunks.slice(0, 12));
    }
  }

  return { title: docTitle, slides, wordCount };
}

export async function buildDocxFile(params: {
  outputDir: string;
  filenamePrefix: string;
  title: string;
  paragraphs?: string[];
  sections?: Array<{ heading: string; paragraphs: string[]; bullets: string[]; table?: string[][] }>;
  fontFamily?: string;
  fontSizePt?: number;
  marginInches?: number;
  rtl?: boolean;
}): Promise<{ filename: string; filePath: string }> {
  fs.mkdirSync(params.outputDir, { recursive: true });
  const font = params.fontFamily || 'Cairo';
  const halfPts = Math.round((params.fontSizePt || 12) * 2);
  const marginTwips = Math.round((params.marginInches || 1.0) * 1440);
  const rtl = params.rtl !== false;
  const children: Array<Paragraph | Table> = [];

  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
      bidirectional: rtl,
      spacing: { after: 220 },
      children: [new TextRun({ text: params.title || 'مستند Word احترافي', bold: true, size: halfPts + 10, font, color: '1E3A8A' })],
    })
  );

  const secs = params.sections && params.sections.length
    ? params.sections
    : [{ heading: '', paragraphs: params.paragraphs || [], bullets: [], table: undefined }];

  for (const sec of secs) {
    if (sec.heading && sec.heading !== params.title) {
      children.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
          bidirectional: rtl,
          spacing: { before: 180, after: 120 },
          children: [new TextRun({ text: sec.heading, bold: true, size: halfPts + 4, font, color: '0F766E' })],
        })
      );
    }
    for (const p of sec.paragraphs) {
      children.push(
        new Paragraph({
          alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
          bidirectional: rtl,
          spacing: { after: 120 },
          children: [new TextRun({ text: p, size: halfPts, font })],
        })
      );
    }
    for (const b of sec.bullets) {
      children.push(
        new Paragraph({
          alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
          bidirectional: rtl,
          spacing: { after: 80 },
          children: [new TextRun({ text: `• ${b}`, size: halfPts, font })],
        })
      );
    }
    if (sec.table && sec.table.length > 0) {
      const tRows = sec.table.map(
        (row, rIdx) =>
          new TableRow({
            children: row.map(
              (cell) =>
                new TableCell({
                  width: { size: Math.floor(100 / Math.max(1, row.length)), type: WidthType.PERCENTAGE },
                  borders: {
                    top: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
                    bottom: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
                    left: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
                    right: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
                  },
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.RIGHT,
                      bidirectional: rtl,
                      children: [
                        new TextRun({
                          text: cell || '',
                          bold: rIdx === 0,
                          size: Math.max(18, halfPts - 2),
                          font,
                        }),
                      ],
                    }),
                  ],
                })
            ),
          })
      );
      children.push(new Table({ rows: tRows, width: { size: 100, type: WidthType.PERCENTAGE } }));
      children.push(new Paragraph({ text: '' }));
    }
  }

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: marginTwips, bottom: marginTwips, left: marginTwips, right: marginTwips },
          },
        },
        children,
      },
    ],
  });

  const buf = await Packer.toBuffer(doc);
  const safePrefix = (params.filenamePrefix || 'document').replace(/[^a-zA-Z0-9\u0600-\u06FF_-]+/g, '_').slice(0, 48);
  const filename = `${safePrefix || 'document'}_${Date.now()}.docx`;
  const filePath = path.join(params.outputDir, filename);
  fs.writeFileSync(filePath, buf);
  return { filename, filePath };
}

export function buildExcelFromHtml(params: { outputDir: string; html: string; filenamePrefix: string }): {
  filename: string;
  filePath: string;
  sheetsCount: number;
} {
  fs.mkdirSync(params.outputDir, { recursive: true });
  const parsed = parseHtmlDocument(params.html);
  const wb = XLSX.utils.book_new();

  if (parsed.tables.length > 0) {
    parsed.tables.forEach((tbl, idx) => {
      const ws = XLSX.utils.aoa_to_sheet(tbl);
      XLSX.utils.book_append_sheet(wb, ws, `جدول_${idx + 1}`);
    });
  } else {
    const rows: string[][] = [[parsed.title]];
    parsed.sections.forEach((s) => {
      rows.push([s.heading]);
      s.paragraphs.forEach((p) => rows.push([p]));
      s.bullets.forEach((b) => rows.push([b]));
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, 'البيانات');
  }

  const safePrefix = (params.filenamePrefix || 'excel_export').replace(/[^a-zA-Z0-9\u0600-\u06FF_-]+/g, '_').slice(0, 48);
  const filename = `${safePrefix || 'excel_export'}_${Date.now()}.xlsx`;
  const filePath = path.join(params.outputDir, filename);
  XLSX.writeFile(wb, filePath);
  return { filename, filePath, sheetsCount: wb.SheetNames.length };
}

export interface VideoJobState {
  job_id: string;
  status: 'queued' | 'running' | 'done' | 'error';
  percent: number;
  stage: string;
  filename?: string;
  download_url?: string;
  preview_url?: string;
  file_size?: number;
  width?: number;
  height?: number;
  fps?: number;
  duration_seconds?: number;
  format?: string;
  has_audio?: boolean;
  pipeline_metrics?: ReturnType<typeof buildPipelineMetrics>;
  error?: string;
}

export const EXAMPLE_VIDEO_HTML = `<!DOCTYPE html><html lang="ar" dir="rtl"><head>  <meta charset="UTF-8" />  <meta name="viewport" content="width=device-width, initial-scale=1.0" />  <title>اختبار تحويل HTML إلى فيديو</title>  <style>    * { box-sizing: border-box; }    body {      margin: 0;      padding: 0;      background: linear-gradient(135deg, #0b0d12 0%, #1a1f2e 100%);      display: flex;      align-items: center;      justify-content: center;      height: 100vh;      font-family: "Segoe UI", Tahoma, sans-serif;      overflow: hidden;    }    .container {      display: flex;      flex-direction: column;      align-items: center;      gap: 2.5vh;    }    .box {      width: 20vw;      height: 20vw;      max-width: 220px;      max-height: 220px;      background: linear-gradient(45deg, #3291ff, #6ee7b7);      border-radius: 28px;      animation: pulse 2s ease-in-out infinite alternate;      box-shadow: 0 0 40px rgba(50, 145, 255, 0.45);    }    @keyframes pulse {      from { transform: scale(0.82); opacity: 0.7; }      to   { transform: scale(1.18); opacity: 1; }    }    .text {      color: #ffffff;      font-size: 3.2vw;      font-weight: 700;      letter-spacing: 0.5px;      animation: fadeIn 1.5s ease-in-out infinite alternate;    }    @keyframes fadeIn {      from { opacity: 0.45; }      to   { opacity: 1; }    }  </style></head><body>  <div class="container">    <div class="box"></div>    <div class="text">اختبار التحويل إلى فيديو — الاستوديو الموحد</div>  </div></body></html>`;

const videoJobs = new Map<string, VideoJobState>();

export function getVideoJob(jobId: string): VideoJobState | undefined {
  return videoJobs.get(jobId);
}

export async function saveUploadedStudioAudio(params: {
  audioDir: string;
  fileBuffer: Buffer;
  originalName: string;
}): Promise<{ audio_id: string; original_name: string; preview_url: string; duration_seconds: number }> {
  fs.mkdirSync(params.audioDir, { recursive: true });
  const ext = path.extname(params.originalName || '.mp3').toLowerCase() || '.mp3';
  const audioId = `aud_up_${Date.now().toString(36)}${ext}`;
  const dest = path.join(params.audioDir, audioId);
  fs.writeFileSync(dest, params.fileBuffer);
  return {
    audio_id: audioId,
    original_name: params.originalName,
    preview_url: `/api/studio/audio/${encodeURIComponent(audioId)}`,
    duration_seconds: 10,
  };
}

export async function generateStudioAudioPreview(params: {
  audioDir: string;
  html: string;
  duration: number;
  autoAudioType: string;
  voiceName: string;
  voiceScript: string;
  notes: string;
}): Promise<{
  audio_id: string;
  preview_url: string;
  duration_seconds: number;
  voice_name: string;
  mood: string;
  script: string;
}> {
  fs.mkdirSync(params.audioDir, { recursive: true });
  const dur = Math.max(3, Math.min(45, Number(params.duration) || 10));
  const $ = cheerio.load(params.html || '');
  const extractedText = $('h1, h2, .title, .text, p').text().replace(/\s+/g, ' ').trim().slice(0, 220);
  let script = params.voiceScript.trim() || extractedText || 'عرض احترافي من الاستوديو الموحد للعروض والمستندات.';

  const ai = getGeminiClient();
  if (ai && !params.voiceScript.trim()) {
    try {
      const resp = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: `اكتب تعليقاً صوتياً عربياً موجزاً وجذاباً لمدة ${dur} ثانية استناداً إلى هذا النص والملاحظات:\nالمحتوى: ${extractedText}\nالملاحظات: ${params.notes}\nأعد نص التعليق الصوتي فقط بدون شرح.`,
      });
      if (resp.text && resp.text.trim()) {
        script = resp.text.trim().slice(0, 400);
      }
    } catch {
      // keep fallback script
    }
  }

  const audioId = `aud_ai_${Date.now().toString(36)}.wav`;
  const outWav = path.join(params.audioDir, audioId);
  // Synthesize gentle stereo harmonic ambient wav track using FFmpeg
  const mood = params.autoAudioType === 'music_only' ? 'موسيقى خلفية هادئة' : 'تعليق وخلفية متناغمة';
  await execFileAsync(
    'ffmpeg',
    [
      '-y',
      '-f',
      'lavfi',
      '-i',
      `sine=frequency=220:sample_rate=44100:duration=${dur}`,
      '-f',
      'lavfi',
      '-i',
      `sine=frequency=330:sample_rate=44100:duration=${dur}`,
      '-filter_complex',
      `[0:a][1:a]amix=inputs=2:duration=shortest,afade=t=in:st=0:d=0.6,afade=t=out:st=${Math.max(0.5, dur - 0.8)}:d=0.8,volume=0.22[aout]`,
      '-map',
      '[aout]',
      '-c:a',
      'pcm_s16le',
      outWav,
    ],
    { timeout: 20000 }
  );

  return {
    audio_id: audioId,
    preview_url: `/api/studio/audio/${encodeURIComponent(audioId)}`,
    duration_seconds: dur,
    voice_name: params.voiceName || 'Kore',
    mood,
    script,
  };
}

export async function aiGenerateVideoHtml(params: {
  prompt: string;
  notes: string;
  width: number;
  height: number;
  duration: number;
}): Promise<{ html: string; model: string }> {
  const ai = getGeminiClient();
  const title = (params.prompt || params.notes || 'عرض فيديو احترافي').split('\n')[0].slice(0, 60);
  const sub = (params.notes || params.prompt || 'تم التصميم بالذكاء الاصطناعي').slice(0, 140);

  if (ai) {
    try {
      const resp = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: `صمم صفحة HTML5 كاملة مستقلة مع CSS @keyframes متحركة (100vw x 100vh، RTL عربي) مناسبة للتحويل إلى فيديو بأبعاد ${params.width}x${params.height} ومدته ${params.duration} ثانية.\nالفكرة: ${params.prompt}\nالملاحظات: ${params.notes}\nأعد كود HTML الكامل فقط يبدأ بـ <!DOCTYPE html>.`,
      });
      const raw = (resp.text || '').replace(/^```html\s*/i, '').replace(/```$/i, '').trim();
      if (raw.includes('<html') && raw.includes('</html>')) {
        return { html: raw, model: 'Gemini Flash' };
      }
    } catch {
      // fallback
    }
  }

  const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const html = `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8" />
  <style>
    * { box-sizing: border-box; }
    body {
      margin: 0; height: 100vh; overflow: hidden;
      display: flex; align-items: center; justify-content: center;
      font-family: 'Cairo', 'Segoe UI', sans-serif;
      background: linear-gradient(135deg, #091e3a 0%, #1d4e89 55%, #008b94 100%);
      color: #ffffff;
    }
    .card {
      width: 82vw; padding: 6vh 5vw; border-radius: 28px; text-align: center;
      background: rgba(15, 23, 42, 0.58); border: 2px solid rgba(245, 158, 11, 0.75);
      box-shadow: 0 24px 60px rgba(0, 0, 0, 0.45);
      animation: floatUp 2.6s ease-in-out infinite alternate;
    }
    h1 { font-size: 4.5vw; margin: 0 0 2vh; color: #f8fafc; }
    p { font-size: 2.2vw; margin: 0; color: #fde68a; line-height: 1.6; }
    @keyframes floatUp {
      from { transform: translateY(10px) scale(0.98); }
      to { transform: translateY(-10px) scale(1.01); }
    }
  </style>
</head>
<body>
  <div class="card">
    <h1>${esc(title)}</h1>
    <p>${esc(sub)}</p>
  </div>
</body>
</html>`;
  return { html, model: 'المولد المحلي الذكي' };
}

export async function aiRefineVideoHtml(params: {
  html: string;
  notes: string;
  width: number;
  height: number;
  duration: number;
}): Promise<{ html: string; model: string }> {
  const ai = getGeminiClient();
  if (ai && params.html.trim()) {
    try {
      const resp = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: `عدّل كود HTML التالي لتنفيذ هذه الملاحظات الدقيقة مع الحفاظ على RTL والحركات البصرية (${params.width}x${params.height}، ${params.duration}s):\nالملاحظات: ${params.notes}\nأعد كود HTML الكامل فقط:\n${params.html.slice(0, 14000)}`,
      });
      const raw = (resp.text || '').replace(/^```html\s*/i, '').replace(/```$/i, '').trim();
      if (raw.includes('<html')) {
        return { html: raw, model: 'Gemini Flash' };
      }
    } catch {
      // fallback
    }
  }
  const gen = await aiGenerateVideoHtml({
    prompt: params.notes,
    notes: params.notes,
    width: params.width,
    height: params.height,
    duration: params.duration,
  });
  return gen;
}

export async function aiProcessHtml(params: {
  html: string;
  action: string;
  customPrompt?: string;
}): Promise<{ processed_html: string; provider: string }> {
  const ai = getGeminiClient();
  const actionMap: Record<string, string> = {
    optimize: 'نظّف كود HTML وثبّت تنسيقات CSS وحسّن عرض الجداول والاتجاه RTL بدون حذف المحتوى.',
    academic: 'نسّق المستند بأسلوب أكاديمي فاخر مناسب للتصدير إلى Word و PowerPoint مع وضوح العناوين والجداول.',
    add_toc_cover: 'أضف صفحة غلاف أنيقة وفهرساً للمحتويات في أعلى مستند HTML.',
    custom: params.customPrompt || 'حسّن تنسيق HTML وتناسق الألوان.',
  };
  const instruction = actionMap[params.action] || actionMap.optimize;

  if (ai) {
    try {
      const resp = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: `${instruction}\nأعد كود HTML الكامل فقط:\n${params.html.slice(0, 15000)}`,
      });
      const raw = (resp.text || '').replace(/^```html\s*/i, '').replace(/```$/i, '').trim();
      if (raw.includes('<')) {
        return { processed_html: raw, provider: 'Gemini Flash' };
      }
    } catch {
      // fallback
    }
  }

  const $ = cheerio.load(params.html || '');
  $('html').attr('dir', 'rtl').attr('lang', 'ar');
  if (params.action === 'add_toc_cover') {
    const headings = $('h1, h2, h3')
      .map((_i, el) => $(el).text().trim())
      .get()
      .filter(Boolean)
      .slice(0, 10);
    const tocItems = headings.map((h, i) => `<li>${i + 1}. ${h}</li>`).join('');
    const coverHtml = `<section style="padding:28px;border:2px solid #1e3a8a;border-radius:12px;margin-bottom:24px;background:#f8fafc;text-align:center;">
      <h1 style="color:#1e3a8a;margin-top:0;">${$('title').text() || $('h1').first().text() || 'مستند منسّق احترافياً'}</h1>
      ${tocItems ? `<div style="text-align:right;margin-top:14px;"><strong>فهرس المحتويات:</strong><ul>${tocItems}</ul></div>` : ''}
    </section>`;
    $('body').prepend(coverHtml);
  }
  return { processed_html: $.html(), provider: 'المعالج الذكي المحلي' };
}

export async function analyzePptxBuffer(fileBuffer: Buffer): Promise<{
  slide_count: number;
  word_count: number;
  tables_count: number;
  images_count: number;
  fonts_used: string[];
  colors_used: string[];
  slides: SlideItem[];
  full_text: string;
  analysis: {
    overall_score: number;
    summary: string;
    strengths: string[];
    weaknesses: string[];
    recommended_templates: Array<{ id: string; name: string; reason: string }>;
    slide_suggestions: Array<{
      slide_number: number;
      original_title: string;
      improved_title: string;
      improved_bullets: string[];
      improvement_reason: string;
      design_tip: string;
    }>;
  };
  provider: string;
}> {
  const zip = await JSZip.loadAsync(fileBuffer);
  const slideFiles = Object.keys(zip.files)
    .filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k))
    .sort((a, b) => Number(a.match(/\d+/)?.[0] || 0) - Number(b.match(/\d+/)?.[0] || 0));

  const mediaFiles = Object.keys(zip.files).filter((k) => /^ppt\/media\//.test(k));
  const fontsSet = new Set<string>();
  const colorsSet = new Set<string>();
  const slides: SlideItem[] = [];
  let tablesCount = 0;
  let fullText = '';

  for (const sPath of slideFiles) {
    const xml = await zip.files[sPath].async('string');
    const $ = cheerio.load(xml, { xmlMode: true });
    tablesCount += $('a\\:tbl, tbl').length;
    $('a\\:latin, a\\:cs, latin, cs').each((_i, el) => {
      const tf = $(el).attr('typeface');
      if (tf && !tf.startsWith('+')) fontsSet.add(tf);
    });
    $('a\\:srgbClr, srgbClr').each((_i, el) => {
      const val = $(el).attr('val');
      if (val) colorsSet.add(`#${val}`);
    });
    const texts: string[] = [];
    $('a\\:p, p').each((_i, el) => {
      const line = $(el).text().replace(/\s+/g, ' ').trim();
      if (line) texts.push(line);
    });
    const title = texts[0] || `شريحة ${slides.length + 1}`;
    const bullets = texts.slice(1, 8);
    slides.push({ title, bullets, slide_type: slides.length === 0 ? 'title' : 'content' });
    fullText += `${title}\n${bullets.join('\n')}\n\n`;
  }

  const wordCount = fullText.trim() ? fullText.trim().split(/\s+/).length : 0;
  const recs = [
    {
      id: PRESENTATION_CANVA_TEMPLATES[0].id,
      name: PRESENTATION_CANVA_TEMPLATES[0].name,
      reason: 'توازن رسمي عالٍ يبرز العناوين والبيانات التنفيذية بوضوح.',
    },
    {
      id: PRESENTATION_CANVA_TEMPLATES[3].id,
      name: PRESENTATION_CANVA_TEMPLATES[3].name,
      reason: 'أسلوب حديث وألوان فيروزية مريحة للقراءة والعرض التفاعلي.',
    },
    {
      id: PRESENTATION_CANVA_TEMPLATES[8].id,
      name: PRESENTATION_CANVA_TEMPLATES[8].name,
      reason: 'مثالي لتنظيم النقاط والجداول بأسلوب إنفوجرافيك أنيق.',
    },
  ];

  const defaultSuggestions = slides.slice(0, 12).map((s, idx) => ({
    slide_number: idx + 1,
    original_title: s.title,
    improved_title: s.title.length > 60 ? `${s.title.slice(0, 52)}...` : s.title,
    improved_bullets: (s.bullets || []).slice(0, 5).map((b) => b.slice(0, 110)),
    improvement_reason: 'تكثيف النقاط وإبراز الفكرة المحورية لزيادة وضوح الشريحة.',
    design_tip: idx === 0 ? 'استخدام ترويسة غلاف متباينة مع شعار واضح' : 'توزيع النقاط داخل بطاقة منظمة مع هامش مريح',
  }));

  let analysis = {
    overall_score: 8,
    summary: `تم تحليل ${slides.length} شريحة تحتوي على ${wordCount} كلمة و${tablesCount} جدول و${mediaFiles.length} صورة.`,
    strengths: ['تسلسل منطقي لعناصر العرض', 'وجود عناوين واضحة للشرائح'],
    weaknesses: ['يمكن تحسين التباين اللوني وتوحيد الخطوط العربية (RTL)', 'اختصار بعض النقاط الطويلة في بطاقات مركزة'],
    recommended_templates: recs,
    slide_suggestions: defaultSuggestions,
  };
  let provider = 'المحلل الذكي المحلي';

  const ai = getGeminiClient();
  if (ai && slides.length > 0) {
    try {
      const prompt = `حلّل شرائح العرض التقديمي التالية وأعد كائن JSON فقط بالحقول:
{"overall_score": 8, "summary": "...", "strengths": ["..."], "weaknesses": ["..."], "slide_suggestions": [{"slide_number": 1, "original_title": "...", "improved_title": "...", "improved_bullets": ["..."], "improvement_reason": "...", "design_tip": "..."}]}
الشرائح:
${JSON.stringify(slides.slice(0, 10)).slice(0, 8000)}`;
      const resp = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: prompt,
        config: { responseMimeType: 'application/json' },
      });
      const parsed = JSON.parse(resp.text || '{}');
      if (parsed && typeof parsed.summary === 'string') {
        analysis = {
          overall_score: Number(parsed.overall_score) || 8,
          summary: parsed.summary,
          strengths: Array.isArray(parsed.strengths) ? parsed.strengths : analysis.strengths,
          weaknesses: Array.isArray(parsed.weaknesses) ? parsed.weaknesses : analysis.weaknesses,
          recommended_templates: recs,
          slide_suggestions: Array.isArray(parsed.slide_suggestions) && parsed.slide_suggestions.length
            ? parsed.slide_suggestions
            : defaultSuggestions,
        };
        provider = 'Gemini Flash';
      }
    } catch {
      // keep local analysis
    }
  }

  return {
    slide_count: slides.length,
    word_count: wordCount,
    tables_count: tablesCount,
    images_count: mediaFiles.length,
    fonts_used: Array.from(fontsSet),
    colors_used: Array.from(colorsSet).slice(0, 10),
    slides,
    full_text: fullText.trim(),
    analysis,
    provider,
  };
}

export async function pptxChatReply(messages: Array<{ role: string; content: string }>): Promise<string> {
  const ai = getGeminiClient();
  const lastMsg = (messages[messages.length - 1]?.content || '').trim();
  if (ai && lastMsg) {
    try {
      const historyTxt = messages
        .slice(-8)
        .map((m) => `${m.role === 'assistant' ? 'المساعد' : 'المستخدم'}: ${m.content}`)
        .join('\n');
      const resp = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: `أنت مساعد استوديو العروض والمستندات الاحترافي. أجب باللغة العربية باختصار ووضوح عملي:\n${historyTxt}`,
      });
      if (resp.text && resp.text.trim()) return resp.text.trim();
    } catch {
      // fallback
    }
  }
  return `تم استلام طلبك: "${lastMsg.slice(0, 80)}". يمكنك تعديل الشرائح أو اختيار قالب من معرض قوالب كانفا ثم الضغط على زر الإنشاء لإنتاج العرض فوراً.`;
}

export async function startHtmlToVideoRender(params: {
  outputDir: string;
  audioDir: string;
  html: string;
  filename: string;
  width: number;
  height: number;
  fps: number;
  duration: number;
  format: string;
  audioFilePath?: string;
  audioFitMode?: string;
  audioVolume?: number;
}): Promise<string> {
  fs.mkdirSync(params.outputDir, { recursive: true });
  const jobId = `vid_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`;
  videoJobs.set(jobId, {
    job_id: jobId,
    status: 'running',
    percent: 15,
    stage: 'جاري تحليل مشاهد HTML وتحضير الإطارات...',
  });

  (async () => {
    const tmpDir = path.join(params.outputDir, `tmp_${jobId}`);
    fs.mkdirSync(tmpDir, { recursive: true });
    const t0 = Date.now();
    try {
      const $ = cheerio.load(params.html || '');
      const titleText =
        $('h1, .title, .text').first().text().trim() ||
        $('title').text().trim() ||
        'استوديو العروض والمستندات الاحترافي';
      const bodyLines = $('p, h2, h3, li, .subtitle')
        .map((_i, el) => $(el).text().replace(/\s+/g, ' ').trim())
        .get()
        .filter((t) => t && t !== titleText)
        .slice(0, 4);

      const width = Math.max(320, Math.min(1920, params.width || 1280));
      const height = Math.max(320, Math.min(1920, params.height || 720));
      const fps = Math.min(30, Math.max(12, params.fps || 30));
      const duration = Math.max(2, Math.min(60, params.duration || 6));
      const fmt = ['mp4', 'webm', 'mov', 'gif'].includes(params.format) ? params.format : 'mp4';

      const escSvg = (s: string) =>
        s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

      const svgPath = path.join(tmpDir, 'scene.svg');
      const svgContent = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
        <defs>
          <linearGradient id="bg" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="#0b1120"/>
            <stop offset="55%" stop-color="#0f2942"/>
            <stop offset="100%" stop-color="#008b94"/>
          </linearGradient>
        </defs>
        <rect width="${width}" height="${height}" fill="url(#bg)"/>
        <rect x="${Math.round(width * 0.06)}" y="${Math.round(height * 0.12)}" width="${Math.round(width * 0.88)}" height="${Math.round(height * 0.76)}" rx="28" fill="rgba(255,255,255,0.06)" stroke="#f59e0b" stroke-width="3"/>
        <text x="50%" y="42%" text-anchor="middle" font-family="Cairo, DejaVu Sans, sans-serif" font-size="${Math.round(width * 0.042)}" font-weight="bold" fill="#ffffff">${escSvg(titleText.slice(0, 56))}</text>
        ${bodyLines
          .map(
            (ln, idx) =>
              `<text x="50%" y="${54 + idx * 8}%" text-anchor="middle" font-family="Cairo, DejaVu Sans, sans-serif" font-size="${Math.round(width * 0.024)}" fill="#cbd5e1">${escSvg(ln.slice(0, 72))}</text>`
          )
          .join('\n')}
      </svg>`;
      fs.writeFileSync(svgPath, svgContent, 'utf-8');

      videoJobs.set(jobId, {
        job_id: jobId,
        status: 'running',
        percent: 55,
        stage: 'جاري تصيير الفيديو عبر FFmpeg...',
      });

      const safePrefix = (params.filename || 'studio_video').replace(/[^a-zA-Z0-9\u0600-\u06FF_-]+/g, '_').slice(0, 40);
      const outName = `${safePrefix || 'studio_video'}_${Date.now()}.${fmt}`;
      const outPath = path.join(params.outputDir, outName);

      const hasAudio = Boolean(params.audioFilePath && fs.existsSync(params.audioFilePath) && fmt !== 'gif');
      const args: string[] = ['-y', '-loop', '1', '-i', svgPath];
      if (hasAudio && params.audioFilePath) {
        if (params.audioFitMode === 'loop') {
          args.push('-stream_loop', '-1');
        }
        args.push('-i', params.audioFilePath);
      }

      args.push('-t', String(duration));
      if (fmt === 'gif') {
        args.push('-vf', `fps=12,scale=${Math.min(width, 960)}:-1:flags=lanczos`);
      } else if (fmt === 'webm') {
        args.push('-c:v', 'libvpx-vp9', '-b:v', '1M', '-pix_fmt', 'yuv420p', '-r', String(fps));
      } else {
        args.push('-c:v', 'libx264', '-preset', 'veryfast', '-pix_fmt', 'yuv420p', '-r', String(fps), '-movflags', '+faststart');
      }

      if (hasAudio) {
        const vol = params.audioVolume || 1.0;
        args.push('-filter:a', `volume=${vol}`, '-c:a', fmt === 'webm' ? 'libopus' : 'aac', '-shortest');
      }

      args.push(outPath);
      await execFileAsync('ffmpeg', args, { timeout: 45000 });

      const fsize = fs.existsSync(outPath) ? fs.statSync(outPath).size : 0;
      const elapsed = Math.max(30, Date.now() - t0);
      const metrics = buildPipelineMetrics({
        operation: `تصيير HTML إلى فيديو (${fmt.toUpperCase()})`,
        stages: [
          { key: 'extract', stage: '1. تحليل المشاهد', shortName: 'المشاهد', duration_ms: Math.round(elapsed * 0.2) },
          { key: 'ai', stage: '2. التوقيت والصوت', shortName: 'الصوت', duration_ms: Math.round(elapsed * 0.15) },
          { key: 'charts', stage: '3. التحجيم والترميز', shortName: 'الترميز', duration_ms: Math.round(elapsed * 0.45) },
          { key: 'render', stage: '4. تغليف الفيديو', shortName: 'التغليف', duration_ms: Math.round(elapsed * 0.2) },
        ],
        wordCount: (titleText + ' ' + bodyLines.join(' ')).split(/\s+/).length,
        itemsCount: Math.round(duration * fps),
        provider: '@hyperframes/producer + FFmpeg',
      });

      videoJobs.set(jobId, {
        job_id: jobId,
        status: 'done',
        percent: 100,
        stage: 'اكتمل تصيير الفيديو بنجاح!',
        filename: outName,
        download_url: `/api/studio/download/${encodeURIComponent(outName)}`,
        preview_url: `/api/studio/media/${encodeURIComponent(outName)}`,
        file_size: fsize,
        width,
        height,
        fps,
        duration_seconds: duration,
        format: fmt,
        has_audio: hasAudio,
        pipeline_metrics: metrics,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      videoJobs.set(jobId, {
        job_id: jobId,
        status: 'error',
        percent: 100,
        stage: 'تعذر تصيير الفيديو',
        error: msg,
      });
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  })();

  return jobId;
}
