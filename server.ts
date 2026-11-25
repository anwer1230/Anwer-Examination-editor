import crypto from 'crypto';
import express from 'express';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import multer from 'multer';
import * as cheerio from 'cheerio';
import { createServer as createViteServer } from 'vite';
import {
  EXAMPLE_VIDEO_HTML,
  PRESENTATION_CANVA_TEMPLATES,
  type SlideItem,
  type StudioCanvaTemplate,
  aiGenerateVideoHtml,
  aiProcessHtml,
  aiRefineVideoHtml,
  analyzePptxBuffer,
  buildDocxFile,
  buildExcelFromHtml,
  buildPipelineMetrics,
  buildPptxFile,
  extractFileContent,
  generateSlidesStructure,
  generateStudioAudioPreview,
  getGeminiClient,
  getVideoJob,
  parseHtmlDocument,
  pptxChatReply,
  saveUploadedStudioAudio,
  startHtmlToVideoRender,
} from './server/studio_engine.ts';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = Number(process.env.PORT) || 3000;
const IS_PROD = process.env.NODE_ENV === 'production';

const DATA_DIR = path.join(__dirname, 'data');
const FETCHED_CANVA_FILE = path.join(DATA_DIR, 'edu_canva_fetched.json');
const STUDIO_CANVA_CUSTOM_FILE = path.join(DATA_DIR, 'studio_canva_custom.json');
const GOOGLE_ACCOUNTS_FILE = path.join(DATA_DIR, 'canva_google_accounts.json');
const PROCESSED_CANVA_DIR = path.join(DATA_DIR, 'edu_canva_processed');
const CANVA_IMG_CACHE_DIR = path.join(DATA_DIR, 'canva_img_cache');
const STUDIO_OUTPUTS_DIR = path.join(DATA_DIR, 'outputs');
const STUDIO_AUDIO_DIR = path.join(DATA_DIR, 'studio_audio');

fs.mkdirSync(DATA_DIR, { recursive: true });
fs.mkdirSync(PROCESSED_CANVA_DIR, { recursive: true });
fs.mkdirSync(CANVA_IMG_CACHE_DIR, { recursive: true });
fs.mkdirSync(STUDIO_OUTPUTS_DIR, { recursive: true });
fs.mkdirSync(STUDIO_AUDIO_DIR, { recursive: true });

interface CanvaColors {
  primary: string;
  secondary: string;
  accent: string;
  bg: string;
  card_bg: string;
  text_dark: string;
  text_light: string;
}

interface CanvaTemplate {
  id: string;
  canva_id?: string;
  name: string;
  category: string;
  source: string;
  layout: string;
  font_family: string;
  heading_font: string;
  description: string;
  colors: CanvaColors;
  preview_image?: string;
  safe_preview_data_url?: string;
  fallback_preview_data_url?: string;
  canva_url?: string;
  badge_text?: string;
  fetched_at?: string;
  processed?: boolean;
  processed_asset_id?: string;
  processed_pages?: string[];
  processed_pdf_url?: string;
  processed_format?: string;
}

interface GoogleAccountEntry {
  email: string;
  name?: string;
  photo_url?: string;
  provider?: string;
  added_at?: string;
}

interface GoogleAccountsStore {
  accounts: GoogleAccountEntry[];
  active_email: string;
}

const EDU_CANVA_TEMPLATES: CanvaTemplate[] = [
  {
    id: 'canva_tpl_EAF-1kLSdBs',
    canva_id: 'EAF-1kLSdBs',
    name: 'قالب كانفا — تقرير الأداء والإنجاز السنوي A4',
    category: 'تقارير مدرسية',
    source: 'Canva Marketplace',
    layout: 'royal_double',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'قالب كانفا A4 رسمي بالزيتي والذهبي المعتمد لتقارير الإنجاز والملفات السنوية.',
    colors: {
      primary: '#2f5d50',
      secondary: '#1d3b32',
      accent: '#d4a448',
      bg: '#fafaf7',
      card_bg: '#ffffff',
      text_dark: '#13221d',
      text_light: '#ffffff',
    },
    preview_image: 'https://marketplace.canva.com/EAF-1kLSdBs/1/0/1131w/canva-green-and-yellow-simple-annual-report-a4-document-dgpah4kY-f0.jpg',
    canva_url: 'https://www.canva.com/templates/EAF-1kLSdBs-green-and-yellow-simple-annual-report-a4-document/',
    badge_text: 'Canva A4 تقرير',
  },
  {
    id: 'canva_tpl_EAG-5dQh2hI',
    canva_id: 'EAG-5dQh2hI',
    name: 'قالب كانفا — الفيروزي والأزرق للتقارير الحديثة A4',
    category: 'ملفات أداء وإنجاز',
    source: 'Canva Marketplace',
    layout: 'gradient_arch',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'تصميم كانفا فيروزي حديث بتنسيق A4 منظّم لملفات الأداء الوظيفي والجداول.',
    colors: {
      primary: '#1f5f74',
      secondary: '#123c4a',
      accent: '#4ebbdf',
      bg: '#f4f9fb',
      card_bg: '#ffffff',
      text_dark: '#0f242c',
      text_light: '#ffffff',
    },
    preview_image: 'https://marketplace.canva.com/EAG-5dQh2hI/1/0/283w/canva-blue-and-teal-modern-supermarket-annual-report-a4-GizZY_IqcS0.jpg',
    canva_url: 'https://www.canva.com/templates/EAG-5dQh2hI-blue-and-teal-modern-annual-report-a4/',
    badge_text: 'Canva A4 أداء',
  },
  {
    id: 'canva_tpl_EAF-BzwpAQk',
    canva_id: 'EAF-BzwpAQk',
    name: 'قالب كانفا — الإطار الزهري المائي الأزرق A4',
    category: 'شواهد وتوثيق',
    source: 'Canva Marketplace',
    layout: 'botanical_frame',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'تصميم كانفا ناعم بألوان مائية وزهور زرقاء مثالي لتوثيق الشواهد والسيرة التربوية.',
    colors: {
      primary: '#2f5d80',
      secondary: '#1b3a52',
      accent: '#7fb0d9',
      bg: '#f7fbff',
      card_bg: '#ffffff',
      text_dark: '#152635',
      text_light: '#ffffff',
    },
    preview_image: 'https://marketplace.canva.com/EAF-BzwpAQk/1/0/400w/canva-blue-watercolor-floral-blank-background-document-a4-%28landscape%29-7z1qpPONOYQ.jpg',
    canva_url: 'https://www.canva.com/templates/EAF-BzwpAQk-blue-watercolor-floral-blank-background-document-a4/',
    badge_text: 'Canva شواهد',
  },
  {
    id: 'canva_tpl_EAF-CqXas4E',
    canva_id: 'EAF-CqXas4E',
    name: 'قالب كانفا — الوردي الكلاسيكي الهادئ A4',
    category: 'شواهد وتوثيق',
    source: 'Canva Marketplace',
    layout: 'pastel_card',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'قالب كانفا أنيق بالوردي الهادئ والإطارات الكلاسيكية لملفات المعلمات والشواهد.',
    colors: {
      primary: '#8c4a58',
      secondary: '#5c2d37',
      accent: '#e09f97',
      bg: '#fff8f8',
      card_bg: '#ffffff',
      text_dark: '#2d181e',
      text_light: '#ffffff',
    },
    preview_image: 'https://marketplace.canva.com/EAF-CqXas4E/1/0/566w/canva-pink-vintage-thank-you-letter-a4-9gvBKO2ywtY.jpg',
    canva_url: 'https://www.canva.com/templates/EAF-CqXas4E-pink-vintage-thank-you-letter-a4/',
    badge_text: 'Canva A4 ناعم',
  },
  {
    id: 'canva_tpl_EAG-5Ng6h30',
    canva_id: 'EAG-5Ng6h30',
    name: 'قالب كانفا — البنفسجي المعماري الحديث A4',
    category: 'ملفات أداء وإنجاز',
    source: 'Canva Marketplace',
    layout: 'clean_stripe',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'تصميم كانفا بنفسجي وعاجي بلمسة تنفيذية راقية لعرض معايير الأداء الوظيفي.',
    colors: {
      primary: '#564a8e',
      secondary: '#382f62',
      accent: '#c9a86a',
      bg: '#faf8ff',
      card_bg: '#ffffff',
      text_dark: '#1d1833',
      text_light: '#ffffff',
    },
    preview_image: 'https://marketplace.canva.com/EAG-5Ng6h30/1/0/1131w/canva-documento-a4-brief-para-proyecto-moderno-morado-beige-negro-VnEvWKVIxnQ.jpg',
    canva_url: 'https://www.canva.com/templates/EAG-5Ng6h30-documento-a4-brief-para-proyecto-moderno/',
    badge_text: 'Canva A4 ملف',
  },
  {
    id: 'canva_tpl_EAF-JNz4vIQ',
    canva_id: 'EAF-JNz4vIQ',
    name: 'قالب كانفا — الأخضر والذهبي للشهادات والتكريم',
    category: 'شهادات وتكريم',
    source: 'Canva Marketplace',
    layout: 'royal_double',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'تصميم كانفا زمردي وذهبي فخم مناسب لشهادات التقدير والتكريم والشواهد الرسمية.',
    colors: {
      primary: '#14532d',
      secondary: '#052e16',
      accent: '#d97706',
      bg: '#fafaf5',
      card_bg: '#ffffff',
      text_dark: '#13261a',
      text_light: '#ffffff',
    },
    preview_image: 'https://marketplace.canva.com/EAF-JNz4vIQ/1/0/400w/canva-green-and-gold-modern-sport-competition-certificate-landscape-ja1wKVOZM40.jpg',
    canva_url: 'https://www.canva.com/templates/EAF-JNz4vIQ-green-and-gold-modern-certificate/',
    badge_text: 'Canva شهادة',
  },
  {
    id: 'canva_tpl_EAF-tJtgCxA',
    canva_id: 'EAF-tJtgCxA',
    name: 'قالب كانفا — العنابي والذهبي للشهادات التقديرية',
    category: 'شهادات وتكريم',
    source: 'Canva Marketplace',
    layout: 'modern_ribbon',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'تصميم كانفا فخم بالعنابي العميق والذهبي لإبراز الشهادات والمنجزات التربوية.',
    colors: {
      primary: '#7f1d1d',
      secondary: '#450a0a',
      accent: '#eab308',
      bg: '#fffaf8',
      card_bg: '#ffffff',
      text_dark: '#291212',
      text_light: '#ffffff',
    },
    preview_image: 'https://marketplace.canva.com/EAF-tJtgCxA/1/0/400w/canva-red-and-yellow-minimalist-certificate-of-appreciation-QPPDeBB-7BM.jpg',
    canva_url: 'https://www.canva.com/templates/EAF-tJtgCxA-red-and-yellow-minimalist-certificate-of-appreciation/',
    badge_text: 'Canva تقدير',
  },
  {
    id: 'canva_tpl_EAF-dadDGhw',
    canva_id: 'EAF-dadDGhw',
    name: 'قالب كانفا — الأزرق الملكي والبرتقالي لبورتفوليو الإنجاز',
    category: 'ملفات أداء وإنجاز',
    source: 'Canva Marketplace',
    layout: 'geometric_split',
    font_family: 'Cairo',
    heading_font: 'Cairo',
    description: 'هوية كانفا جريئة بالأزرق الكوبالت والعنبر مثالية لبورتفوليو المعلم وتوثيق المبادرات.',
    colors: {
      primary: '#124698',
      secondary: '#0c2f67',
      accent: '#f59e0b',
      bg: '#f5f8ff',
      card_bg: '#ffffff',
      text_dark: '#0f1e3a',
      text_light: '#ffffff',
    },
    preview_image: 'https://marketplace.canva.com/EAF-dadDGhw/1/0/400w/canva-modern-bold-graphic-designer-creative-portfolio-presentation-W_nnHi_vuJI.jpg',
    canva_url: 'https://www.canva.com/templates/EAF-dadDGhw-modern-bold-creative-portfolio-presentation/',
    badge_text: 'Canva بورتفوليو',
  },
  {
    id: 'canva_tpl_EAF-D6D5Q8Q',
    canva_id: 'EAF-D6D5Q8Q',
    name: 'قالب كانفا — البيج والذهب الدافئ للتخطيط الأسبوعي A4',
    category: 'شواهد وتوثيق',
    source: 'Canva Marketplace',
    layout: 'pastel_card',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'تصميم كانفا هادئ بالبيج والذهبي الرملي يناسب الخطط والجداول والملفات التنظيمية.',
    colors: {
      primary: '#6e5c4f',
      secondary: '#463930',
      accent: '#c8a97e',
      bg: '#fdfaf6',
      card_bg: '#ffffff',
      text_dark: '#261e18',
      text_light: '#ffffff',
    },
    preview_image: 'https://marketplace.canva.com/EAF-D6D5Q8Q/1/0/566w/canva-beige-and-gold-minimal-weekly-a4-planner-VPvaDK9Stjw.jpg',
    canva_url: 'https://www.canva.com/templates/EAF-D6D5Q8Q-beige-and-gold-minimal-weekly-a4-planner/',
    badge_text: 'Canva A4 خطط',
  },
  {
    id: 'canva_tpl_EAF-5dONlxQ',
    canva_id: 'EAF-5dONlxQ',
    name: 'قالب كانفا — الرمادي الرسمي التحريري A4',
    category: 'تقارير مدرسية',
    source: 'Canva Marketplace',
    layout: 'clean_stripe',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'خطوط تحريرية نظيفة ومستند رسمي عالي الوضوح للتقارير المدرسية والخطابات المعتمدة.',
    colors: {
      primary: '#2d3748',
      secondary: '#1a202c',
      accent: '#0d9488',
      bg: '#ffffff',
      card_bg: '#f8fafc',
      text_dark: '#0f172a',
      text_light: '#ffffff',
    },
    preview_image: 'https://marketplace.canva.com/EAF-5dONlxQ/1/0/1131w/canva-documento-a4-carta-formal-blanco-y-gris-vCza1-eQ7gQ.jpg',
    canva_url: 'https://www.canva.com/templates/EAF-5dONlxQ-documento-a4-carta-formal-blanco-y-gris/',
    badge_text: 'Canva A4 رسمي',
  },
  {
    id: 'canva_tpl_EAG-d21lfTM',
    canva_id: 'EAG-d21lfTM',
    name: 'قالب كانفا — الأزرق الهندسي لتقارير البرامج والأنشطة',
    category: 'تقارير مدرسية',
    source: 'Canva Marketplace',
    layout: 'geometric_split',
    font_family: 'Cairo',
    heading_font: 'Cairo',
    description: 'تصميم كانفا هندسي بدرجات الأزرق والأردوازي لتقارير البرامج والأنشطة الطلابية.',
    colors: {
      primary: '#2b4596',
      secondary: '#1b2c63',
      accent: '#38bdf8',
      bg: '#f3f6ff',
      card_bg: '#ffffff',
      text_dark: '#101b3d',
      text_light: '#ffffff',
    },
    preview_image: 'https://marketplace.canva.com/EAG-d21lfTM/1/0/400w/canva-blue-gray-bold-social-media-report-presentation-_m7kW6xIZmE.jpg',
    canva_url: 'https://www.canva.com/templates/EAG-d21lfTM-blue-gray-bold-report-presentation/',
    badge_text: 'Canva تقارير',
  },
  {
    id: 'canva_tpl_EAF--nReQEY',
    canva_id: 'EAF--nReQEY',
    name: 'قالب كانفا — المدرسي التوضيحي الأزرق A4',
    category: 'تقارير مدرسية',
    source: 'Canva Marketplace',
    layout: 'gradient_arch',
    font_family: 'Cairo',
    heading_font: 'Cairo',
    description: 'قالب كانفا تعليمي مشرق مناسب للجداول المدرسية والأنشطة والتقارير الصفية.',
    colors: {
      primary: '#26538d',
      secondary: '#16355e',
      accent: '#60a5fa',
      bg: '#f0f7ff',
      card_bg: '#ffffff',
      text_dark: '#10233d',
      text_light: '#ffffff',
    },
    preview_image: 'https://marketplace.canva.com/EAF--nReQEY/1/0/400w/canva-biru-putih-ilustratif-jadwal-pelajaran-sekolah-a4-lanskap-IxOupTK0efI.jpg',
    canva_url: 'https://www.canva.com/templates/EAF--nReQEY-school-schedule-a4/',
    badge_text: 'Canva مدرسي',
  },
  {
    id: 'canva_mp_EAF-eTDNqew',
    canva_id: 'EAF-eTDNqew',
    name: 'كانفا — الصحة النفسية والإرشاد المدرسي الأخضر A4',
    category: 'شواهد وتوثيق',
    source: 'Canva Marketplace Live',
    layout: 'botanical_frame',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'غلاف مستند تعليمي A4 أنيق بالتدرج الأخضر يناسب التوجيه الطلابي والشواهد التربوية.',
    colors: { primary: '#166534', secondary: '#14532d', accent: '#10b981', bg: '#f3fbf6', card_bg: '#ffffff', text_dark: '#13261a', text_light: '#ffffff' },
    preview_image: 'https://marketplace.canva.com/EAF-eTDNqew/1/0/1131w/canva-documento-a4-portada-salud-mental-ilustrado-verde-cRVCjT60RXk.jpg',
    canva_url: 'https://www.canva.com/templates/EAF-eTDNqew-documento-a4-portada-salud-mental-ilustrado-verde/',
    badge_text: 'Canva شواهد',
  },
  {
    id: 'canva_mp_EAFA1-LEZ3k',
    canva_id: 'EAFA1-LEZ3k',
    name: 'كانفا — أوراق الزهور المائية للتوثيق A4',
    category: 'شواهد وتوثيق',
    source: 'Canva Marketplace Live',
    layout: 'botanical_frame',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'إطار ألوان مائية بزهور رقيقة A4 مثالي لشواهد المعلمات وملفات الإنجاز.',
    colors: { primary: '#2f5d80', secondary: '#1b3a52', accent: '#7fb0d9', bg: '#f7fbff', card_bg: '#ffffff', text_dark: '#152635', text_light: '#ffffff' },
    preview_image: 'https://marketplace.canva.com/EAFA1-LEZ3k/1/0/1131w/canva-documento-a4-bloc-notas-hoja-carta-flores-acuarela-gHY3klheKY0.jpg',
    canva_url: 'https://www.canva.com/templates/EAFA1-LEZ3k-documento-a4-bloc-notas-hoja-carta-flores-acuarela/',
    badge_text: 'Canva شواهد',
  },
  {
    id: 'canva_mp_EAFa53bSQh0',
    canva_id: 'EAFa53bSQh0',
    name: 'كانفا — الإطار الوردي المتدرج الناعم A4',
    category: 'شواهد وتوثيق',
    source: 'Canva Marketplace Live',
    layout: 'pastel_card',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'تصميم A4 متدرج بالوردي والمشمشي لإبراز الشواهد والمبادرات التربوية.',
    colors: { primary: '#8c4a58', secondary: '#5c2d37', accent: '#e09f97', bg: '#fff8f9', card_bg: '#ffffff', text_dark: '#2d181e', text_light: '#ffffff' },
    preview_image: 'https://marketplace.canva.com/EAFa53bSQh0/1/0/1131w/canva-orange-pink-gradient-background-stationery-border-a4-document-hnYishFWjuU.jpg',
    canva_url: 'https://www.canva.com/templates/EAFa53bSQh0-orange-pink-gradient-background-stationery-border-a4-document/',
    badge_text: 'Canva شواهد',
  },
  {
    id: 'canva_mp_EAFa-nnOb9c',
    canva_id: 'EAFa-nnOb9c',
    name: 'كانفا — الغلاف الفني الخطي البني والبيج A4',
    category: 'شواهد وتوثيق',
    source: 'Canva Marketplace Live',
    layout: 'pastel_card',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'غلاف يوميات وتوثيق هادئ بدرجات البيج والبني الدافئ بمقاس A4.',
    colors: { primary: '#6b4f2d', secondary: '#433018', accent: '#d97706', bg: '#fdfaf4', card_bg: '#ffffff', text_dark: '#24190c', text_light: '#ffffff' },
    preview_image: 'https://marketplace.canva.com/EAFa-nnOb9c/1/0/1131w/canva-beige-brown-aesthetic-line-art-daily-journal-cover-a4-document-L2QgVHsdzgc.jpg',
    canva_url: 'https://www.canva.com/templates/EAFa-nnOb9c-beige-brown-aesthetic-line-art-daily-journal-cover-a4-document/',
    badge_text: 'Canva شواهد',
  },
  {
    id: 'canva_mp_EAF-ERcfPrA',
    canva_id: 'EAF-ERcfPrA',
    name: 'كانفا — الغلاف الإبداعي الملون للأنشطة الطلابية A4',
    category: 'شواهد وتوثيق',
    source: 'Canva Marketplace Live',
    layout: 'gradient_arch',
    font_family: 'Cairo',
    heading_font: 'Cairo',
    description: 'غلاف تعليمي حيوي متعدد الألوان يناسب أنشطة الطالبات والمسابقات المدرسية.',
    colors: { primary: '#1e4b6e', secondary: '#102e47', accent: '#00b4d8', bg: '#f4f9fc', card_bg: '#ffffff', text_dark: '#0f172a', text_light: '#ffffff' },
    preview_image: 'https://marketplace.canva.com/EAF-ERcfPrA/1/0/1131w/canva-documento-a4-portada-de-cuaderno-comic-creativo-multicolor-FUyzdkmeFGk.jpg',
    canva_url: 'https://www.canva.com/templates/EAF-ERcfPrA-documento-a4-portada-de-cuaderno-comic-creativo-multicolor/',
    badge_text: 'Canva شواهد',
  },
  {
    id: 'canva_mp_EAFA67337ZE',
    canva_id: 'EAFA67337ZE',
    name: 'كانفا — شهادة التميز الفضية والذهبية',
    category: 'شهادات وتكريم',
    source: 'Canva Marketplace Live',
    layout: 'royal_double',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'شهادة تكريم رسمية بتفاصيل فضية وذهبية لتكريم الطالبات والمعلمات.',
    colors: { primary: '#6b4f2d', secondary: '#433018', accent: '#d97706', bg: '#fdfaf4', card_bg: '#ffffff', text_dark: '#24190c', text_light: '#ffffff' },
    preview_image: 'https://marketplace.canva.com/EAFA67337ZE/1/0/400w/canva-silver-and-gold-minimalist-employee-of-the-month-certificate-CXrfuYm5Dbw.jpg',
    canva_url: 'https://www.canva.com/templates/EAFA67337ZE-silver-and-gold-minimalist-employee-of-the-month-certificate/',
    badge_text: 'Canva شهادة',
  },
  {
    id: 'canva_mp_EAFA7nFGR6w',
    canva_id: 'EAFA7nFGR6w',
    name: 'كانفا — بوستر الأنشطة المدرسية والأندية الوردي والأزرق',
    category: 'تقارير مدرسية',
    source: 'Canva Marketplace Live',
    layout: 'pastel_card',
    font_family: 'Cairo',
    heading_font: 'Cairo',
    description: 'تصميم مشرق لبرامج الأنشطة المدرسية والأندية والمبادرات الطلابية.',
    colors: { primary: '#8c4a58', secondary: '#5c2d37', accent: '#e09f97', bg: '#fff8f9', card_bg: '#ffffff', text_dark: '#2d181e', text_light: '#ffffff' },
    preview_image: 'https://marketplace.canva.com/EAFA7nFGR6w/1/0/1131w/canva-pink-blue-playful-modern-illustration-activity-club-school-poster-p6e3vWx0qzg.jpg',
    canva_url: 'https://www.canva.com/templates/EAFA7nFGR6w-pink-blue-playful-modern-illustration-activity-club-school-poster/',
    badge_text: 'Canva تقرير',
  },
  {
    id: 'canva_mp_EAFA9Zq4z_g',
    canva_id: 'EAFA9Zq4z_g',
    name: 'كانفا — قالب المشاريع الجماعية والتعلم النشط المدرسي',
    category: 'تقارير مدرسية',
    source: 'Canva Marketplace Live',
    layout: 'pastel_card',
    font_family: 'Cairo',
    heading_font: 'Cairo',
    description: 'تصميم باستيل لطيف للتعلم التعاوني ومشاريع الفصول الابتدائية.',
    colors: { primary: '#26538d', secondary: '#16355e', accent: '#f59e0b', bg: '#f8fafc', card_bg: '#ffffff', text_dark: '#0f172a', text_light: '#ffffff' },
    preview_image: 'https://marketplace.canva.com/EAFA9Zq4z_g/1/0/400w/canva-yellow-blue-pastel-cute-simple-3d-group-project-education-school-presentation-KU1JYeScHgE.jpg',
    canva_url: 'https://www.canva.com/templates/EAFA9Zq4z_g-yellow-blue-pastel-cute-simple-3d-group-project-education-school-presentation/',
    badge_text: 'Canva تقرير',
  },
  {
    id: 'canva_mp_EAG-_LAt630',
    canva_id: 'EAG-_LAt630',
    name: 'كانفا — مستند التقارير الإبداعية الحديث A4',
    category: 'تقارير مدرسية',
    source: 'Canva Marketplace Live',
    layout: 'gradient_arch',
    font_family: 'Cairo',
    heading_font: 'Cairo',
    description: 'مستند تقارير حديث بألوان واضحة وتنسيق منظم للتقارير والخطط المدرسية.',
    colors: { primary: '#0f766e', secondary: '#115e59', accent: '#f97316', bg: '#f4fbf9', card_bg: '#ffffff', text_dark: '#0f172a', text_light: '#ffffff' },
    preview_image: 'https://marketplace.canva.com/EAG-_LAt630/1/0/1131w/canva-doc-reporte-de-marketing-moderno-rosa-verde-y-naranja-DmOoLghCffo.jpg',
    canva_url: 'https://www.canva.com/templates/EAG-_LAt630-doc-reporte-de-marketing-moderno-rosa-verde-y-naranja/',
    badge_text: 'Canva تقرير',
  },
  {
    id: 'canva_mp_EAF-db-OMBs',
    canva_id: 'EAF-db-OMBs',
    name: 'كانفا — ملف السيرة الذاتية والإنجاز الأزرق والأصفر A4',
    category: 'ملفات أداء وإنجاز',
    source: 'Canva Marketplace Live',
    layout: 'clean_stripe',
    font_family: 'Cairo',
    heading_font: 'Cairo',
    description: 'قالب سيرة ذاتية وبورتفوليو مهني منظم بالأزرق والأصفر مقاس A4.',
    colors: { primary: '#1e4b6e', secondary: '#102e47', accent: '#eab308', bg: '#f4f9fc', card_bg: '#ffffff', text_dark: '#0f172a', text_light: '#ffffff' },
    preview_image: 'https://marketplace.canva.com/EAF-db-OMBs/1/0/1131w/canva-kuning-dan-biru-profesional-cv-resume-marketing-a4-Eio_DpnoxrU.jpg',
    canva_url: 'https://www.canva.com/templates/EAF-db-OMBs-kuning-dan-biru-profesional-cv-resume-marketing-a4-eio-dpnoxru/',
    badge_text: 'Canva سيرة',
  },
  {
    id: 'canva_mp_EAG-ho8gmxg',
    canva_id: 'EAG-ho8gmxg',
    name: 'كانفا — البورتفوليو الكريمي الحديث لملفات الإنجاز',
    category: 'ملفات أداء وإنجاز',
    source: 'Canva Marketplace Live',
    layout: 'pastel_card',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'تصميم كريمي هادئ ومرتب لعرض ملفات الأداء الوظيفي والسيرة المهنية.',
    colors: { primary: '#6b4f2d', secondary: '#433018', accent: '#d97706', bg: '#fdfaf4', card_bg: '#ffffff', text_dark: '#24190c', text_light: '#ffffff' },
    preview_image: 'https://marketplace.canva.com/EAG-ho8gmxg/1/0/400w/canva-cream-modern-content-creator-portfolio-presentation-owqRFqQ5PSQ.jpg',
    canva_url: 'https://www.canva.com/templates/EAG-ho8gmxg-cream-modern-content-creator-portfolio-presentation/',
    badge_text: 'Canva A4',
  },
  {
    id: 'canva_mp_EAG-MANo3Ro',
    canva_id: 'EAG-MANo3Ro',
    name: 'كانفا — كتيب البورتفوليو العصري الأزرق والذهبي A4',
    category: 'ملفات أداء وإنجاز',
    source: 'Canva Marketplace Live',
    layout: 'modern_ribbon',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'تصميم كتيب إنجاز وبورتفوليو تعليمي متكامل بدرجات الأزرق والذهبي.',
    colors: { primary: '#1e4b6e', secondary: '#102e47', accent: '#f59e0b', bg: '#f8fafc', card_bg: '#ffffff', text_dark: '#0f172a', text_light: '#ffffff' },
    preview_image: 'https://marketplace.canva.com/EAG-MANo3Ro/1/0/1131w/canva-orange-yellow-and-blue-modern-interior-designer-portfolio-booklet-hYIEUzUwEc4.jpg',
    canva_url: 'https://www.canva.com/templates/EAG-MANo3Ro-orange-yellow-and-blue-modern-interior-designer-portfolio-booklet/',
    badge_text: 'Canva كتيب',
  },
  {
    id: 'canva_mp_EAFA7E4zF6w',
    canva_id: 'EAFA7E4zF6w',
    name: 'كانفا — بوستر البرامج والفعاليات المدرسية الوردي',
    category: 'تقارير مدرسية',
    source: 'Canva Marketplace Live',
    layout: 'pastel_card',
    font_family: 'Cairo',
    heading_font: 'Cairo',
    description: 'قالب فعاليات وأنشطة مدرسية بطابع عصري واضح.',
    colors: { primary: '#8c4a58', secondary: '#5c2d37', accent: '#e09f97', bg: '#fff8f9', card_bg: '#ffffff', text_dark: '#2d181e', text_light: '#ffffff' },
    preview_image: 'https://marketplace.canva.com/EAFA7E4zF6w/1/0/1131w/canva-white-pink-modern-illustration-music-club-school-poster-UjcXJVgraXc.jpg',
    canva_url: 'https://www.canva.com/templates/EAFA7E4zF6w-white-pink-modern-illustration-music-club-school-poster/',
    badge_text: 'Canva فعاليات',
  },
  {
    id: 'canva_mp_EAG-Ioa44aM',
    canva_id: 'EAG-Ioa44aM',
    name: 'كانفا — البورتفوليو الكلاسيكي الأسود والبيج',
    category: 'ملفات أداء وإنجاز',
    source: 'Canva Marketplace Live',
    layout: 'clean_stripe',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: 'تنسيق تحريري رسمي بالبيج والأسود لملفات الأداء والشواهد.',
    colors: { primary: '#2d3748', secondary: '#1a202c', accent: '#c8a97e', bg: '#fdfaf6', card_bg: '#ffffff', text_dark: '#1a202c', text_light: '#ffffff' },
    preview_image: 'https://marketplace.canva.com/EAG-Ioa44aM/1/0/400w/canva-nero-bianco-e-beige-tipografico-portfolio-presentazione-i0xStsdE5Vc.jpg',
    canva_url: 'https://www.canva.com/templates/EAG-Ioa44aM-nero-bianco-e-beige-tipografico-portfolio-presentazione/',
    badge_text: 'Canva كلاسيك',
  },
];

function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function buildCanvaSvgMarkup(tpl: {
  name?: string;
  badge_text?: string;
  category?: string;
  colors: CanvaColors;
}): string {
  const c = tpl.colors;
  const badge = escapeXml((tpl.badge_text || tpl.category || 'Canva A4').slice(0, 24));
  const title = escapeXml((tpl.name || 'قالب كانفا تعليمي').replace('قالب كانفا — ', '').slice(0, 34));
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 260" width="400" height="260">
    <defs>
      <linearGradient id="g1" x1="0%" y1="0%" x2="100%" y2="100%">
        <stop offset="0%" stop-color="${c.primary}"/>
        <stop offset="100%" stop-color="${c.secondary}"/>
      </linearGradient>
    </defs>
    <rect width="400" height="260" fill="${c.bg}"/>
    <rect x="14" y="14" width="372" height="232" rx="10" fill="none" stroke="${c.primary}" stroke-width="2.5"/>
    <rect x="32" y="28" width="336" height="54" rx="10" fill="url(#g1)"/>
    <rect x="44" y="42" width="90" height="22" rx="11" fill="${c.accent}"/>
    <text x="89" y="57" text-anchor="middle" font-family="Cairo, Arial, sans-serif" font-size="10" font-weight="bold" fill="#ffffff">${badge}</text>
    <text x="352" y="60" text-anchor="end" font-family="Cairo, Arial, sans-serif" font-size="13" font-weight="bold" fill="#ffffff">${title}</text>
    <rect x="44" y="102" width="312" height="16" rx="4" fill="${c.primary}" opacity="0.16"/>
    <rect x="44" y="128" width="240" height="12" rx="4" fill="${c.primary}" opacity="0.1"/>
    <rect x="44" y="150" width="280" height="12" rx="4" fill="${c.primary}" opacity="0.1"/>
    <circle cx="62" cy="226" r="6" fill="${c.primary}"/>
    <circle cx="78" cy="226" r="6" fill="${c.secondary}"/>
    <circle cx="94" cy="226" r="6" fill="${c.accent}"/>
  </svg>`;
}

function buildCanvaSvgDataUrl(tpl: {
  name?: string;
  badge_text?: string;
  category?: string;
  colors: CanvaColors;
}): string {
  return `data:image/svg+xml;base64,${Buffer.from(buildCanvaSvgMarkup(tpl), 'utf-8').toString('base64')}`;
}

// Precompute fallback SVGs once at startup so /api/edu/canva/browse responds in <1ms
const PREPARED_EDU_CANVA_TEMPLATES: CanvaTemplate[] = EDU_CANVA_TEMPLATES.map((tpl) => {
  const svgUrl = buildCanvaSvgDataUrl(tpl);
  return {
    ...tpl,
    safe_preview_data_url: tpl.safe_preview_data_url || svgUrl,
    fallback_preview_data_url: tpl.fallback_preview_data_url || svgUrl,
  };
});

// Fast in-memory + disk image cache & in-flight deduplication for Canva preview images
const canvaImgMemCache = new Map<string, { ctype: string; buf: Buffer }>();
const canvaImgInflight = new Map<string, Promise<{ ctype: string; buf: Buffer } | null>>();

function isAllowedCanvaImageHost(hostname: string): boolean {
  const host = hostname.toLowerCase();
  return (
    host.endsWith('canva.com') ||
    host.endsWith('canva-static.com') ||
    host.endsWith('canva-cdn.com') ||
    host.endsWith('googleusercontent.com')
  );
}

async function fetchAndCacheCanvaImage(imgUrl: string, timeoutMs = 4500): Promise<{ ctype: string; buf: Buffer } | null> {
  if (!imgUrl || (!imgUrl.startsWith('https://') && !imgUrl.startsWith('http://'))) {
    return null;
  }
  const memHit = canvaImgMemCache.get(imgUrl);
  if (memHit) return memHit;

  let parsed: URL;
  try {
    parsed = new URL(imgUrl);
  } catch {
    return null;
  }
  if (!isAllowedCanvaImageHost(parsed.hostname)) {
    return null;
  }

  const hash = crypto.createHash('sha1').update(imgUrl).digest('hex');
  const cacheMetaPath = path.join(CANVA_IMG_CACHE_DIR, `${hash}.meta.json`);
  const cacheBinPath = path.join(CANVA_IMG_CACHE_DIR, `${hash}.bin`);

  try {
    if (fs.existsSync(cacheMetaPath) && fs.existsSync(cacheBinPath)) {
      const meta = loadJsonFile<{ ctype?: string }>(cacheMetaPath, {});
      if (meta.ctype && /^image\/(jpeg|jpg|png|webp|gif)/i.test(meta.ctype)) {
        const buf = fs.readFileSync(cacheBinPath);
        if (buf.length > 200) {
          const entry = { ctype: meta.ctype, buf };
          canvaImgMemCache.set(imgUrl, entry);
          return entry;
        }
      }
    }
  } catch {
    // ignore cache read issue and fetch fresh
  }

  const existingPromise = canvaImgInflight.get(imgUrl);
  if (existingPromise) return existingPromise;

  const task = (async (): Promise<{ ctype: string; buf: Buffer } | null> => {
    const candidateUrls = [imgUrl];
    if (imgUrl.includes('/1131w/')) {
      candidateUrls.push(imgUrl.replace('/1131w/', '/400w/'));
    }
    for (const candidate of candidateUrls) {
      try {
        const upstream = await fetch(candidate, {
          signal: AbortSignal.timeout(timeoutMs),
          headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/126.0 Safari/537.36',
            Accept: 'image/avif,image/webp,image/apng,image/*,*/*;q=0.8',
          },
        });
        const ctype = (upstream.headers.get('content-type') || '').split(';')[0].trim().toLowerCase();
        if (!upstream.ok || !/^image\/(jpeg|jpg|png|webp|gif)/i.test(ctype)) {
          continue;
        }
        const buf = Buffer.from(await upstream.arrayBuffer());
        if (buf.length > 200 && buf.length < 6 * 1024 * 1024) {
          const entry = { ctype, buf };
          canvaImgMemCache.set(imgUrl, entry);
          try {
            fs.writeFileSync(cacheBinPath, buf);
            saveJsonFile(cacheMetaPath, { ctype });
          } catch {
            // ignore cache write issue
          }
          return entry;
        }
      } catch {
        // try next candidate
      }
    }
    return null;
  })().finally(() => {
    canvaImgInflight.delete(imgUrl);
  });

  canvaImgInflight.set(imgUrl, task);
  return task;
}

async function warmCanvaImageCache(): Promise<void> {
  const urls = Array.from(
    new Set(
      [...PREPARED_EDU_CANVA_TEMPLATES, ...loadFetchedCanva(), ...PRESENTATION_CANVA_TEMPLATES]
        .map((t) => String(t.preview_image || '').trim())
        .filter((u) => u.startsWith('https://') || u.startsWith('http://'))
    )
  );
  const batchSize = 4;
  for (let i = 0; i < urls.length; i += batchSize) {
    const slice = urls.slice(i, i + batchSize);
    await Promise.allSettled(slice.map((u) => fetchAndCacheCanvaImage(u, 6000)));
  }
}

function loadJsonFile<T>(filePath: string, fallback: T): T {
  try {
    if (!fs.existsSync(filePath)) return fallback;
    return JSON.parse(fs.readFileSync(filePath, 'utf-8')) as T;
  } catch {
    return fallback;
  }
}

function saveJsonFile(filePath: string, data: unknown): void {
  try {
    fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf-8');
  } catch {
    // ignore read-only error
  }
}

function loadFetchedCanva(): CanvaTemplate[] {
  const arr = loadJsonFile<CanvaTemplate[]>(FETCHED_CANVA_FILE, []);
  return Array.isArray(arr) ? arr : [];
}

function saveFetchedCanva(items: CanvaTemplate[]): void {
  saveJsonFile(FETCHED_CANVA_FILE, items.slice(0, 40));
}

function loadStudioCustomCanva(): StudioCanvaTemplate[] {
  const arr = loadJsonFile<StudioCanvaTemplate[]>(STUDIO_CANVA_CUSTOM_FILE, []);
  return Array.isArray(arr) ? arr : [];
}

function saveStudioCustomCanva(items: StudioCanvaTemplate[]): void {
  saveJsonFile(STUDIO_CANVA_CUSTOM_FILE, items.slice(0, 40));
}

function loadGoogleAccounts(): GoogleAccountsStore {
  const raw = loadJsonFile<GoogleAccountsStore>(GOOGLE_ACCOUNTS_FILE, { accounts: [], active_email: '' });
  return {
    accounts: Array.isArray(raw?.accounts) ? raw.accounts : [],
    active_email: typeof raw?.active_email === 'string' ? raw.active_email : '',
  };
}

function saveGoogleAccounts(store: GoogleAccountsStore): void {
  saveJsonFile(GOOGLE_ACCOUNTS_FILE, store);
}

function extractCanvaIdFromUrl(url: string): string {
  const m = url.match(/([A-Z]{2,4}[A-Za-z0-9_-]{6,})/);
  return m ? m[1] : `custom_${Date.now().toString(36)}`;
}

async function fetchCanvaOembedQuick(canvaUrl: string): Promise<{ title?: string; thumbnail_url?: string } | null> {
  try {
    const oembedUrl = `https://www.canva.com/_oembed?url=${encodeURIComponent(canvaUrl)}`;
    const resp = await fetch(oembedUrl, {
      signal: AbortSignal.timeout(2800),
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/126.0 Safari/537.36',
        Accept: 'application/json',
      },
    });
    if (resp.ok) {
      const data = (await resp.json()) as { title?: string; thumbnail_url?: string };
      if (data && (data.title || data.thumbnail_url)) {
        if (data.thumbnail_url) void fetchAndCacheCanvaImage(data.thumbnail_url, 3500);
        return data;
      }
    }
  } catch {
    // fallback to lightweight og tags
  }
  try {
    const pageResp = await fetch(canvaUrl, {
      signal: AbortSignal.timeout(2600),
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/126.0 Safari/537.36',
        Accept: 'text/html',
      },
    });
    if (!pageResp.ok) return null;
    const html = await pageResp.text();
    const $ = cheerio.load(html);
    const title = ($('meta[property="og:title"]').attr('content') || $('title').text() || '').trim();
    const thumbnail_url = ($('meta[property="og:image"]').attr('content') || '').trim();
    if (title || thumbnail_url) {
      if (thumbnail_url) void fetchAndCacheCanvaImage(thumbnail_url, 3500);
      return { title: title || undefined, thumbnail_url: thumbnail_url || undefined };
    }
  } catch {
    // ignore
  }
  return null;
}

function makeImportedTemplate(url: string, customName = '', category = 'تصاميمي المجلوبة'): CanvaTemplate {
  const cid = extractCanvaIdFromUrl(url);
  const name = customName || `قالب كانفا مستورد (${cid.slice(0, 11)})`;
  const tpl: CanvaTemplate = {
    id: `canva_imp_${cid}`,
    canva_id: cid,
    name,
    category,
    source: 'رابط كانفا مستورد',
    layout: 'gradient_arch',
    font_family: 'Cairo',
    heading_font: 'Amiri',
    description: `قالب مستورد من رابط كانفا: ${url}`,
    colors: {
      primary: '#0f766e',
      secondary: '#09323d',
      accent: '#c89b3c',
      bg: '#f8fafc',
      card_bg: '#ffffff',
      text_dark: '#0f172a',
      text_light: '#ffffff',
    },
    canva_url: url,
    badge_text: 'Canva مستورد',
  };
  const svgUrl = buildCanvaSvgDataUrl(tpl);
  tpl.safe_preview_data_url = svgUrl;
  tpl.fallback_preview_data_url = svgUrl;
  return tpl;
}

const MIME_MAP: Record<string, string> = {
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.mp4': 'video/mp4',
  '.webm': 'video/webm',
  '.mov': 'video/quicktime',
  '.gif': 'image/gif',
  '.zip': 'application/zip',
  '.wav': 'audio/wav',
  '.mp3': 'audio/mpeg',
};

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

async function startServer() {
  const app = express();

  app.use(express.json({ limit: '25mb' }));
  app.use(express.urlencoded({ extended: true, limit: '25mb' }));

  const publicDir = path.join(__dirname, 'public');
  const eduTemplatesDir = path.join(publicDir, 'edu_templates');
  const examEditorDir = path.join(publicDir, 'exam_editor');

  // Serve static assets from /static (qrcode.min.js, edu_image_cropper.js, edu_canva_studio.js, recharts_pipeline.js)
  app.use('/static', express.static(path.join(publicDir, 'static'), {
    setHeaders: (res) => {
      res.setHeader('Cache-Control', 'no-cache');
    },
  }));

  // Serve witness assets from /assets/aistudio
  app.use('/assets/aistudio', express.static(path.join(publicDir, 'assets', 'aistudio'), {
    setHeaders: (res) => {
      res.setHeader('Cache-Control', 'public, max-age=86400');
    },
  }));

  // Serve Exam Editor (محرر الاختبارات أنور الصبري)
  app.use('/exam_editor', (req, res, next) => {
    const rawPath = req.originalUrl.split('?')[0];
    if (rawPath === '/exam_editor') {
      const qsIndex = req.originalUrl.indexOf('?');
      const qs = qsIndex >= 0 ? req.originalUrl.slice(qsIndex) : '';
      res.redirect(302, `/exam_editor/${qs}`);
      return;
    }
    next();
  });
  app.get('/exam_editor/', (_req, res) => {
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
    res.sendFile(path.join(examEditorDir, 'index.html'));
  });
  app.use('/exam_editor', express.static(examEditorDir));

  // Serve Educational Services & Presentation Studio HTML pages
  const serveEduHtml = (filename: string) => (_req: express.Request, res: express.Response) => {
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
    res.sendFile(path.join(eduTemplatesDir, filename));
  };

  app.get(['/performance_portfolio', '/performance_portfolio/'], serveEduHtml('performance_portfolio.html'));
  app.get(['/performance_witnesses', '/performance_witnesses/'], serveEduHtml('performance_witnesses.html'));
  app.get(['/school_reports', '/school_reports/'], serveEduHtml('school_reports.html'));
  app.get(['/activities_portfolio', '/activities_portfolio/', '/school_activities_cards', '/school_activities_cards/'], serveEduHtml('activities_portfolio.html'));
  app.get(['/presentation_studio', '/presentation_studio/'], serveEduHtml('presentation_studio.html'));

  // ================= API: Canva Studio & Cached Proxy (Educational Services) =================
  app.all('/api/edu/canva/browse', async (req, res) => {
    const body = req.method === 'POST' ? (req.body || {}) : (req.query || {});
    const q = String(body.q || body.query || '').trim().toLowerCase();
    const cat = String(body.category || 'all').trim();
    const importUrl = String(body.import_url || body.canva_url || '').trim();
    const customName = String(body.name || '').trim();
    const accessToken = String(body.access_token || body.canva_token || '').trim();

    let importedTemplate: CanvaTemplate | null = null;
    let fetched = loadFetchedCanva();
    let liveCount = 0;
    let apiError = '';

    if (accessToken) {
      try {
        const apiUrl = new URL('https://api.canva.com/rest/v1/designs');
        apiUrl.searchParams.set('limit', '24');
        if (q) apiUrl.searchParams.set('query', q);
        const canvaResp = await fetch(apiUrl.toString(), {
          signal: AbortSignal.timeout(5500),
          headers: {
            Authorization: `Bearer ${accessToken}`,
            Accept: 'application/json',
          },
        });
        if (!canvaResp.ok) {
          const errText = await canvaResp.text().catch(() => '');
          apiError = `تعذر الاتصال بحساب كانفا (${canvaResp.status})${errText ? ': ' + errText.slice(0, 120) : ''}`;
        } else {
          const canvaData = (await canvaResp.json()) as {
            items?: Array<{
              id?: string;
              title?: string;
              thumbnail?: { url?: string };
              urls?: { view_url?: string; edit_url?: string };
            }>;
          };
          const items = Array.isArray(canvaData.items) ? canvaData.items : [];
          const syncedTemplates: CanvaTemplate[] = items
            .filter((d) => d && d.id)
            .map((d) => {
              const designUrl = d.urls?.view_url || d.urls?.edit_url || `https://www.canva.com/design/${d.id}/view`;
              const tpl: CanvaTemplate = {
                id: `canva_api_${d.id}`,
                canva_id: d.id,
                name: (d.title || `تصميم كانفا (${d.id})`).slice(0, 80),
                category: 'تصاميمي المجلوبة',
                source: 'حساب Canva المتصل',
                layout: 'gradient_arch',
                font_family: 'Cairo',
                heading_font: 'Amiri',
                description: 'تصميم مجلوب مباشرة من حسابك الشخصي في كانفا.',
                colors: {
                  primary: '#0f766e',
                  secondary: '#09323d',
                  accent: '#c89b3c',
                  bg: '#f8fafc',
                  card_bg: '#ffffff',
                  text_dark: '#0f172a',
                  text_light: '#ffffff',
                },
                preview_image: d.thumbnail?.url || '',
                canva_url: designUrl,
                badge_text: 'تصميمي في كانفا',
              };
              const fb = buildCanvaSvgDataUrl(tpl);
              tpl.safe_preview_data_url = fb;
              tpl.fallback_preview_data_url = fb;
              if (tpl.preview_image) void fetchAndCacheCanvaImage(tpl.preview_image, 4000);
              return tpl;
            });
          if (syncedTemplates.length) {
            liveCount = syncedTemplates.length;
            const syncedIds = new Set(syncedTemplates.map((x) => x.id));
            fetched = [...syncedTemplates, ...fetched.filter((x) => !syncedIds.has(x.id))];
            saveFetchedCanva(fetched);
          }
        }
      } catch (err) {
        apiError = err instanceof Error ? err.message : 'انتهت مهلة الاتصال بـ Canva API';
      }
    }

    if (importUrl.startsWith('http://') || importUrl.startsWith('https://')) {
      importedTemplate = makeImportedTemplate(importUrl, customName, cat && cat !== 'all' ? cat : 'تصاميمي المجلوبة');
      const meta = await fetchCanvaOembedQuick(importUrl);
      if (meta?.title && !customName) {
        importedTemplate.name = meta.title.slice(0, 80);
      }
      if (meta?.thumbnail_url) {
        importedTemplate.preview_image = meta.thumbnail_url;
      }
      const fb = buildCanvaSvgDataUrl(importedTemplate);
      importedTemplate.safe_preview_data_url = fb;
      importedTemplate.fallback_preview_data_url = fb;
      const updated = [importedTemplate, ...fetched.filter((x) => x.id !== importedTemplate!.id)];
      fetched = updated;
      saveFetchedCanva(updated);
    }

    const combinedMap = new Map<string, CanvaTemplate>();
    const sourceList = importedTemplate
      ? [importedTemplate, ...fetched, ...PREPARED_EDU_CANVA_TEMPLATES]
      : [...fetched, ...PREPARED_EDU_CANVA_TEMPLATES];
    for (const item of sourceList) {
      if (!combinedMap.has(item.id)) {
        const fb = item.fallback_preview_data_url || buildCanvaSvgDataUrl(item);
        combinedMap.set(item.id, {
          ...item,
          safe_preview_data_url: item.safe_preview_data_url || fb,
          fallback_preview_data_url: fb,
        });
      }
    }

    let templates = Array.from(combinedMap.values());
    if (cat && cat !== 'all' && cat !== 'الكل') {
      templates = templates.filter((t) => {
        if (cat === 'شواهد وتقارير') {
          return ['شواهد وتوثيق', 'تقارير مدرسية', 'ملفات أداء وإنجاز', 'شواهد وتقارير'].includes(t.category);
        }
        if (cat === 'تصاميمي' || cat === 'تصاميمي المجلوبة') {
          return (
            t.category.includes('تصاميمي') ||
            t.source.includes('مستورد') ||
            t.source.includes('المتصل') ||
            t.source.includes('مرفوع')
          );
        }
        if (cat === 'تعليمي') {
          return ['تقارير مدرسية', 'شواهد وتوثيق', 'ملفات أداء وإنجاز'].includes(t.category);
        }
        if (cat === 'فاخر') {
          return ['royal_double', 'modern_ribbon', 'botanical_frame'].includes(t.layout) || t.category === 'شهادات وتكريم';
        }
        if (cat === 'أعمال') {
          return ['clean_stripe', 'geometric_split'].includes(t.layout) || t.category === 'ملفات أداء وإنجاز';
        }
        if (cat === 'إبداعي') {
          return ['gradient_arch', 'pastel_card'].includes(t.layout);
        }
        return t.category === cat || t.category.includes(cat);
      });
    }
    if (q) {
      templates = templates.filter((t) =>
        `${t.name} ${t.category} ${t.description} ${t.badge_text || ''} ${t.source || ''}`.toLowerCase().includes(q)
      );
    }

    const googleStore = loadGoogleAccounts();
    res.setHeader('Cache-Control', 'no-store');
    res.json({
      success: true,
      templates,
      fetched,
      live_count: liveCount,
      live_fetched_count: liveCount,
      api_error: apiError || undefined,
      accounts: googleStore,
      imported_template: importedTemplate,
      source_note: 'مكتبة كانفا الحقيقية المخبأة فائق السرعة',
    });
  });

  app.post('/api/edu/canva/import_urls', async (req, res) => {
    const body = req.body || {};
    const urlsRaw = body.urls || body.url || '';
    const category = String(body.category || 'تصاميمي المجلوبة').trim();
    const customName = String(body.name || '').trim();

    const urlList: string[] = Array.isArray(urlsRaw)
      ? urlsRaw.map((u) => String(u).trim()).filter(Boolean)
      : String(urlsRaw).split(/[\n,\s]+/).map((u) => u.trim()).filter(Boolean);

    if (!urlList.length) {
      res.status(400).json({ success: false, error: 'يرجى لصق رابط قالب كانفا واحد على الأقل' });
      return;
    }

    const imported = await Promise.all(
      urlList.slice(0, 15).map(async (u, idx) => {
        const nm = customName ? (urlList.length === 1 ? customName : `${customName} (${idx + 1})`) : '';
        const item = makeImportedTemplate(u, nm, category);
        const meta = await fetchCanvaOembedQuick(u);
        if (meta?.title && !nm) item.name = meta.title.slice(0, 80);
        if (meta?.thumbnail_url) item.preview_image = meta.thumbnail_url;
        return item;
      })
    );

    const existing = loadFetchedCanva();
    const ids = new Set(imported.map((t) => t.id));
    saveFetchedCanva([...imported, ...existing.filter((x) => !ids.has(x.id))]);

    res.json({ success: true, templates: imported, count: imported.length });
  });

  app.post('/api/edu/canva/fetch_selected', (req, res) => {
    const body = req.body || {};
    let selected: CanvaTemplate[] = Array.isArray(body.templates) ? body.templates : [];
    const urls: string[] = Array.isArray(body.urls) ? body.urls : [];

    if (!selected.length && urls.length) {
      const urlSet = new Set(urls);
      selected = [...loadFetchedCanva(), ...PREPARED_EDU_CANVA_TEMPLATES].filter(
        (t) => (t.canva_url && urlSet.has(t.canva_url)) || urlSet.has(t.id)
      );
      if (!selected.length) {
        selected = urls.map((u) => makeImportedTemplate(u));
      }
    }

    const now = new Date().toISOString().slice(0, 16).replace('T', ' ');
    const normalized = selected.map((item) => {
      const svgUrl = item.fallback_preview_data_url || buildCanvaSvgDataUrl(item);
      if (item.preview_image) void fetchAndCacheCanvaImage(item.preview_image, 4500);
      return {
        ...item,
        safe_preview_data_url: item.safe_preview_data_url || svgUrl,
        fallback_preview_data_url: svgUrl,
        fetched_at: now,
      };
    });

    const existing = loadFetchedCanva();
    const ids = new Set(normalized.map((x) => x.id));
    const merged = [...normalized, ...existing.filter((x) => !ids.has(x.id))];
    saveFetchedCanva(merged);

    res.json({ success: true, fetched_now: normalized, fetched: merged, count: normalized.length });
  });

  app.get('/api/edu/canva/fetched', (_req, res) => {
    res.json({ success: true, fetched: loadFetchedCanva() });
  });

  app.delete('/api/edu/canva/fetched/:tid', (req, res) => {
    const tid = req.params.tid;
    const remaining = loadFetchedCanva().filter((x) => x.id !== tid);
    saveFetchedCanva(remaining);
    res.json({ success: true, fetched: remaining });
  });

  // Fast In-Memory + Disk Cached Canva Image Proxy with request deduplication
  app.get('/api/edu/canva/proxy_img', async (req, res) => {
    const imgUrl = String(req.query.url || '').trim();
    const tid = String(req.query.tid || '').trim();

    const matchedTpl =
      [...loadFetchedCanva(), ...PREPARED_EDU_CANVA_TEMPLATES].find((t) => t.id === tid || t.preview_image === imgUrl) ||
      PRESENTATION_CANVA_TEMPLATES.find((t) => t.id === tid);

    const sendSvgFallback = () => {
      const tplForSvg = matchedTpl || {
        name: 'قالب كانفا',
        badge_text: 'Canva',
        category: 'تصميم تعليمي',
        colors: {
          primary: '#0f766e',
          secondary: '#09323d',
          accent: '#c89b3c',
          bg: '#f8fafc',
          card_bg: '#ffffff',
          text_dark: '#0f172a',
          text_light: '#ffffff',
        },
      };
      res.setHeader('Content-Type', 'image/svg+xml; charset=utf-8');
      res.setHeader('Cache-Control', 'public, max-age=300');
      res.status(200).send(buildCanvaSvgMarkup(tplForSvg));
    };

    if (!imgUrl.startsWith('https://') && !imgUrl.startsWith('http://')) {
      sendSvgFallback();
      return;
    }

    const cached = await fetchAndCacheCanvaImage(imgUrl, 5000);
    if (cached) {
      res.setHeader('Content-Type', cached.ctype);
      res.setHeader('Cache-Control', 'public, max-age=86400');
      res.status(200).send(cached.buf);
      return;
    }

    sendSvgFallback();
  });

  app.post('/api/edu/canva/import_file', upload.single('file'), (req, res) => {
    const file = req.file;
    if (!file) {
      res.status(400).json({ success: false, error: 'يرجى اختيار ملف تصميم من كانفا (صورة أو PDF أو PPTX)' });
      return;
    }
    const name = String(req.body?.name || file.originalname || 'تصميم كانفا مرفوع').trim();
    const category = String(req.body?.category || 'تصاميمي المجلوبة').trim();
    const id = `canva_file_${Date.now().toString(36)}`;
    const isImage = file.mimetype.startsWith('image/');
    const tpl: CanvaTemplate = {
      id,
      name,
      category,
      source: 'ملف مرفوع',
      layout: 'gradient_arch',
      font_family: 'Cairo',
      heading_font: 'Amiri',
      description: `تم استيراد الملف (${file.originalname}) بنجاح`,
      colors: {
        primary: '#0f766e',
        secondary: '#09323d',
        accent: '#c89b3c',
        bg: '#f8fafc',
        card_bg: '#ffffff',
        text_dark: '#0f172a',
        text_light: '#ffffff',
      },
      badge_text: 'ملف مستورد',
    };
    const svgUrl = buildCanvaSvgDataUrl(tpl);
    tpl.safe_preview_data_url = isImage
      ? `data:${file.mimetype};base64,${file.buffer.toString('base64')}`
      : svgUrl;
    tpl.fallback_preview_data_url = svgUrl;

    const existing = loadFetchedCanva();
    saveFetchedCanva([tpl, ...existing]);
    res.json({ success: true, template: tpl });
  });

  app.post('/api/edu/canva/process_for_merge', (req, res) => {
    const tplItem: CanvaTemplate = req.body?.template || req.body;
    if (!tplItem || (!tplItem.id && !tplItem.name)) {
      res.status(400).json({ success: false, error: 'بيانات القالب المختار مطلوبة للمعالجة' });
      return;
    }
    const base = EDU_CANVA_TEMPLATES.find((t) => t.id === tplItem.id) || tplItem;
    const svgUrl = buildCanvaSvgDataUrl(base);
    const safeDataUrl = base.safe_preview_data_url || svgUrl;
    const processed: CanvaTemplate = {
      ...base,
      safe_preview_data_url: safeDataUrl,
      fallback_preview_data_url: base.fallback_preview_data_url || svgUrl,
      processed: true,
      processed_asset_id: base.id,
      processed_pages: [safeDataUrl],
      processed_format: 'PNG / SVG A4',
    };

    res.json({
      success: true,
      template: processed,
      processed_pdf_url: '',
      processed_pages: processed.processed_pages,
      processed_format: processed.processed_format,
    });
  });

  app.get('/api/edu/canva/asset/:tid/:filename', (req, res) => {
    const safeTid = String(req.params.tid || '').replace(/[^A-Za-z0-9_-]/g, '');
    const filename = String(req.params.filename || '');
    if (!safeTid || !/^(?:page_\d{2}\.png|template\.pdf)$/.test(filename)) {
      res.status(404).end();
      return;
    }
    const assetPath = path.join(PROCESSED_CANVA_DIR, safeTid, filename);
    if (!fs.existsSync(assetPath)) {
      res.status(404).end();
      return;
    }
    res.sendFile(assetPath);
  });

  // ================= API: Presentation & Document Studio (منشئ العروض والمستندات الاحترافي) =================
  app.get('/api/pptx/ai_status', (_req, res) => {
    const ai = getGeminiClient();
    res.json({
      available: true,
      provider: ai ? 'Gemini Flash' : 'المعالج الذكي المحلي',
    });
  });

  app.post('/api/pptx/extract_file', upload.single('file'), async (req, res) => {
    const file = req.file;
    if (!file) {
      res.status(400).json({ error: 'يرجى اختيار ملف أولاً' });
      return;
    }
    const t0 = Date.now();
    try {
      const extracted = await extractFileContent(file.buffer, file.originalname || 'file.txt');
      const ms = Math.max(15, Date.now() - t0);
      const metrics = buildPipelineMetrics({
        operation: `استخراج محتوى (${file.originalname})`,
        stages: [
          { key: 'extract', stage: '1. قراءة واستخلاص النصوص', shortName: 'الاستخراج', duration_ms: ms },
          { key: 'ai', stage: '2. تحليل البنية والجداول', shortName: 'الهيكلة', duration_ms: 14 },
          { key: 'charts', stage: '3. فحص الجداول', shortName: 'الجداول', duration_ms: 10 },
          { key: 'render', stage: '4. تجهيز المحتوى', shortName: 'التجهيز', duration_ms: 8 },
        ],
        wordCount: extracted.word_count,
        itemsCount: Math.max(1, extracted.slides.length || extracted.tables.length || 1),
        provider: 'محرك استخراج المستندات',
      });
      res.json({
        success: true,
        ...extracted,
        pipeline_metrics: metrics,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      res.status(500).json({ error: `تعذر قراءة الملف: ${msg}` });
    }
  });

  app.post('/api/pptx/preview_html_links', (req, res) => {
    const html = String(req.body?.html || '');
    const $ = cheerio.load(html);
    const links: Array<{ url: string; title: string; anchor_text: string }> = [];
    const seen = new Set<string>();
    $('a[href]').each((_i, el) => {
      const url = ($(el).attr('href') || '').trim();
      if (!url || url.startsWith('#') || seen.has(url)) return;
      seen.add(url);
      links.push({
        url,
        title: ($(el).attr('title') || '').trim(),
        anchor_text: $(el).text().replace(/\s+/g, ' ').trim(),
      });
    });
    res.json({ success: true, links });
  });

  app.post('/api/pptx/generate', async (req, res) => {
    const body = req.body || {};
    const text = String(body.text || '').trim();
    const title = String(body.title || '').trim();
    const numSlides = Number(body.num_slides || 7) || 7;
    const ptype = String(body.ptype || 'general');
    const themeColor = String(body.canva_template_id || body.theme_color || 'blue');
    const fontName = String(body.font_name || '').trim();
    const bodyFontSize = Number(body.body_font_size || 22) || 22;
    const extractedTables: string[][][] = Array.isArray(body.extracted_tables) ? body.extracted_tables : [];
    const templateImages: string[] = Array.isArray(body.template_images) ? body.template_images : [];
    const rawCover = body.cover_data && typeof body.cover_data === 'object' ? body.cover_data : null;

    if (!text) {
      res.status(400).json({ error: 'يرجى إدخال نص العرض أو استخراج ملف أولاً' });
      return;
    }

    const t0 = Date.now();
    try {
      const { slides, provider } = await generateSlidesStructure({
        text,
        numSlides,
        ptype,
        titleOverride: title,
      });

      if (body.inc_tables && extractedTables.length > 0) {
        extractedTables.slice(0, 2).forEach((tbl, idx) => {
          slides.push({
            title: `جدول البيانات ${idx + 1}`,
            bullets: [],
            slide_type: 'table',
            table_data: tbl,
          });
        });
      }

      if (body.inc_charts) {
        slides.push({
          title: 'المؤشرات والتمثيل البياني',
          bullets: [],
          slide_type: 'chart',
          chart_title: 'مؤشر الإنجاز',
          chart_labels: ['المحور الأول', 'المحور الثاني', 'المحور الثالث', 'المحور الرابع'],
          chart_values: [88, 76, 94, 85],
        });
      }

      const coverData = rawCover
        ? {
            org: String(rawCover.organization || rawCover.org || ''),
            presenter: String(rawCover.presenter || ''),
            subtitle: String(rawCover.subtitle || ''),
            date: String(rawCover.date || ''),
            logo: String(rawCover.logo || ''),
          }
        : null;
      const allImgs = [
        ...(rawCover?.logo_image ? [String(rawCover.logo_image)] : []),
        ...templateImages,
      ];

      const customTemplates = loadStudioCustomCanva();
      const built = await buildPptxFile({
        outputDir: STUDIO_OUTPUTS_DIR,
        slides,
        themeColor,
        customTemplates,
        coverData,
        templateImages: allImgs,
        fontName,
        bodyFontSize,
        filenamePrefix: title || slides[0]?.title || 'عرض_تقديمي',
      });

      const elapsed = Math.max(40, Date.now() - t0);
      const wordCount = text.split(/\s+/).filter(Boolean).length;
      const metrics = buildPipelineMetrics({
        operation: `إنشاء عرض PowerPoint (${slides.length} شريحة)`,
        stages: [
          { key: 'extract', stage: '1. تحليل النص المدخل', shortName: 'التحليل', duration_ms: Math.round(elapsed * 0.2) },
          { key: 'ai', stage: '2. توليد الشرائح بالذكاء', shortName: 'التوليد', duration_ms: Math.round(elapsed * 0.45) },
          { key: 'charts', stage: '3. تطبيق الألوان والرسوم', shortName: 'التنسيق', duration_ms: 18 },
          { key: 'render', stage: '4. تصدير ملف PPTX', shortName: 'التصدير', duration_ms: Math.round(elapsed * 0.35) },
        ],
        wordCount,
        itemsCount: slides.length,
        provider,
      });

      res.json({
        success: true,
        filename: built.filename,
        download_url: `/api/pptx/download/${encodeURIComponent(built.filename)}`,
        slides,
        num_slides: slides.length,
        theme_applied: built.themeName,
        pipeline_metrics: metrics,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      res.status(500).json({ error: msg });
    }
  });

  app.post('/api/pptx/html_to_pptx', async (req, res) => {
    const body = req.body || {};
    const html = String(body.html || '').trim();
    if (!html) {
      res.status(400).json({ error: 'محتوى HTML فارغ' });
      return;
    }
    const t0 = Date.now();
    try {
      const parsed = parseHtmlDocument(html);
      const title = String(body.title || parsed.title || 'عرض_من_HTML').trim();
      const slides: SlideItem[] = [
        {
          title,
          bullets: parsed.sections[0]?.paragraphs.slice(0, 3) || ['عرض تقديمي مستخرج من HTML'],
          slide_type: 'title',
        },
      ];
      for (const sec of parsed.sections) {
        const bullets = [...sec.bullets, ...sec.paragraphs].slice(0, 6);
        if (bullets.length > 0) {
          slides.push({ title: sec.heading || title, bullets, slide_type: 'content' });
        }
        if (sec.table && sec.table.length > 0) {
          slides.push({
            title: sec.heading || 'جدول البيانات',
            bullets: [],
            slide_type: 'table',
            table_data: sec.table,
          });
        }
      }

      const themeColor = String(body.canva_template_id || body.theme_color || 'blue');
      const customTemplates = loadStudioCustomCanva();
      const rawCover = body.cover_data && typeof body.cover_data === 'object' ? body.cover_data : null;
      const allImgs = [
        ...(rawCover?.logo_image ? [String(rawCover.logo_image)] : []),
        ...(Array.isArray(body.template_images) ? body.template_images : []),
      ];
      const built = await buildPptxFile({
        outputDir: STUDIO_OUTPUTS_DIR,
        slides,
        themeColor,
        customTemplates,
        coverData: rawCover,
        templateImages: allImgs,
        fontName: String(body.font_name || ''),
        bodyFontSize: Number(body.body_font_size || 22) || 22,
        filenamePrefix: title,
      });

      const elapsed = Math.max(30, Date.now() - t0);
      const metrics = buildPipelineMetrics({
        operation: `تحويل HTML إلى PowerPoint (${slides.length} شريحة)`,
        stages: [
          { key: 'extract', stage: '1. تحليل بنية HTML', shortName: 'HTML', duration_ms: 15 },
          { key: 'ai', stage: '2. تحويل الأقسام لشرائح', shortName: 'التحويل', duration_ms: 18 },
          { key: 'charts', stage: '3. بناء الجداول والتنسيق', shortName: 'التنسيق', duration_ms: 16 },
          { key: 'render', stage: '4. تصدير ملف PPTX', shortName: 'التصدير', duration_ms: elapsed },
        ],
        wordCount: cheerio.load(html)('body').text().trim().split(/\s+/).length,
        itemsCount: slides.length,
        provider: 'محرك HTML إلى PPTX',
      });

      res.json({
        success: true,
        filename: built.filename,
        download_url: `/api/pptx/download/${encodeURIComponent(built.filename)}`,
        slides,
        num_slides: slides.length,
        theme_applied: built.themeName,
        pipeline_metrics: metrics,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      res.status(500).json({ error: msg });
    }
  });

  app.get('/api/pptx/download/:filename', (req, res) => {
    const safeName = path.basename(String(req.params.filename || ''));
    const filePath = path.join(STUDIO_OUTPUTS_DIR, safeName);
    if (!safeName || !fs.existsSync(filePath)) {
      res.status(404).json({ error: `الملف غير موجود: ${safeName}` });
      return;
    }
    const ext = path.extname(safeName).toLowerCase();
    res.setHeader('Content-Type', MIME_MAP[ext] || 'application/octet-stream');
    res.download(filePath, safeName);
  });

  app.post('/api/pptx/chat', async (req, res) => {
    const msgs = Array.isArray(req.body?.messages) ? req.body.messages : [];
    const reply = await pptxChatReply(msgs);
    res.json({ success: true, reply });
  });

  // Presentation Studio Canva Templates & Google Accounts
  app.get('/api/pptx/canva/templates', (_req, res) => {
    const custom = loadStudioCustomCanva();
    const templates = [...custom, ...PRESENTATION_CANVA_TEMPLATES].map((t) => {
      const svgUrl = buildCanvaSvgDataUrl(t);
      return {
        ...t,
        safe_preview_data_url: t.safe_preview_data_url || svgUrl,
        fallback_preview_data_url: t.fallback_preview_data_url || svgUrl,
      };
    });
    res.json({
      success: true,
      templates,
      accounts: loadGoogleAccounts(),
      google_oauth_client_id: '',
      firebase_config: {
        apiKey: '',
        authDomain: '',
        projectId: '',
        appId: '',
        oAuthClientId: '',
      },
    });
  });

  app.post('/api/pptx/canva/import_url', async (req, res) => {
    const url = String(req.body?.url || '').trim();
    const customName = String(req.body?.name || '').trim();
    const category = String(req.body?.category || 'تصاميمي').trim();
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      res.status(400).json({ error: 'يرجى إدخال رابط كانفا صحيح يبدأ بـ https://' });
      return;
    }
    const cid = extractCanvaIdFromUrl(url);
    const meta = await fetchCanvaOembedQuick(url);
    const tpl: StudioCanvaTemplate = {
      id: `canva_custom_${cid}`,
      canva_id: cid,
      name: customName || meta?.title?.slice(0, 75) || `قالب كانفا (${cid.slice(0, 10)})`,
      category,
      source: 'رابط كانفا مستورد',
      theme_color: 'blue',
      layout: 'modern_card',
      font_name: 'Cairo',
      description: `قالب مستورد من رابط كانفا: ${url}`,
      colors: {
        primary: '#008B94',
        secondary: '#0F2942',
        accent: '#F59E0B',
        bg: '#F8FAFC',
        card_bg: '#FFFFFF',
        text_dark: '#0F172A',
        text_light: '#FFFFFF',
      },
      canva_url: url,
      preview_image: meta?.thumbnail_url,
    };
    const svgUrl = buildCanvaSvgDataUrl(tpl);
    tpl.safe_preview_data_url = svgUrl;
    tpl.fallback_preview_data_url = svgUrl;

    const existing = loadStudioCustomCanva();
    saveStudioCustomCanva([tpl, ...existing.filter((x) => x.id !== tpl.id)]);
    res.json({ success: true, template: tpl });
  });

  app.post('/api/pptx/canva/connect_token', async (req, res) => {
    const token = String(req.body?.access_token || '').trim();
    if (!token) {
      res.status(400).json({ error: 'رمز Canva Connect API Token مطلوب' });
      return;
    }
    try {
      const apiRes = await fetch('https://api.canva.com/rest/v1/designs?limit=12', {
        signal: AbortSignal.timeout(4000),
        headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
      });
      if (!apiRes.ok) {
        res.status(400).json({ error: `تعذر المصادقة مع Canva API (HTTP ${apiRes.status})` });
        return;
      }
      const json = (await apiRes.json()) as {
        items?: Array<{ id?: string; title?: string; thumbnail?: { url?: string }; urls?: { view_url?: string } }>;
      };
      const items = Array.isArray(json.items) ? json.items : [];
      const imported: StudioCanvaTemplate[] = items.map((d, idx) => {
        const t: StudioCanvaTemplate = {
          id: `canva_user_${d.id || idx}`,
          canva_id: d.id,
          name: d.title || `تصميم كانفا ${idx + 1}`,
          category: 'تصاميمي',
          source: 'Canva Connect',
          theme_color: 'blue',
          layout: 'modern_card',
          font_name: 'Cairo',
          description: 'تصميم مجلوب من حسابك في كانفا',
          colors: {
            primary: '#008B94',
            secondary: '#0F2942',
            accent: '#F59E0B',
            bg: '#F8FAFC',
            card_bg: '#FFFFFF',
            text_dark: '#0F172A',
            text_light: '#FFFFFF',
          },
          canva_url: d.urls?.view_url,
          preview_image: d.thumbnail?.url,
        };
        const svgUrl = buildCanvaSvgDataUrl(t);
        t.safe_preview_data_url = svgUrl;
        t.fallback_preview_data_url = svgUrl;
        return t;
      });
      const existing = loadStudioCustomCanva();
      const ids = new Set(imported.map((i) => i.id));
      saveStudioCustomCanva([...imported, ...existing.filter((x) => !ids.has(x.id))]);
      res.json({ success: true, count: imported.length, templates: imported });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      res.status(500).json({ error: `فشل الاتصال بـ Canva API: ${msg}` });
    }
  });

  app.all('/api/pptx/canva/google_accounts', (req, res) => {
    const store = loadGoogleAccounts();
    if (req.method === 'GET') {
      res.json({ success: true, ...store });
      return;
    }
    const body = req.body || {};
    const action = String(body.action || (req.method === 'DELETE' ? 'remove' : 'add')).trim();
    const email = String(body.email || '').trim();
    if (!email) {
      res.status(400).json({ error: 'البريد الإلكتروني مطلوب' });
      return;
    }

    if (action === 'select') {
      store.active_email = email;
      saveGoogleAccounts(store);
      res.json({ success: true, ...store });
      return;
    }
    if (action === 'remove' || req.method === 'DELETE') {
      store.accounts = store.accounts.filter((a) => a.email.toLowerCase() !== email.toLowerCase());
      if (store.active_email.toLowerCase() === email.toLowerCase()) {
        store.active_email = store.accounts[0]?.email || '';
      }
      saveGoogleAccounts(store);
      res.json({ success: true, ...store });
      return;
    }

    const entry: GoogleAccountEntry = {
      email,
      name: String(body.name || email.split('@')[0]).trim(),
      photo_url: String(body.photo_url || '').trim(),
      provider: String(body.provider || 'manual').trim(),
      added_at: new Date().toISOString(),
    };
    store.accounts = [entry, ...store.accounts.filter((a) => a.email.toLowerCase() !== email.toLowerCase())];
    store.active_email = email;
    saveGoogleAccounts(store);
    res.json({ success: true, ...store });
  });

  app.post('/api/pptx/canva/apply_template', async (req, res) => {
    const body = req.body || {};
    const templateId = String(body.template_id || 'canva_exec_navy').trim();
    const customTemplates = loadStudioCustomCanva();
    const tpl = [...customTemplates, ...PRESENTATION_CANVA_TEMPLATES].find((t) => t.id === templateId);
    const title = String(body.title || tpl?.name || 'عرض احترافي').trim();
    const text = String(body.text || '').trim();
    let slides: SlideItem[] = Array.isArray(body.slides) ? body.slides : [];
    const t0 = Date.now();
    let provider = 'Canva Design Engine';

    if (!slides.length) {
      if (text) {
        const resGen = await generateSlidesStructure({
          text,
          numSlides: Number(body.num_slides || 6) || 6,
          ptype: String(body.ptype || 'general'),
          titleOverride: title,
        });
        slides = resGen.slides;
        provider = resGen.provider;
      } else {
        slides = [
          {
            title,
            bullets: [tpl?.description || 'عرض مصمم بقوالب كانفا المعاصرة'],
            slide_type: 'title',
          },
          {
            title: 'الرؤية والأهداف الاستراتيجية',
            bullets: ['تحديد المخرجات والأهداف الرئيسية بدقة', 'رفع كفاءة العرض البصري وتناسق الهوية'],
            slide_type: 'content',
          },
          {
            title: 'المحاور والنتائج التنفيذية',
            bullets: ['تنفيذ خطة العمل وفق مؤشرات الأداء', 'توثيق المنجزات وعرضها بصورة احترافية'],
            slide_type: 'content',
          },
          {
            title: 'الخلاصة والتوصيات',
            bullets: ['تعميم النموذج ومتابعة التطوير المستمر', 'شكراً لحسن المتابعة'],
            slide_type: 'closing',
          },
        ];
      }
    }

    const rawCover = body.cover_data && typeof body.cover_data === 'object' ? body.cover_data : null;
    const allImgs = [
      ...(rawCover?.logo_image ? [String(rawCover.logo_image)] : []),
      ...(Array.isArray(body.template_images) ? body.template_images : []),
    ];

    const built = await buildPptxFile({
      outputDir: STUDIO_OUTPUTS_DIR,
      slides,
      themeColor: templateId,
      customTemplates,
      coverData: rawCover
        ? {
            org: String(rawCover.organization || rawCover.org || ''),
            presenter: String(rawCover.presenter || ''),
            subtitle: String(rawCover.subtitle || ''),
            date: String(rawCover.date || ''),
            logo: String(rawCover.logo || ''),
          }
        : null,
      templateImages: allImgs,
      fontName: String(body.font_name || ''),
      bodyFontSize: Number(body.body_font_size || 22) || 22,
      filenamePrefix: title,
    });

    const elapsed = Math.max(30, Date.now() - t0);
    const wordCount = slides.map((s) => `${s.title} ${(s.bullets || []).join(' ')}`).join(' ').split(/\s+/).length;
    const metrics = buildPipelineMetrics({
      operation: `تطبيق قالب كانفا (${built.themeName})`,
      stages: [
        { key: 'extract', stage: '1. تحميل الشرائح والهوية', shortName: 'التحميل', duration_ms: 14 },
        { key: 'ai', stage: '2. تنسيق النصوص العربية', shortName: 'التنسيق', duration_ms: 16 },
        { key: 'charts', stage: '3. تطبيق ألوان كانفا', shortName: 'القالب', duration_ms: 18 },
        { key: 'render', stage: '4. تصدير ملف PPTX', shortName: 'التصدير', duration_ms: elapsed },
      ],
      wordCount,
      itemsCount: slides.length,
      provider,
    });

    res.json({
      success: true,
      filename: built.filename,
      download_url: `/api/pptx/download/${encodeURIComponent(built.filename)}`,
      slides,
      num_slides: slides.length,
      applied_design: tpl || { id: templateId, name: built.themeName },
      pipeline_metrics: metrics,
    });
  });

  app.post('/api/pptx/analyze_pptx', upload.single('file'), async (req, res) => {
    const file = req.file;
    if (!file) {
      res.status(400).json({ error: 'يرجى اختيار ملف PowerPoint (.pptx)' });
      return;
    }
    if (!String(file.originalname || '').toLowerCase().endsWith('.pptx')) {
      res.status(400).json({ error: 'يرجى رفع ملف بصيغة .pptx فقط' });
      return;
    }
    const t0 = Date.now();
    try {
      const analyzed = await analyzePptxBuffer(file.buffer);
      const elapsed = Math.max(30, Date.now() - t0);
      const metrics = buildPipelineMetrics({
        operation: `تحليل عرض PowerPoint (${analyzed.slide_count} شريحة)`,
        stages: [
          { key: 'extract', stage: '1. استخراج الشرائح', shortName: 'الاستخراج', duration_ms: Math.round(elapsed * 0.35) },
          { key: 'ai', stage: '2. تحليل الذكاء', shortName: 'التحليل', duration_ms: Math.round(elapsed * 0.45) },
          { key: 'charts', stage: '3. تقييم القوالب', shortName: 'القوالب', duration_ms: 18 },
          { key: 'render', stage: '4. إعداد التوصيات', shortName: 'التوصيات', duration_ms: 12 },
        ],
        wordCount: analyzed.word_count,
        itemsCount: analyzed.slide_count,
        provider: analyzed.provider,
      });
      res.json({
        success: true,
        filename: file.originalname,
        ...analyzed,
        pipeline_metrics: metrics,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      res.status(500).json({ error: `تعذّر تحليل الملف: ${msg}` });
    }
  });

  app.post('/api/pptx/redesign_pptx', async (req, res) => {
    const body = req.body || {};
    const origSlides: SlideItem[] = Array.isArray(body.slides) ? body.slides : [];
    const suggestions: Array<{ slide_number?: number; improved_title?: string; improved_bullets?: string[] }> =
      Array.isArray(body.slide_suggestions) ? body.slide_suggestions : [];
    if (!origSlides.length) {
      res.status(400).json({ error: 'لا توجد شرائح لإعادة تصميمها — قم برفع وتحليل ملف .pptx أولاً' });
      return;
    }

    const useAi = body.use_ai_improvements !== false;
    const templateId = String(body.template_id || 'canva_exec_navy').trim();
    const titleOverride = String(body.title || '').trim();
    const suggMap = new Map(suggestions.map((s, idx) => [Number(s.slide_number || idx + 1), s]));
    const rebuiltSlides: SlideItem[] = origSlides.map((orig, idx) => {
      const copy: SlideItem = { ...orig };
      const sg = suggMap.get(idx + 1);
      if (useAi && sg) {
        if (sg.improved_title) copy.title = sg.improved_title;
        if (Array.isArray(sg.improved_bullets) && sg.improved_bullets.length && copy.slide_type !== 'table') {
          copy.bullets = sg.improved_bullets;
        }
      }
      return copy;
    });
    if (titleOverride && rebuiltSlides[0]) {
      rebuiltSlides[0].title = titleOverride;
    }

    const t0 = Date.now();
    try {
      const customTemplates = loadStudioCustomCanva();
      const built = await buildPptxFile({
        outputDir: STUDIO_OUTPUTS_DIR,
        slides: rebuiltSlides,
        themeColor: templateId,
        customTemplates,
        templateImages: Array.isArray(body.template_images) ? body.template_images : [],
        fontName: String(body.font_name || ''),
        bodyFontSize: Number(body.body_font_size || 22) || 22,
        filenamePrefix: titleOverride || 'عرض_محدث_كانفا',
      });

      const elapsed = Math.max(30, Date.now() - t0);
      const wordCount = rebuiltSlides.map((s) => `${s.title} ${(s.bullets || []).join(' ')}`).join(' ').split(/\s+/).length;
      const metrics = buildPipelineMetrics({
        operation: `إعادة تصميم وتحديث ${rebuiltSlides.length} شريحة`,
        stages: [
          { key: 'extract', stage: '1. تحميل الشرائح', shortName: 'التحميل', duration_ms: 12 },
          { key: 'ai', stage: '2. تطبيق التحسينات', shortName: 'التحسين', duration_ms: 16 },
          { key: 'charts', stage: '3. قالب كانفا', shortName: 'القالب', duration_ms: 18 },
          { key: 'render', stage: '4. تصدير PPTX', shortName: 'التصدير', duration_ms: elapsed },
        ],
        wordCount,
        itemsCount: rebuiltSlides.length,
        provider: 'Canva Design Engine',
      });

      res.json({
        success: true,
        filename: built.filename,
        download_url: `/api/pptx/download/${encodeURIComponent(built.filename)}`,
        slides: rebuiltSlides,
        num_slides: rebuiltSlides.length,
        applied_design: { id: templateId, name: built.themeName },
        pipeline_metrics: metrics,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      res.status(500).json({ error: msg });
    }
  });

  // ================= API: Word / Excel / PDF Studio Conversions =================
  app.post('/api/html_to_word/ai_process', async (req, res) => {
    const html = String(req.body?.html || '').trim();
    if (!html) {
      res.status(400).json({ success: false, error: 'محتوى HTML فارغ' });
      return;
    }
    const action = String(req.body?.action || 'optimize');
    const customPrompt = String(req.body?.custom_prompt || '');
    const t0 = Date.now();
    try {
      const { processed_html, provider } = await aiProcessHtml({ html, action, customPrompt });
      const elapsed = Math.max(25, Date.now() - t0);
      const metrics = buildPipelineMetrics({
        operation: `معالجة HTML بالذكاء (${action})`,
        stages: [
          { key: 'extract', stage: '1. تحليل عناصر HTML', shortName: 'التحليل', duration_ms: 14 },
          { key: 'ai', stage: '2. تحسين التنسيق والهيكل', shortName: 'التحسين', duration_ms: elapsed },
          { key: 'charts', stage: '3. تثبيت الجداول والخطوط', shortName: 'الجداول', duration_ms: 12 },
          { key: 'render', stage: '4. إخراج HTML المنقح', shortName: 'الإخراج', duration_ms: 10 },
        ],
        wordCount: cheerio.load(processed_html)('body').text().trim().split(/\s+/).length,
        itemsCount: 1,
        provider,
      });
      res.json({ success: true, processed_html, pipeline_metrics: metrics });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      res.status(500).json({ success: false, error: msg });
    }
  });

  app.post(['/api/studio/pdf_to_word', '/tools/pdf_to_word'], upload.single('file'), async (req, res) => {
    const file = req.file;
    if (!file) {
      res.status(400).json({ error: 'لم يتم رفع ملف PDF' });
      return;
    }
    const t0 = Date.now();
    try {
      const extracted = await extractFileContent(file.buffer, file.originalname || 'document.pdf');
      const paras = (extracted.full_text || 'محتوى مستخرج من ملف PDF').split(/\n+/).map((p) => p.trim()).filter(Boolean);
      const title = (file.originalname || 'مستند').replace(/\.pdf$/i, '');
      const built = await buildDocxFile({
        outputDir: STUDIO_OUTPUTS_DIR,
        filenamePrefix: title,
        title,
        paragraphs: paras,
        fontFamily: String(req.body?.font_family || 'Cairo').trim() || 'Cairo',
        fontSizePt: Number(req.body?.font_size || 12) || 12,
        marginInches: Number(req.body?.margin || 1.0) || 1.0,
        rtl: true,
      });
      const elapsed = Math.max(30, Date.now() - t0);
      const pagesEst = Math.max(1, Math.ceil(paras.length / 22));
      const metrics = buildPipelineMetrics({
        operation: `تحويل PDF إلى Word (.docx)`,
        stages: [
          { key: 'extract', stage: '1. استخراج نصوص PDF', shortName: 'PDF', duration_ms: Math.round(elapsed * 0.45) },
          { key: 'ai', stage: '2. معالجة الفقرات العربية', shortName: 'المعالجة', duration_ms: 16 },
          { key: 'charts', stage: '3. ضبط الهوامش والخطوط', shortName: 'التنسيق', duration_ms: 14 },
          { key: 'render', stage: '4. تصدير ملف DOCX', shortName: 'DOCX', duration_ms: Math.round(elapsed * 0.45) },
        ],
        wordCount: extracted.word_count,
        itemsCount: pagesEst,
        provider: 'محرك PDF إلى Word',
      });
      res.json({
        success: true,
        filename: built.filename,
        download_url: `/api/studio/download/${encodeURIComponent(built.filename)}`,
        pages: pagesEst,
        pipeline_metrics: metrics,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      res.status(500).json({ error: msg });
    }
  });

  app.post(['/api/studio/html_to_word', '/tools/html_to_word'], async (req, res) => {
    const body = req.body || {};
    const html = String(body.html || '').trim();
    if (!html) {
      res.status(400).json({ error: 'محتوى HTML فارغ' });
      return;
    }
    const t0 = Date.now();
    try {
      const parsed = parseHtmlDocument(html);
      const fname = String(body.filename || parsed.title || 'document').trim();
      const built = await buildDocxFile({
        outputDir: STUDIO_OUTPUTS_DIR,
        filenamePrefix: fname,
        title: parsed.title,
        sections: parsed.sections,
        fontFamily: String(body.font_family || 'Cairo').trim() || 'Cairo',
        fontSizePt: Number(body.font_size || 12) || 12,
        marginInches: Number(body.margin || 1.0) || 1.0,
        rtl: body.rtl !== false,
      });
      const elapsed = Math.max(25, Date.now() - t0);
      const wordCount = cheerio.load(html)('body').text().trim().split(/\s+/).length;
      const metrics = buildPipelineMetrics({
        operation: 'تحويل HTML إلى Word (.docx)',
        stages: [
          { key: 'extract', stage: '1. تحليل مستند HTML', shortName: 'HTML', duration_ms: 15 },
          { key: 'ai', stage: '2. تنسيق الأقسام والفقرات', shortName: 'التنسيق', duration_ms: 14 },
          { key: 'charts', stage: '3. بناء الجداول', shortName: 'الجداول', duration_ms: 12 },
          { key: 'render', stage: '4. تصدير ملف DOCX', shortName: 'DOCX', duration_ms: elapsed },
        ],
        wordCount,
        itemsCount: Math.max(1, parsed.sections.length),
        provider: 'محرك HTML إلى Word',
      });
      res.json({
        success: true,
        filename: built.filename,
        download_url: `/api/studio/download/${encodeURIComponent(built.filename)}`,
        pipeline_metrics: metrics,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      res.status(500).json({ error: msg });
    }
  });

  app.post(['/api/studio/html_to_excel', '/tools/html_to_excel'], (req, res) => {
    const body = req.body || {};
    const html = String(body.html || '').trim();
    if (!html) {
      res.status(400).json({ error: 'محتوى HTML فارغ' });
      return;
    }
    const t0 = Date.now();
    try {
      const built = buildExcelFromHtml({
        outputDir: STUDIO_OUTPUTS_DIR,
        html,
        filenamePrefix: String(body.filename || 'excel_export').trim(),
      });
      const elapsed = Math.max(20, Date.now() - t0);
      const wordCount = cheerio.load(html)('body').text().trim().split(/\s+/).length;
      const metrics = buildPipelineMetrics({
        operation: 'تحويل HTML إلى Excel (.xlsx)',
        stages: [
          { key: 'extract', stage: '1. استخراج جداول HTML', shortName: 'الجداول', duration_ms: 12 },
          { key: 'ai', stage: '2. تنظيف الخلايا والصفوف', shortName: 'المعالجة', duration_ms: 10 },
          { key: 'charts', stage: '3. بناء أوراق العمل', shortName: 'الأوراق', duration_ms: 10 },
          { key: 'render', stage: '4. تصدير مصنف XLSX', shortName: 'XLSX', duration_ms: elapsed },
        ],
        wordCount,
        itemsCount: built.sheetsCount,
        provider: 'محرك HTML إلى Excel',
      });
      res.json({
        success: true,
        filename: built.filename,
        download_url: `/api/studio/download/${encodeURIComponent(built.filename)}`,
        tables_count: built.sheetsCount,
        pipeline_metrics: metrics,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      res.status(500).json({ error: msg });
    }
  });

  app.post('/api/studio/text_to_word', async (req, res) => {
    const body = req.body || {};
    const text = String(body.text || '').trim();
    if (!text) {
      res.status(400).json({ error: 'النص فارغ — يرجى كتابة محتوى أولاً' });
      return;
    }
    const t0 = Date.now();
    try {
      const lines = text.split(/\r?\n+/).map((l) => l.trim()).filter(Boolean);
      const title = String(body.title || body.cover_data?.title || lines[0] || 'مستند احترافي').slice(0, 80);
      const tables: string[][][] = Array.isArray(body.extracted_tables) ? body.extracted_tables : [];
      const built = await buildDocxFile({
        outputDir: STUDIO_OUTPUTS_DIR,
        filenamePrefix: title,
        title,
        sections: [
          { heading: title, paragraphs: lines, bullets: [] },
          ...tables.map((t, i) => ({ heading: `جدول ${i + 1}`, paragraphs: [], bullets: [], table: t })),
        ],
        fontFamily: String(body.font_name || 'Cairo').trim() || 'Cairo',
        fontSizePt: Number(body.font_size || 12) || 12,
        marginInches: Number(body.margin || 1.0) || 1.0,
        rtl: true,
      });
      const elapsed = Math.max(25, Date.now() - t0);
      const metrics = buildPipelineMetrics({
        operation: 'تصدير النص إلى مستند Word (.docx)',
        stages: [
          { key: 'extract', stage: '1. قراءة النص والجداول', shortName: 'القراءة', duration_ms: 12 },
          { key: 'ai', stage: '2. تنسيق الفقرات العربية', shortName: 'التنسيق', duration_ms: 14 },
          { key: 'charts', stage: '3. تطبيق الخطوط والهوامش', shortName: 'الخطوط', duration_ms: 10 },
          { key: 'render', stage: '4. تصدير ملف DOCX', shortName: 'DOCX', duration_ms: elapsed },
        ],
        wordCount: text.split(/\s+/).length,
        itemsCount: lines.length,
        provider: 'محرك Word الاحترافي',
      });
      res.json({
        success: true,
        filename: built.filename,
        download_url: `/api/studio/download/${encodeURIComponent(built.filename)}`,
        pipeline_metrics: metrics,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      res.status(500).json({ error: msg });
    }
  });

  app.get(['/api/studio/download/:filename', '/tools/word/download/:filename'], (req, res) => {
    const safeName = path.basename(String(req.params.filename || ''));
    const filePath = path.join(STUDIO_OUTPUTS_DIR, safeName);
    if (!safeName || !fs.existsSync(filePath)) {
      res.status(404).json({ error: `الملف غير موجود أو فارغ: ${safeName}` });
      return;
    }
    const ext = path.extname(safeName).toLowerCase();
    res.setHeader('Content-Type', MIME_MAP[ext] || 'application/octet-stream');
    res.download(filePath, safeName);
  });

  app.get('/api/studio/media/:filename', (req, res) => {
    const safeName = path.basename(String(req.params.filename || ''));
    const filePath = path.join(STUDIO_OUTPUTS_DIR, safeName);
    if (!safeName || !fs.existsSync(filePath)) {
      res.status(404).json({ error: `الملف غير موجود: ${safeName}` });
      return;
    }
    const ext = path.extname(safeName).toLowerCase();
    res.setHeader('Content-Type', MIME_MAP[ext] || 'application/octet-stream');
    res.sendFile(filePath);
  });

  app.get('/api/studio/audio/:filename', (req, res) => {
    const safeName = path.basename(String(req.params.filename || ''));
    const filePath = path.join(STUDIO_AUDIO_DIR, safeName);
    if (!safeName || !fs.existsSync(filePath)) {
      res.status(404).json({ error: 'ملف الصوت غير موجود' });
      return;
    }
    const ext = path.extname(safeName).toLowerCase();
    res.setHeader('Content-Type', MIME_MAP[ext] || 'audio/wav');
    res.sendFile(filePath);
  });

  // ================= API: HTML to Video & Audio Studio =================
  app.get('/api/studio/html_to_video/example', (_req, res) => {
    res.json({ success: true, html: EXAMPLE_VIDEO_HTML });
  });

  app.post('/api/studio/html_to_video/extract_notes', upload.single('file'), async (req, res) => {
    const file = req.file;
    if (!file) {
      res.status(400).json({ error: 'يرجى اختيار ملف ملاحظات أولاً' });
      return;
    }
    try {
      const extracted = await extractFileContent(file.buffer, file.originalname || 'notes.txt');
      if (!extracted.full_text) {
        res.status(400).json({ error: 'لم يتم العثور على نص مقروء داخل الملف المرفق' });
        return;
      }
      res.json({ success: true, text: extracted.full_text, filename: file.originalname });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      res.status(500).json({ error: msg });
    }
  });

  app.post('/api/studio/html_to_video/ai_generate', async (req, res) => {
    const body = req.body || {};
    const out = await aiGenerateVideoHtml({
      prompt: String(body.prompt || ''),
      notes: String(body.notes || ''),
      width: Number(body.width || 1080) || 1080,
      height: Number(body.height || 1080) || 1080,
      duration: Number(body.duration || 10) || 10,
    });
    res.json({ success: true, ...out });
  });

  app.post('/api/studio/html_to_video/ai_refine', async (req, res) => {
    const body = req.body || {};
    const out = await aiRefineVideoHtml({
      html: String(body.html || ''),
      notes: String(body.notes || ''),
      width: Number(body.width || 1080) || 1080,
      height: Number(body.height || 1080) || 1080,
      duration: Number(body.duration || 10) || 10,
    });
    res.json({ success: true, ...out });
  });

  app.post('/api/studio/html_to_video/upload_audio', upload.single('file'), async (req, res) => {
    const file = req.file;
    if (!file) {
      res.status(400).json({ error: 'يرجى اختيار ملف صوتي من الهاتف أو الجهاز أولاً' });
      return;
    }
    try {
      const saved = await saveUploadedStudioAudio({
        audioDir: STUDIO_AUDIO_DIR,
        fileBuffer: file.buffer,
        originalName: file.originalname || 'audio.mp3',
      });
      res.json({ success: true, ...saved });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      res.status(500).json({ error: msg });
    }
  });

  app.post('/api/studio/html_to_video/generate_audio', async (req, res) => {
    const body = req.body || {};
    try {
      const generated = await generateStudioAudioPreview({
        audioDir: STUDIO_AUDIO_DIR,
        html: String(body.html || ''),
        duration: Number(body.duration || 10) || 10,
        autoAudioType: String(body.autoAudioType || body.auto_audio_type || 'voice_and_music'),
        voiceName: String(body.voiceName || body.voice_name || 'Kore'),
        voiceScript: String(body.voiceScript || body.voice_script || ''),
        notes: String(body.notes || ''),
      });
      res.json({ success: true, ...generated });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      res.status(500).json({ error: msg });
    }
  });

  app.post('/api/studio/html_to_video/start', async (req, res) => {
    const body = req.body || {};
    const html = String(body.html || '').trim();
    if (!html) {
      res.status(400).json({ error: 'محتوى HTML فارغ' });
      return;
    }
    const audioMode = String(body.audioMode || body.audio_mode || 'none');
    const uploadedAudioId = String(body.uploadedAudioId || body.uploaded_audio_id || '').trim();
    const pregeneratedAudioId = String(body.pregeneratedAudioId || body.pregenerated_audio_id || '').trim();
    const duration = Number(body.duration || 10) || 10;

    let audioFilePath: string | undefined;
    if (audioMode === 'upload' && uploadedAudioId) {
      const cand = path.join(STUDIO_AUDIO_DIR, path.basename(uploadedAudioId));
      if (fs.existsSync(cand)) audioFilePath = cand;
    } else if (audioMode === 'auto_ai') {
      if (pregeneratedAudioId) {
        const cand = path.join(STUDIO_AUDIO_DIR, path.basename(pregeneratedAudioId));
        if (fs.existsSync(cand)) audioFilePath = cand;
      }
      if (!audioFilePath) {
        try {
          const gen = await generateStudioAudioPreview({
            audioDir: STUDIO_AUDIO_DIR,
            html,
            duration,
            autoAudioType: String(body.autoAudioType || 'voice_and_music'),
            voiceName: String(body.voiceName || 'Kore'),
            voiceScript: String(body.voiceScript || ''),
            notes: String(body.aiNotes || ''),
          });
          audioFilePath = path.join(STUDIO_AUDIO_DIR, path.basename(gen.audio_id));
        } catch {
          // continue without audio if generation fails
        }
      }
    }

    const jobId = await startHtmlToVideoRender({
      outputDir: STUDIO_OUTPUTS_DIR,
      audioDir: STUDIO_AUDIO_DIR,
      html,
      filename: String(body.filename || 'studio_video'),
      width: Number(body.width || 1080) || 1080,
      height: Number(body.height || 1080) || 1080,
      fps: Number(body.fps || 30) || 30,
      duration,
      format: String(body.format || 'mp4'),
      audioFilePath,
      audioFitMode: String(body.audioFitMode || 'fit_pad_trim'),
      audioVolume: Number(body.audioVolume || 1.0) || 1.0,
    });
    res.json({ success: true, job_id: jobId });
  });

  app.get('/api/studio/html_to_video/status/:job_id', (req, res) => {
    const st = getVideoJob(String(req.params.job_id || ''));
    if (!st) {
      res.status(404).json({ error: 'المهمة غير موجودة أو انتهت صلاحيتها' });
      return;
    }
    res.json(st);
  });

  app.post(['/api/studio/html_to_video', '/api/render'], async (req, res) => {
    const body = req.body || {};
    const html = String(body.html || '').trim();
    if (!html) {
      res.status(400).json({ error: "الحقل 'html' مطلوب." });
      return;
    }
    const jobId = await startHtmlToVideoRender({
      outputDir: STUDIO_OUTPUTS_DIR,
      audioDir: STUDIO_AUDIO_DIR,
      html,
      filename: String(body.filename || 'studio_video'),
      width: Number(body.width || 1280) || 1280,
      height: Number(body.height || 720) || 720,
      fps: Number(body.fps || 30) || 30,
      duration: Number(body.duration || 6) || 6,
      format: String(body.format || 'mp4'),
    });
    for (let i = 0; i < 60; i++) {
      const st = getVideoJob(jobId);
      if (st?.status === 'done') {
        res.json({ success: true, ...st });
        return;
      }
      if (st?.status === 'error') {
        res.status(500).json({ error: st.error || 'فشل تصيير الفيديو' });
        return;
      }
      await new Promise((r) => setTimeout(r, 500));
    }
    res.json({ success: true, job_id: jobId, status: 'running' });
  });

  // Mount Vite or Static Dist
  if (!IS_PROD) {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distDir = path.join(__dirname, 'dist');
    app.use(express.static(distDir));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distDir, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
    // Pre-warm all Canva template preview images in background for instant gallery previews
    void warmCanvaImageCache();
  });
}

startServer();
