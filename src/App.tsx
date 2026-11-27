import React, { useEffect, useRef, useState } from 'react';
import {
  Award,
  BookOpen,
  CalendarCheck,
  ClipboardList,
  ExternalLink,
  Eye,
  FileCode2,
  FileDown,
  FilePlus2,
  FileSignature,
  FileSpreadsheet,
  Film,
  FolderKanban,
  FolderPlus,
  GraduationCap,
  Layers,
  LayoutGrid,
  Palette,
  Presentation,
  Printer,
  RotateCcw,
  Save,
  Sparkles,
  Trash2,
  Upload,
  X,
} from 'lucide-react';

type EduToolId =
  | 'performance_portfolio'
  | 'performance_witnesses'
  | 'school_reports'
  | 'exam_editor'
  | 'activities_portfolio'
  | 'presentation_studio';

interface EduToolConfig {
  id: EduToolId;
  title: string;
  shortTitle: string;
  description: string;
  path: string;
  accentHex: string;
  headerGradient: string;
  supportsPrintActions: boolean;
  icon: React.ComponentType<{ className?: string }>;
  subLinks?: Array<{ label: string; path: string }>;
}

const PRESENTATION_STUDIO_TOOL: EduToolConfig = {
  id: 'presentation_studio',
  title: 'منشئ العروض والمستندات الاحترافي',
  shortTitle: 'منشئ العروض والمستندات',
  description: 'PPTX · Word · Excel · قوالب كانفا والمعاينة الحية · MP4',
  path: '/presentation_studio?tab=text',
  accentHex: '#d24726',
  headerGradient: 'linear-gradient(135deg, #d24726, #991b1b)',
  supportsPrintActions: false,
  icon: Layers,
  subLinks: [
    { label: 'قوالب كانفا والمعاينة الحية', path: '/presentation_studio?tab=canva' },
    { label: 'تحليل وتحديث PPTX', path: '/presentation_studio?tab=analyze' },
    { label: 'محول Word و Excel', path: '/presentation_studio?tab=docs' },
    { label: 'محول HTML إلى فيديو MP4', path: '/presentation_studio?tab=video' },
  ],
};

const EDU_TOOLS: EduToolConfig[] = [
  {
    id: 'performance_portfolio',
    title: 'ملف الأداء الوظيفي',
    shortTitle: 'ملف الأداء',
    description: 'تطبيق تفاعلي كامل قابل للتعديل والحفظ وتصدير PDF والطباعة',
    path: '/performance_portfolio',
    accentHex: '#0f766e',
    headerGradient: 'linear-gradient(135deg, #3b6978, #62929e)',
    supportsPrintActions: true,
    icon: ClipboardList,
  },
  {
    id: 'performance_witnesses',
    title: 'شواهد الأداء',
    shortTitle: 'شواهد الأداء',
    description: 'سجل توثيق شواهد الأداء الوظيفي بالصور والروابط وباركود QR',
    path: '/performance_witnesses',
    accentHex: '#6d5ca0',
    headerGradient: 'linear-gradient(135deg, #6d5ca0, #8b7cb5)',
    supportsPrintActions: true,
    icon: Award,
  },
  {
    id: 'school_reports',
    title: 'ملف التقارير والبرامج التربوية والأنشطة المدرسية',
    shortTitle: 'التقارير والشواهد',
    description: 'تقارير البرامج والأنشطة بالصور وباركود QR وإدارة الصفحات وتصدير PDF',
    path: '/school_reports',
    accentHex: '#1e3a8a',
    headerGradient: 'linear-gradient(135deg, #1e3a8a, #3b82f6)',
    supportsPrintActions: true,
    icon: BookOpen,
  },
  {
    id: 'exam_editor',
    title: 'محرر الاختبارات أنور الصبري',
    shortTitle: 'محرر الاختبارات',
    description: 'إنشاء وتحرير الاختبارات المدرسية بالقوالب الجاهزة وتحليل النتائج وتصدير PDF',
    path: '/exam_editor/',
    accentHex: '#1a5276',
    headerGradient: 'linear-gradient(135deg, #1a5276, #2874a6)',
    supportsPrintActions: false,
    icon: FileSignature,
  },
  {
    id: 'activities_portfolio',
    title: 'ملف الأنشطة المدرسية وبطاقات التنفيذ',
    shortTitle: 'الأنشطة والتنفيذ',
    description:
      'بوابة توثيق النشاط الطلابي (مدرسة آمنة بنت عفان - 1448هـ): خطة الأنشطة المدرسية للأسابيع 1–19 وبطاقات تنفيذ الأنشطة المعبأة',
    path: '/activities_portfolio',
    accentHex: '#0f4c5c',
    headerGradient: 'linear-gradient(135deg, #09323d, #0f4c5c)',
    supportsPrintActions: true,
    icon: FolderKanban,
    subLinks: [
      { label: 'ملف الأنشطة المدرسية', path: '/activities_portfolio' },
      { label: 'بطاقة تنفيذ نشاط', path: '/activities_portfolio?tab=cards' },
    ],
  },
];

export default function App() {
  const [allServicesOpen, setAllServicesOpen] = useState(false);
  const [activeTool, setActiveTool] = useState<EduToolConfig | null>(null);
  const [activeToolPath, setActiveToolPath] = useState<string>('');
  const [activitiesView, setActivitiesView] = useState<'plan' | 'cards'>('plan');
  const [toolStatus, setToolStatus] = useState<string>('جاهز للتعديل');
  const [toolModel, setToolModel] = useState<'primary' | 'secondary'>('primary');
  const [isPreviewActive, setIsPreviewActive] = useState<boolean>(false);
  const [mergedTemplateName, setMergedTemplateName] = useState<string>('');
  const iframeRef = useRef<HTMLIFrameElement | null>(null);

  useEffect(() => {
    const readMergedTemplate = () => {
      try {
        const raw = localStorage.getItem('ecs_merged_template_reports');
        const parsed = raw ? JSON.parse(raw) : null;
        if (parsed && typeof parsed.name === 'string' && parsed.name.trim()) {
          setMergedTemplateName(parsed.name.trim());
        } else {
          setMergedTemplateName('');
        }
      } catch {
        setMergedTemplateName('');
      }
    };

    readMergedTemplate();
    window.addEventListener('storage', readMergedTemplate);
    return () => window.removeEventListener('storage', readMergedTemplate);
  }, [allServicesOpen, activeTool]);

  useEffect(() => {
    let cancelled = false;
    const warmCanvaCatalog = async () => {
      try {
        const res = await fetch('/api/edu/canva/browse', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ category: 'all', source: 'cached' }),
        });
        if (!res.ok || cancelled) return;
        const data = await res.json();
        const templates = Array.isArray(data?.templates) ? data.templates : [];
        const fetched = Array.isArray(data?.fetched) ? data.fetched : [];
        if (!templates.length || cancelled) return;

        try {
          const slim = (arr: Array<Record<string, unknown>>) =>
            arr.slice(0, 40).map((item) => ({
              ...item,
              safe_preview_data_url:
                typeof item.safe_preview_data_url === 'string' &&
                /^data:image\/(jpeg|jpg|png|webp|gif)/i.test(item.safe_preview_data_url)
                  ? item.safe_preview_data_url
                  : '',
              fallback_preview_data_url: '',
            }));
          localStorage.setItem(
            'edu_canva_catalog_cache_v2',
            JSON.stringify({
              templates: slim(templates),
              fetched: slim(fetched),
              updatedAt: Date.now(),
            })
          );
        } catch {
          // Ignore localStorage quota errors
        }

        templates.slice(0, 10).forEach((tpl: { thumb_url?: string; id?: string }) => {
          const src = tpl?.thumb_url || (tpl?.id ? `/api/edu/canva/thumb/${encodeURIComponent(tpl.id)}` : '');
          if (src) {
            const img = new Image();
            img.decoding = 'async';
            img.src = src;
          }
        });
      } catch {
        // Silent background warmup
      }
    };
    const timer = setTimeout(() => {
      void warmCanvaCatalog();
    }, 350);
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, []);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        if (activeTool) {
          setActiveTool(null);
        } else if (allServicesOpen) {
          setAllServicesOpen(false);
        }
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [activeTool, allServicesOpen]);

  useEffect(() => {
    const handleMessage = (ev: MessageEvent) => {
      if (ev?.data?.type === 'schoolActivitiesViewChanged') {
        const nextView = ev.data.view === 'cards' ? 'cards' : 'plan';
        setActivitiesView(nextView);
        setActiveToolPath(
          nextView === 'cards' ? '/activities_portfolio?view=cards' : '/activities_portfolio?view=plan'
        );
      } else if (ev?.data?.type === 'eduToolState') {
        if (typeof ev.data.status === 'string' && ev.data.status.trim()) {
          setToolStatus(ev.data.status.trim());
        }
        if (ev.data.model === 'primary' || ev.data.model === 'secondary') {
          setToolModel(ev.data.model);
        }
        if (typeof ev.data.preview === 'boolean') {
          setIsPreviewActive(ev.data.preview);
        }
        if (ev.data.view === 'plan' || ev.data.view === 'cards') {
          setActivitiesView(ev.data.view);
        }
      }
    };
    window.addEventListener('message', handleMessage);

    const win = window as unknown as Record<string, unknown>;
    const printViaParentHost = (
      stylesHtml: string,
      bodyHtml: string,
      onUnavailable?: () => void,
      options?: { rootCssText?: string; hostClassName?: string }
    ) => {
      const host = document.getElementById('eduDirectPrintHost');
      if (!host) {
        if (typeof onUnavailable === 'function') onUnavailable();
        return;
      }
      const scopedStyles = String(stylesHtml || '')
        .replace(/:root\b/g, '#eduDirectPrintHost')
        .replace(/html\.embedded-mode/gi, '#eduDirectPrintHost')
        .replace(/html\s*,\s*body/gi, '#eduDirectPrintHost')
        .replace(/\bhtml\b(?![\w-])/gi, '#eduDirectPrintHost')
        .replace(/\bbody\b(?![\w-])/gi, '#eduDirectPrintHost')
        .replace(/(^|[}\n;>])\s*\*\s*\{/g, '$1#eduDirectPrintHost, #eduDirectPrintHost * {');

      host.className = options?.hostClassName ? String(options.hostClassName) : '';
      host.style.cssText = options?.rootCssText ? String(options.rootCssText) : '';
      host.innerHTML = scopedStyles + String(bodyHtml || '');
      document.body.classList.add('edu-direct-printing');

      let topBeforePrintFired = false;
      let safetyTimer: ReturnType<typeof setTimeout> | null = null;
      const cleanup = () => {
        if (safetyTimer) {
          clearTimeout(safetyTimer);
          safetyTimer = null;
        }
        document.body.classList.remove('edu-direct-printing');
        host.className = '';
        host.style.cssText = '';
        host.innerHTML = '';
        window.removeEventListener('beforeprint', onTopBeforePrint);
        window.removeEventListener('afterprint', cleanup);
      };
      const onTopBeforePrint = () => {
        topBeforePrintFired = true;
      };
      window.addEventListener('beforeprint', onTopBeforePrint, { once: true });
      window.addEventListener('afterprint', cleanup, { once: true });

      requestAnimationFrame(() => {
        try {
          window.focus();
          window.print();
        } catch {
          // Ignore blocked print
        }
        setTimeout(() => {
          if (!topBeforePrintFired) {
            cleanup();
            if (typeof onUnavailable === 'function') onUnavailable();
          } else {
            safetyTimer = setTimeout(cleanup, 60000);
          }
        }, 500);
      });
    };

    win.printActivitiesViaParent = printViaParentHost;
    win.printEduHtmlViaParent = printViaParentHost;

    return () => {
      window.removeEventListener('message', handleMessage);
      delete win.printActivitiesViaParent;
      delete win.printEduHtmlViaParent;
    };
  }, []);

  const openTool = (tool: EduToolConfig, overridePath?: string) => {
    setAllServicesOpen(false);
    setActiveTool(tool);
    setToolStatus('جاهز للتعديل');
    setToolModel('primary');
    setIsPreviewActive(false);
    const targetPath = overridePath || tool.path;
    if (tool.id === 'activities_portfolio') {
      const isCards = targetPath.includes('tab=cards') || targetPath.includes('view=cards');
      setActivitiesView(isCards ? 'cards' : 'plan');
      setActiveToolPath(isCards ? '/activities_portfolio?view=cards' : '/activities_portfolio?view=plan');
    } else {
      setActiveToolPath(targetPath);
    }
  };

  const switchActivitiesSubview = (viewName: 'plan' | 'cards') => {
    setActivitiesView(viewName);
    const frame = iframeRef.current;
    if (frame?.contentWindow) {
      try {
        const win = frame.contentWindow as unknown as Record<string, unknown>;
        if (typeof win.switchView === 'function') {
          (win.switchView as (v: string, skip?: boolean) => void)(viewName, true);
          return;
        }
      } catch {
        // fallback to reloading iframe URL
      }
    }
    setActiveToolPath(
      viewName === 'cards' ? '/activities_portfolio?view=cards' : '/activities_portfolio?view=plan'
    );
  };

  type FrameActionName =
    | 'saveData'
    | 'loadData'
    | 'addNewPage'
    | 'addNewSection'
    | 'openCanvaHubModal'
    | 'openEduCanvaStudio'
    | 'togglePrintPreview'
    | 'exportPDF'
    | 'exportActivitiesPDF'
    | 'exportHTML'
    | 'exportActivitiesHTML'
    | 'printDocument'
    | 'printActivities'
    | 'resetAll'
    | 'openModelUploadModal'
    | 'switchReportsModel'
    | 'switchActivitiesModel';

  const callFrameAction = (fnName: FrameActionName, ...args: unknown[]) => {
    const frame = iframeRef.current;
    if (!frame || !frame.contentWindow) return;
    try {
      const win = frame.contentWindow as unknown as Record<string, unknown>;
      const fn = win[fnName];
      if (typeof fn === 'function') {
        (fn as (...a: unknown[]) => void)(...args);
      } else if (fnName === 'openEduCanvaStudio' && win.EduCanvaStudio && typeof (win.EduCanvaStudio as Record<string, unknown>).toggle === 'function') {
        ((win.EduCanvaStudio as Record<string, unknown>).toggle as (scope: string) => void)('reports');
      } else if (fnName === 'printActivities' || fnName === 'printDocument') {
        frame.contentWindow.focus();
        frame.contentWindow.print();
      }
    } catch {
      // Cross-frame safe fallback
    }
  };

  const switchToolModel = (nextModel: 'primary' | 'secondary') => {
    setToolModel(nextModel);
    if (!activeTool) return;
    if (activeTool.id === 'school_reports') {
      callFrameAction('switchReportsModel', nextModel);
    } else if (activeTool.id === 'activities_portfolio') {
      callFrameAction('switchActivitiesModel', nextModel);
    }
  };

  const standaloneHref = (targetPath: string) => {
    const sep = targetPath.includes('?') ? '&' : '?';
    return `${targetPath}${sep}standalone=1`;
  };

  return (
    <div className="min-h-screen bg-[#f8fafc] text-slate-900 flex flex-col">
      {/* ================= Top Navigation Bar (Single-Row 3-Zone Contract) ================= */}
      <header className="sticky top-0 z-30 bg-white/95 backdrop-blur border-b border-slate-200">
        <div className="max-w-[1200px] mx-auto px-6 py-3.5 flex items-center justify-between gap-8">
          <a
            href="/"
            className="text-lg font-extrabold tracking-tight text-[#09323d] whitespace-nowrap shrink-0 focus-visible:outline-2 focus-visible:outline-[#0f766e] rounded"
          >
            وظائف الخدمات التعليمية
          </a>

          <nav className="hidden lg:flex items-center gap-6 text-sm font-semibold text-slate-600" aria-label="التنقل السريع للأدوات">
            {EDU_TOOLS.map((tool) => (
              <button
                key={tool.id}
                type="button"
                onClick={() => openTool(tool)}
                className="hover:text-[#0f766e] transition-colors whitespace-nowrap shrink-0 cursor-pointer py-1 border-b-2 border-transparent hover:border-[#0f766e] focus-visible:outline-2 focus-visible:outline-[#0f766e]"
              >
                {tool.shortTitle}
              </button>
            ))}
            <button
              type="button"
              onClick={() => openTool(PRESENTATION_STUDIO_TOOL, '/presentation_studio?tab=text')}
              className="hover:text-[#d24726] transition-colors whitespace-nowrap shrink-0 cursor-pointer py-1 border-b-2 border-transparent hover:border-[#d24726] focus-visible:outline-2 focus-visible:outline-[#d24726]"
            >
              {PRESENTATION_STUDIO_TOOL.shortTitle}
            </button>
          </nav>

          <div className="flex items-center shrink-0">
            <button
              type="button"
              id="btnEduServicesTools"
              onClick={() => setAllServicesOpen(true)}
              className="px-4 py-2 text-sm font-bold text-white bg-[#0f766e] rounded-lg hover:bg-[#0d5f58] transition-colors whitespace-nowrap shrink-0 cursor-pointer inline-flex items-center gap-2 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-[#0f766e]"
            >
              <LayoutGrid className="w-4 h-4" />
              <span>فتح الكل</span>
            </button>
          </div>
        </div>
      </header>

      {/* ================= Main Educational Services Hub Workspace ================= */}
      <main className="flex-1 max-w-[1200px] w-full mx-auto px-6 py-10 space-y-10">
        {/* Hero Header */}
        <section className="flex flex-col md:flex-row md:items-end justify-between gap-6 pb-8 border-b border-slate-200">
          <div className="space-y-3 max-w-2xl">
            <div className="flex items-center gap-2 text-xs font-bold text-[#0f766e]">
              <span>منصة التوثيق والتقارير التربوية</span>
              <span aria-hidden="true">·</span>
              <span>العام الدراسي 1448هـ</span>
            </div>
            <h1 className="text-3xl md:text-4xl font-extrabold text-[#09323d] leading-snug text-balance">
              وظائف الخدمات التعليمية والأنشطة المدرسية
            </h1>
            <p className="text-base text-slate-600 leading-relaxed">
              أدوات وملفات تعليمية متكاملة جاهزة للتحرير الفوري، الحفظ المحلي، إضافة الشواهد وباركود QR، والطباعة وتصدير ملفات PDF.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-3 shrink-0">
            <button
              type="button"
              onClick={() => openTool(EDU_TOOLS[4], '/activities_portfolio')}
              className="px-5 py-2.5 text-sm font-bold text-white bg-[#0f4c5c] rounded-lg hover:bg-[#09323d] transition-colors whitespace-nowrap shrink-0 cursor-pointer inline-flex items-center gap-2"
            >
              <FolderKanban className="w-4 h-4 text-[#f3d796]" />
              <span>ملف الأنشطة المدرسية</span>
            </button>
            <button
              type="button"
              onClick={() => openTool(EDU_TOOLS[4], '/activities_portfolio?tab=cards')}
              className="px-5 py-2.5 text-sm font-bold text-[#09323d] bg-[#fdf7ea] border border-[#c89b3c] rounded-lg hover:bg-[#faefd6] transition-colors whitespace-nowrap shrink-0 cursor-pointer inline-flex items-center gap-2"
            >
              <CalendarCheck className="w-4 h-4 text-[#b45309]" />
              <span>بطاقة تنفيذ نشاط</span>
            </button>
          </div>
        </section>

        {/* Iconic Educational Services Card Container */}
        <section
          aria-labelledby="edu-hub-heading"
          className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm"
        >
          <div
            className="px-6 py-5 text-white flex flex-wrap items-center justify-between gap-4"
            style={{ background: 'linear-gradient(135deg, #0f766e, #0d9488)' }}
          >
            <div className="flex items-center gap-3.5">
              <span className="w-11 h-11 rounded-xl bg-white/15 flex items-center justify-center shrink-0">
                <GraduationCap className="w-6 h-6 text-[#fcd34d]" />
              </span>
              <div>
                <h2 id="edu-hub-heading" className="text-lg font-extrabold text-white">
                  وظائف الخدمات التعليمية
                </h2>
                <p className="text-xs font-medium text-teal-100 mt-0.5">
                  أدوات وملفات تعليمية جاهزة للتحرير والطباعة والتصدير
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={() => setAllServicesOpen(true)}
              className="px-4 py-2 bg-white text-slate-900 rounded-lg text-xs font-extrabold hover:bg-teal-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer inline-flex items-center gap-2"
            >
              <LayoutGrid className="w-4 h-4 text-[#0f766e]" />
              <span>فتح الكل</span>
            </button>
          </div>

          <div className="p-6 bg-slate-50/70">
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
              {EDU_TOOLS.map((tool) => {
                const IconComponent = tool.icon;
                const isActivities = tool.id === 'activities_portfolio';
                return (
                  <div
                    key={tool.id}
                    className={`bg-white rounded-xl border border-slate-200 p-5 flex flex-col justify-between gap-4 transition-transform duration-150 hover:-translate-y-0.5 ${
                      isActivities ? 'md:col-span-2 lg:col-span-2' : ''
                    }`}
                  >
                    <div className="space-y-3">
                      <div className="flex items-start justify-between gap-3">
                        <button
                          type="button"
                          onClick={() => openTool(tool)}
                          className="text-right flex items-center gap-3 group cursor-pointer focus-visible:outline-2 focus-visible:outline-[#0f766e] rounded"
                        >
                          <span
                            className="w-12 h-12 rounded-xl flex items-center justify-center shrink-0"
                            style={{
                              backgroundColor: `${tool.accentHex}14`,
                              color: tool.accentHex,
                            }}
                          >
                            <IconComponent className="w-6 h-6" />
                          </span>
                          <span className="text-base font-extrabold text-slate-900 group-hover:text-[#0f766e] transition-colors">
                            {tool.title}
                          </span>
                        </button>
                      </div>

                      <p className="text-sm text-slate-600 leading-relaxed">
                        {tool.description}
                      </p>

                      {tool.id === 'school_reports' && mergedTemplateName && (
                        <div className="text-xs font-bold text-[#0f766e]">
                          قالب كانفا المدمج: {mergedTemplateName}
                        </div>
                      )}
                    </div>

                    <div className="pt-3 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2">
                      {tool.subLinks ? (
                        <div className="flex flex-wrap items-center gap-2">
                          {tool.subLinks.map((sub) => (
                            <button
                              key={sub.path}
                              type="button"
                              onClick={() => openTool(tool, sub.path)}
                              className="px-3.5 py-2 text-xs font-bold rounded-lg bg-[#eef6f8] text-[#09323d] border border-[#9fb8c0] hover:bg-[#0f4c5c] hover:text-white transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                            >
                              {sub.label}
                            </button>
                          ))}
                        </div>
                      ) : (
                        <button
                          type="button"
                          onClick={() => openTool(tool)}
                          className="px-4 py-2 text-xs font-bold rounded-lg text-white transition-opacity hover:opacity-95 whitespace-nowrap shrink-0 cursor-pointer"
                          style={{ backgroundColor: tool.accentHex }}
                        >
                          فتح الأداة
                        </button>
                      )}

                      <a
                        href={standaloneHref(tool.path)}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="px-3 py-2 text-xs font-bold text-slate-600 hover:text-slate-900 inline-flex items-center gap-1.5 whitespace-nowrap shrink-0"
                      >
                        <ExternalLink className="w-3.5 h-3.5" />
                        <span>تبويب مستقل</span>
                      </a>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </section>

        {/* Professional Presentation & Document Studio Card */}
        <section
          aria-labelledby="studio-hub-heading"
          className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm"
        >
          <div
            className="px-6 py-5 text-white flex flex-wrap items-center justify-between gap-4"
            style={{ background: 'linear-gradient(135deg, #d24726, #b91c1c)' }}
          >
            <div className="flex items-center gap-3.5">
              <span className="w-11 h-11 rounded-xl bg-white/15 flex items-center justify-center shrink-0">
                <Layers className="w-6 h-6 text-[#fcd34d]" />
              </span>
              <div>
                <h2 id="studio-hub-heading" className="text-lg font-extrabold text-white">
                  منشئ العروض والمستندات الاحترافي
                </h2>
                <p className="text-xs font-medium text-red-100 mt-0.5">
                  PPTX · Word · Excel · قوالب كانفا والمعاينة الحية · MP4
                </p>
              </div>
            </div>

            <span className="text-xs font-extrabold bg-amber-300 text-slate-900 px-3 py-1.5 rounded-lg">
              AI · Canva
            </span>
          </div>

          <div className="p-6 bg-slate-50/70 space-y-4">
            <button
              type="button"
              id="btnQuickPresentationCreator"
              onClick={() => openTool(PRESENTATION_STUDIO_TOOL, '/presentation_studio?tab=text')}
              title="إنشاء عروض .pptx ومستندات .docx · تحويل PDF و HTML إلى Word و Excel و PowerPoint وفيديو MP4 · قوالب كانفا وحسابات Google · تحليل وتحديث .pptx + مراقبة السرعة عبر Recharts"
              className="w-full py-3.5 px-5 rounded-xl text-white font-extrabold text-sm shadow-sm flex items-center justify-center gap-2.5 transition-opacity hover:opacity-95 cursor-pointer"
              style={{ background: 'linear-gradient(135deg, #d24726, #991b1b)' }}
            >
              <Sparkles className="w-5 h-5" />
              <span>فتح استوديو العروض والمستندات الاحترافي</span>
            </button>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
              <button
                type="button"
                onClick={() => openTool(PRESENTATION_STUDIO_TOOL, '/presentation_studio?tab=canva')}
                className="bg-white border border-red-200 hover:border-red-400 text-red-700 rounded-xl px-4 py-3 text-xs font-extrabold flex items-center justify-center gap-2 transition-colors cursor-pointer"
              >
                <Palette className="w-4 h-4 shrink-0" />
                <span>قوالب كانفا والمعاينة الحية</span>
              </button>
              <button
                type="button"
                onClick={() => openTool(PRESENTATION_STUDIO_TOOL, '/presentation_studio?tab=analyze')}
                className="bg-white border border-blue-200 hover:border-blue-400 text-blue-700 rounded-xl px-4 py-3 text-xs font-extrabold flex items-center justify-center gap-2 transition-colors cursor-pointer"
              >
                <Presentation className="w-4 h-4 shrink-0" />
                <span>تحليل وتحديث PPTX</span>
              </button>
              <button
                type="button"
                onClick={() => openTool(PRESENTATION_STUDIO_TOOL, '/presentation_studio?tab=docs')}
                className="bg-white border border-emerald-200 hover:border-emerald-400 text-emerald-700 rounded-xl px-4 py-3 text-xs font-extrabold flex items-center justify-center gap-2 transition-colors cursor-pointer"
              >
                <FileSpreadsheet className="w-4 h-4 shrink-0" />
                <span>محول Word و Excel</span>
              </button>
              <button
                type="button"
                onClick={() => openTool(PRESENTATION_STUDIO_TOOL, '/presentation_studio?tab=video')}
                className="bg-white border border-slate-300 hover:border-slate-500 text-slate-800 rounded-xl px-4 py-3 text-xs font-extrabold flex items-center justify-center gap-2 transition-colors cursor-pointer"
              >
                <Film className="w-4 h-4 shrink-0" />
                <span>محول HTML إلى فيديو MP4</span>
              </button>
            </div>
          </div>
        </section>
      </main>

      {/* ================= Footer ================= */}
      <footer className="border-t border-slate-200 bg-white">
        <div className="max-w-[1200px] mx-auto px-6 py-4 flex flex-wrap items-center justify-between gap-4 text-xs text-slate-500 font-semibold">
          <span>بوابة الخدمات التعليمية والأنشطة المدرسية — ملف الأداء، الشواهد، التقارير، المحرر، وبطاقات التنفيذ</span>
          <span>1448هـ</span>
        </div>
      </footer>

      {/* ================= "فتح الكل" Educational Services Modal ================= */}
      {allServicesOpen && (
        <div
          className="fixed inset-0 z-40 bg-slate-900/55 backdrop-blur-xs flex items-center justify-center p-4"
          role="dialog"
          aria-modal="true"
          aria-labelledby="edu-modal-title"
          onClick={() => setAllServicesOpen(false)}
        >
          <div
            className="bg-white rounded-2xl max-w-4xl w-full overflow-hidden shadow-xl border border-slate-200 max-h-[90vh] flex flex-col"
            onClick={(e) => e.stopPropagation()}
          >
            <div
              className="px-6 py-4 text-white flex items-center justify-between gap-4 shrink-0"
              style={{ background: 'linear-gradient(135deg, #0f766e, #0d9488)' }}
            >
              <div className="flex items-center gap-3">
                <GraduationCap className="w-6 h-6 text-[#fcd34d]" />
                <div>
                  <h2 id="edu-modal-title" className="text-lg font-extrabold">
                    وظائف الخدمات التعليمية
                  </h2>
                  <p className="text-xs text-teal-100">
                    أدوات وملفات تعليمية جاهزة للتحرير والطباعة والتصدير
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setAllServicesOpen(false)}
                aria-label="إغلاق"
                className="w-9 h-9 rounded-lg bg-white/15 hover:bg-white/25 flex items-center justify-center text-white cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-6 bg-slate-50 overflow-y-auto grid grid-cols-1 md:grid-cols-2 gap-4">
              {EDU_TOOLS.filter((t) => t.id !== 'activities_portfolio').map((tool) => {
                const IconComponent = tool.icon;
                return (
                  <button
                    key={tool.id}
                    type="button"
                    onClick={() => openTool(tool)}
                    className="bg-white p-5 rounded-xl border text-center flex flex-col items-center justify-center gap-3 hover:shadow-md transition-transform duration-150 hover:-translate-y-0.5 cursor-pointer"
                    style={{ borderColor: tool.accentHex }}
                  >
                    <span
                      className="w-14 h-14 rounded-full flex items-center justify-center"
                      style={{
                        backgroundColor: `${tool.accentHex}18`,
                        color: tool.accentHex,
                      }}
                    >
                      <IconComponent className="w-7 h-7" />
                    </span>
                    <span className="text-base font-extrabold text-slate-900">
                      {tool.title}
                    </span>
                    <span className="text-xs text-slate-600 leading-relaxed max-w-md">
                      {tool.description}
                    </span>
                    {tool.id === 'school_reports' && mergedTemplateName && (
                      <span className="text-xs font-bold text-[#0f766e]">
                        قالب مدمج: {mergedTemplateName}
                      </span>
                    )}
                  </button>
                );
              })}

              <button
                type="button"
                id="btnOpenSchoolActivitiesPlan"
                onClick={() => openTool(EDU_TOOLS[4], '/activities_portfolio?view=plan')}
                className="bg-white p-5 rounded-xl border text-center flex flex-col items-center justify-center gap-3 hover:shadow-md transition-transform duration-150 hover:-translate-y-0.5 cursor-pointer"
                style={{ borderColor: '#0f4c5c' }}
              >
                <span
                  className="w-14 h-14 rounded-full flex items-center justify-center"
                  style={{ backgroundColor: 'rgba(15,76,92,0.12)', color: '#0f4c5c' }}
                >
                  <FolderKanban className="w-7 h-7" />
                </span>
                <span className="text-base font-extrabold text-slate-900">
                  ملف الأنشطة المدرسية
                </span>
                <span className="text-xs text-slate-600 leading-relaxed max-w-md">
                  خطة النشاط الطلابي والتكامل المنهجي لمادة الدراسات الاجتماعية (الأسابيع 1–19)
                </span>
              </button>

              <button
                type="button"
                id="btnOpenSchoolActivitiesCards"
                onClick={() => openTool(EDU_TOOLS[4], '/activities_portfolio?view=cards')}
                className="bg-white p-5 rounded-xl border text-center flex flex-col items-center justify-center gap-3 hover:shadow-md transition-transform duration-150 hover:-translate-y-0.5 cursor-pointer"
                style={{ borderColor: '#c89b3c' }}
              >
                <span
                  className="w-14 h-14 rounded-full flex items-center justify-center"
                  style={{ backgroundColor: 'rgba(200,155,60,0.16)', color: '#8a6617' }}
                >
                  <CalendarCheck className="w-7 h-7" />
                </span>
                <span className="text-base font-extrabold text-slate-900">
                  بطاقة تنفيذ نشاط
                </span>
                <span className="text-xs text-slate-600 leading-relaxed max-w-md">
                  بطاقات تنفيذ برامج النشاط الطلابي الرسمية المعبأة وجاهزة للطباعة
                </span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ================= Fullscreen Interactive Workspace Modal ================= */}
      {activeTool && (
        <div
          className="fixed inset-0 z-50 bg-white flex flex-col"
          role="dialog"
          aria-modal="true"
          aria-label={activeTool.title}
        >
          <div
            className="text-white shrink-0 shadow-sm"
            style={{ background: activeTool.headerGradient }}
          >
            {/* Row 1: Identity, Status & Window Controls */}
            <div className="px-4 py-2 flex flex-wrap items-center justify-between gap-3">
              <div className="flex flex-wrap items-center gap-3 min-w-0">
                <activeTool.icon className="w-5 h-5 text-[#fcd34d] shrink-0" />
                <h2 className="text-sm sm:text-base font-extrabold truncate">
                  {activeTool.id === 'activities_portfolio'
                    ? activitiesView === 'cards'
                      ? 'بطاقة تنفيذ نشاط'
                      : 'ملف الأنشطة المدرسية'
                    : activeTool.title}
                </h2>
                {activeTool.supportsPrintActions && (
                  <span className="text-xs font-semibold text-white/85 whitespace-nowrap">
                    — {toolStatus}
                  </span>
                )}
              </div>

              <div className="flex items-center gap-2 shrink-0">
                {activeTool.id === 'presentation_studio' && (
                  <div className="hidden xl:flex items-center gap-1.5 ml-2">
                    {[
                      { label: 'إنشاء وتصدير', path: '/presentation_studio?tab=text' },
                      { label: 'قوالب كانفا', path: '/presentation_studio?tab=canva' },
                      { label: 'تحليل وتحديث PPTX', path: '/presentation_studio?tab=analyze' },
                      { label: 'تحويل PDF و HTML', path: '/presentation_studio?tab=docs' },
                      { label: 'HTML إلى فيديو', path: '/presentation_studio?tab=video' },
                    ].map((tabItem) => (
                      <button
                        key={tabItem.path}
                        type="button"
                        onClick={() => setActiveToolPath(tabItem.path)}
                        className={`px-2.5 py-1 rounded-md text-xs font-bold transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                          activeToolPath === tabItem.path
                            ? 'bg-white text-slate-900'
                            : 'bg-white/15 text-white hover:bg-white/25'
                        }`}
                      >
                        {tabItem.label}
                      </button>
                    ))}
                  </div>
                )}

                <a
                  href={standaloneHref(activeToolPath || activeTool.path)}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="px-2.5 py-1.5 rounded-md bg-white text-slate-900 hover:bg-slate-100 text-xs font-extrabold inline-flex items-center gap-1.5 whitespace-nowrap shrink-0"
                >
                  <ExternalLink className="w-3.5 h-3.5" />
                  <span>تبويب مستقل</span>
                </a>

                <button
                  type="button"
                  onClick={() => setActiveTool(null)}
                  aria-label="إغلاق"
                  className="w-8 h-8 rounded-md bg-white/15 hover:bg-white/25 flex items-center justify-center text-white cursor-pointer shrink-0"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
            </div>

            {/* Presentation Studio sub-tabs on smaller screens */}
            {activeTool.id === 'presentation_studio' && (
              <div className="xl:hidden px-4 py-2 border-t border-white/15 flex flex-wrap items-center gap-1.5">
                {[
                  { label: 'إنشاء وتصدير', path: '/presentation_studio?tab=text' },
                  { label: 'قوالب كانفا', path: '/presentation_studio?tab=canva' },
                  { label: 'تحليل وتحديث PPTX', path: '/presentation_studio?tab=analyze' },
                  { label: 'تحويل PDF و HTML', path: '/presentation_studio?tab=docs' },
                  { label: 'HTML إلى فيديو', path: '/presentation_studio?tab=video' },
                ].map((tabItem) => (
                  <button
                    key={tabItem.path}
                    type="button"
                    onClick={() => setActiveToolPath(tabItem.path)}
                    className={`px-2.5 py-1 rounded-md text-xs font-bold transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                      activeToolPath === tabItem.path
                        ? 'bg-white text-slate-900'
                        : 'bg-white/15 text-white hover:bg-white/25'
                    }`}
                  >
                    {tabItem.label}
                  </button>
                ))}
              </div>
            )}

            {/* Row 2: Unified Single Action Bar for Educational Documents (No Duplicates!) */}
            {activeTool.supportsPrintActions && (
              <div className="px-4 py-2 border-t border-white/15 bg-black/15 flex flex-wrap items-center justify-between gap-2">
                <div className="flex flex-wrap items-center gap-1.5">
                  {activeTool.id === 'activities_portfolio' && (
                    <div className="inline-flex items-center rounded-md bg-white/12 p-0.5 gap-0.5 ml-1">
                      <button
                        type="button"
                        onClick={() => switchActivitiesSubview('plan')}
                        className={`px-2.5 py-1 rounded text-xs font-bold transition-colors whitespace-nowrap cursor-pointer ${
                          activitiesView === 'plan'
                            ? 'bg-white text-[#09323d]'
                            : 'text-white hover:bg-white/15'
                        }`}
                      >
                        ملف الأنشطة المدرسية
                      </button>
                      <button
                        type="button"
                        onClick={() => switchActivitiesSubview('cards')}
                        className={`px-2.5 py-1 rounded text-xs font-bold transition-colors whitespace-nowrap cursor-pointer ${
                          activitiesView === 'cards'
                            ? 'bg-white text-[#09323d]'
                            : 'text-white hover:bg-white/15'
                        }`}
                      >
                        بطاقة تنفيذ نشاط
                      </button>
                    </div>
                  )}

                  {(activeTool.id === 'school_reports' ||
                    activeTool.id === 'activities_portfolio') && (
                    <div className="inline-flex items-center rounded-md bg-white/12 p-0.5 gap-0.5 ml-1">
                      <button
                        type="button"
                        onClick={() => switchToolModel('primary')}
                        className={`px-2.5 py-1 rounded text-xs font-bold transition-colors whitespace-nowrap cursor-pointer ${
                          toolModel === 'primary'
                            ? 'bg-[#fcd34d] text-slate-900'
                            : 'text-white hover:bg-white/15'
                        }`}
                      >
                        النموذج الأساسي
                      </button>
                      <button
                        type="button"
                        onClick={() => switchToolModel('secondary')}
                        className={`px-2.5 py-1 rounded text-xs font-bold transition-colors whitespace-nowrap cursor-pointer ${
                          toolModel === 'secondary'
                            ? 'bg-[#fcd34d] text-slate-900'
                            : 'text-white hover:bg-white/15'
                        }`}
                      >
                        النموذج الثانوي
                      </button>
                    </div>
                  )}

                  <button
                    type="button"
                    onClick={() => callFrameAction('saveData')}
                    className="px-2.5 py-1.5 rounded-md bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold inline-flex items-center gap-1.5 whitespace-nowrap cursor-pointer"
                  >
                    <Save className="w-3.5 h-3.5" />
                    <span>حفظ</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => callFrameAction('loadData')}
                    className="px-2.5 py-1.5 rounded-md bg-white/15 hover:bg-white/25 text-white text-xs font-bold inline-flex items-center gap-1.5 whitespace-nowrap cursor-pointer"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    <span>استرجاع</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => callFrameAction('addNewPage')}
                    className="px-2.5 py-1.5 rounded-md bg-white/15 hover:bg-white/25 text-white text-xs font-bold inline-flex items-center gap-1.5 whitespace-nowrap cursor-pointer"
                  >
                    <FilePlus2 className="w-3.5 h-3.5" />
                    <span>
                      {activeTool.id === 'activities_portfolio' && activitiesView === 'cards'
                        ? 'بطاقة جديدة'
                        : 'صفحة جديدة'}
                    </span>
                  </button>

                  {activeTool.id === 'performance_portfolio' && (
                    <button
                      type="button"
                      onClick={() => callFrameAction('openModelUploadModal')}
                      className="px-2.5 py-1.5 rounded-md bg-teal-600 hover:bg-teal-500 text-white text-xs font-bold inline-flex items-center gap-1.5 whitespace-nowrap cursor-pointer"
                    >
                      <Upload className="w-3.5 h-3.5" />
                      <span>إضافة شاهد / رفع نموذج</span>
                    </button>
                  )}

                  {activeTool.id === 'school_reports' && (
                    <button
                      type="button"
                      onClick={() => callFrameAction('addNewSection')}
                      className="px-2.5 py-1.5 rounded-md bg-white/15 hover:bg-white/25 text-white text-xs font-bold inline-flex items-center gap-1.5 whitespace-nowrap cursor-pointer"
                    >
                      <FolderPlus className="w-3.5 h-3.5" />
                      <span>إضافة تقرير</span>
                    </button>
                  )}

                  <button
                    type="button"
                    onClick={() =>
                      callFrameAction(
                        activeTool.id === 'school_reports'
                          ? 'openEduCanvaStudio'
                          : 'openCanvaHubModal'
                      )
                    }
                    className="px-2.5 py-1.5 rounded-md text-white text-xs font-bold inline-flex items-center gap-1.5 whitespace-nowrap cursor-pointer"
                    style={{ background: 'linear-gradient(135deg, #00c4cc, #7c3aed)' }}
                  >
                    <Palette className="w-3.5 h-3.5" />
                    <span>قوالب كانفا</span>
                  </button>
                </div>

                <div className="flex flex-wrap items-center gap-1.5">
                  <button
                    type="button"
                    onClick={() => callFrameAction('togglePrintPreview')}
                    className={`px-2.5 py-1.5 rounded-md text-xs font-extrabold inline-flex items-center gap-1.5 whitespace-nowrap cursor-pointer transition-colors ${
                      isPreviewActive
                        ? 'bg-white text-slate-900 ring-2 ring-amber-300'
                        : 'bg-amber-400 text-slate-900 hover:bg-amber-300'
                    }`}
                  >
                    <Eye className="w-3.5 h-3.5" />
                    <span>{isPreviewActive ? 'إنهاء المعاينة' : 'معاينة قبل الطباعة'}</span>
                  </button>

                  <button
                    type="button"
                    onClick={() =>
                      callFrameAction(
                        activeTool.id === 'activities_portfolio'
                          ? 'exportActivitiesPDF'
                          : 'exportPDF'
                      )
                    }
                    className="px-2.5 py-1.5 rounded-md bg-emerald-500 hover:bg-emerald-400 text-white text-xs font-extrabold inline-flex items-center gap-1.5 whitespace-nowrap cursor-pointer"
                  >
                    <FileDown className="w-3.5 h-3.5" />
                    <span>تصدير PDF</span>
                  </button>

                  <button
                    type="button"
                    onClick={() =>
                      callFrameAction(
                        activeTool.id === 'activities_portfolio'
                          ? 'exportActivitiesHTML'
                          : 'exportHTML'
                      )
                    }
                    className="px-2.5 py-1.5 rounded-md bg-sky-600 hover:bg-sky-500 text-white text-xs font-extrabold inline-flex items-center gap-1.5 whitespace-nowrap cursor-pointer"
                  >
                    <FileCode2 className="w-3.5 h-3.5" />
                    <span>تصدير HTML</span>
                  </button>

                  <button
                    type="button"
                    onClick={() =>
                      callFrameAction(
                        activeTool.id === 'activities_portfolio'
                          ? 'printActivities'
                          : 'printDocument'
                      )
                    }
                    className="px-2.5 py-1.5 rounded-md bg-white/15 hover:bg-white/25 text-white text-xs font-bold inline-flex items-center gap-1.5 whitespace-nowrap cursor-pointer"
                  >
                    <Printer className="w-3.5 h-3.5" />
                    <span>طباعة</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => callFrameAction('resetAll')}
                    className="px-2.5 py-1.5 rounded-md bg-red-600/90 hover:bg-red-500 text-white text-xs font-bold inline-flex items-center gap-1.5 whitespace-nowrap cursor-pointer"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>إعادة تعيين</span>
                  </button>
                </div>
              </div>
            )}
          </div>

          <div className="flex-1 min-h-0 bg-slate-200">
            <iframe
              ref={iframeRef}
              key={
                activeTool.id === 'presentation_studio'
                  ? activeToolPath
                  : activeTool.id
              }
              src={activeToolPath || activeTool.path}
              title={activeTool.title}
              allow="fullscreen"
              className="w-full h-full border-0 block"
            />
          </div>
        </div>
      )}

      <div id="eduDirectPrintHost" aria-hidden="true" />
    </div>
  );
}
