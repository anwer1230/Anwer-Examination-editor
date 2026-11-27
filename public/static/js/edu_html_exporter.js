(function (global) {
  'use strict';

  const DEFAULT_REMOVE_SELECTORS = [
    '.edit-only',
    '.no-print',
    '.toolbar',
    '.preview-bar',
    '.app-topbar',
    '.sub-filter-bar',
    '.fmt-bar',
    '.table-controls',
    '.row-del-btn',
    'td.action-cell',
    '.add-page-bar',
    '.page-ctrl-bar',
    '.witness-actions',
    '.witness-box-tools',
    '.witness-quick-bar',
    '.witness-below-hint',
    '.witness-del-btn',
    '.witness-thumb .remove-thumb',
    '.witness-info .link-count',
    '.card-meta-editbar',
    '.custom-sheet-bar',
    '.canva-embedded-controls',
    '.preview-page-badge',
    '.canva-hub-modal',
    '#eduCanvaStudioInline',
    '#ecsMergedStatusBanner',
    '#ecsInlineDock',
    '.pdf-overlay',
    '.pdf-ready-banner',
    '.modal-overlay',
    '.app-confirm-overlay',
    '.cropper-modal',
    '.img-lightbox',
    '.witness-modal',
    'script'
  ];

  function escHtml(str) {
    return String(str == null ? '' : str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function sanitizeFilename(name, fallback) {
    const raw = String(name || fallback || 'document').trim();
    const cleaned = raw
      .replace(/[\\/:*?"<>|]+/g, '_')
      .replace(/[^\u0600-\u06FFa-zA-Z0-9_-]+/g, '_')
      .replace(/^_+|_+$/g, '')
      .replace(/_+/g, '_');
    return cleaned || 'document';
  }

  function syncFormValuesToDom(sourceRoot, cloneRoot) {
    if (!sourceRoot || !cloneRoot) return;
    const srcInputs = Array.from(sourceRoot.querySelectorAll('input'));
    const dstInputs = Array.from(cloneRoot.querySelectorAll('input'));
    srcInputs.forEach((src, idx) => {
      const dst = dstInputs[idx];
      if (!dst) return;
      const t = (src.type || '').toLowerCase();
      if (t === 'checkbox' || t === 'radio') {
        if (src.checked) dst.setAttribute('checked', 'checked');
        else dst.removeAttribute('checked');
      } else {
        dst.setAttribute('value', src.value != null ? String(src.value) : '');
      }
      dst.setAttribute('readonly', 'readonly');
    });

    const srcAreas = Array.from(sourceRoot.querySelectorAll('textarea'));
    const dstAreas = Array.from(cloneRoot.querySelectorAll('textarea'));
    srcAreas.forEach((src, idx) => {
      const dst = dstAreas[idx];
      if (!dst) return;
      dst.textContent = src.value != null ? String(src.value) : '';
      dst.setAttribute('readonly', 'readonly');
    });

    const srcSelects = Array.from(sourceRoot.querySelectorAll('select'));
    const dstSelects = Array.from(cloneRoot.querySelectorAll('select'));
    srcSelects.forEach((src, idx) => {
      const dst = dstSelects[idx];
      if (!dst) return;
      const srcOpts = Array.from(src.options || []);
      const dstOpts = Array.from(dst.options || []);
      srcOpts.forEach((opt, oi) => {
        if (!dstOpts[oi]) return;
        if (opt.selected) dstOpts[oi].setAttribute('selected', 'selected');
        else dstOpts[oi].removeAttribute('selected');
      });
    });
  }

  function replaceCanvasesWithImages(sourceRoot, cloneRoot) {
    if (!cloneRoot) return;
    const srcCanvases = sourceRoot ? Array.from(sourceRoot.querySelectorAll('canvas')) : [];
    const dstCanvases = Array.from(cloneRoot.querySelectorAll('canvas'));
    dstCanvases.forEach((dstCv, idx) => {
      const liveCv = srcCanvases[idx] || dstCv;
      if (!dstCv.parentNode) return;
      try {
        const dataUrl = liveCv.toDataURL('image/png');
        if (dataUrl && dataUrl.startsWith('data:image')) {
          const img = document.createElement('img');
          img.src = dataUrl;
          img.alt = liveCv.getAttribute('aria-label') || 'QR Code';
          const w = liveCv.style.width || (liveCv.width ? liveCv.width + 'px' : '68px');
          const h = liveCv.style.height || (liveCv.height ? liveCv.height + 'px' : '68px');
          img.style.cssText = liveCv.style.cssText || `width:${w};height:${h};display:block;`;
          dstCv.parentNode.replaceChild(img, dstCv);
        }
      } catch (_) {}
    });
  }

  async function fetchAsDataUrl(rawUrl, cacheMap) {
    if (!rawUrl || typeof rawUrl !== 'string') return rawUrl;
    const trimmed = rawUrl.trim();
    if (!trimmed || trimmed.startsWith('data:') || trimmed.startsWith('#')) return trimmed;

    let absUrl = trimmed;
    try {
      absUrl = new URL(trimmed, global.location.href).href;
    } catch (_) {
      return trimmed;
    }

    if (cacheMap.has(absUrl)) return cacheMap.get(absUrl);

    const isSameOrigin = absUrl.startsWith(global.location.origin) || absUrl.startsWith('blob:');
    const promise = (async () => {
      if (!isSameOrigin) {
        return absUrl;
      }
      try {
        const controller = typeof AbortController === 'function' ? new AbortController() : null;
        const timer = controller ? setTimeout(() => controller.abort(), 3500) : null;
        const res = await fetch(absUrl, controller ? { signal: controller.signal } : undefined);
        if (timer) clearTimeout(timer);
        if (!res.ok) return absUrl;
        const blob = await res.blob();
        if (!blob || blob.size > 4 * 1024 * 1024) return absUrl;
        return await new Promise((resolve) => {
          const reader = new FileReader();
          reader.onloadend = () => resolve(typeof reader.result === 'string' ? reader.result : absUrl);
          reader.onerror = () => resolve(absUrl);
          reader.readAsDataURL(blob);
        });
      } catch (_) {
        return absUrl;
      }
    })();

    cacheMap.set(absUrl, promise);
    return promise;
  }

  async function inlineCloneAssets(cloneRoot) {
    if (!cloneRoot) return;
    const cacheMap = new Map();
    const tasks = [];

    cloneRoot.querySelectorAll('img[src]').forEach((img) => {
      const rawSrc = img.getAttribute('src') || '';
      if (!rawSrc || rawSrc.startsWith('data:')) return;
      tasks.push(
        fetchAsDataUrl(rawSrc, cacheMap).then((resolved) => {
          if (resolved) img.setAttribute('src', resolved);
        })
      );
    });

    cloneRoot.querySelectorAll('[style]').forEach((el) => {
      const styleVal = el.getAttribute('style') || '';
      if (!styleVal || !styleVal.includes('url(')) return;
      const matches = Array.from(styleVal.matchAll(/url\(\s*(['"]?)([^'")]+)\1\s*\)/gi));
      if (!matches.length) return;
      const p = Promise.all(
        matches.map(async (m) => ({
          full: m[0],
          url: m[2],
          resolved: await fetchAsDataUrl(m[2], cacheMap),
        }))
      ).then((replacements) => {
        let nextStyle = styleVal;
        replacements.forEach((r) => {
          if (r.resolved && r.resolved !== r.url) {
            nextStyle = nextStyle.split(r.full).join(`url("${r.resolved}")`);
          }
        });
        el.setAttribute('style', nextStyle);
      });
      tasks.push(p);
    });

    if (tasks.length) {
      await Promise.allSettled(tasks);
    }
  }

  function collectHeadStylesHtml() {
    const nodes = Array.from(document.querySelectorAll('link[rel="stylesheet"], style'));
    const parts = [];
    nodes.forEach((node) => {
      if (node.id === 'eduHtmlExporterTempCss') return;
      if (node.tagName === 'LINK') {
        const href = node.getAttribute('href') || '';
        if (!href) return;
        let absHref = href;
        try {
          if (!/^https?:\/\//i.test(href)) {
            absHref = new URL(href, global.location.href).href;
          }
        } catch (_) {}
        parts.push(`<link rel="stylesheet" href="${escHtml(absHref)}">`);
      } else if (node.tagName === 'STYLE') {
        const cssText = node.textContent || '';
        if (!cssText.trim()) return;
        const idAttr = node.id ? ` id="${escHtml(node.id)}"` : '';
        parts.push(`<style${idAttr}>${cssText}</style>`);
      }
    });
    return parts.join('\n');
  }

  function stripInteractiveChrome(cloneRoot, extraSelectors) {
    if (!cloneRoot) return;
    const selectors = DEFAULT_REMOVE_SELECTORS.concat(Array.isArray(extraSelectors) ? extraSelectors : []);
    selectors.forEach((sel) => {
      try {
        cloneRoot.querySelectorAll(sel).forEach((el) => el.remove());
      } catch (_) {}
    });
    cloneRoot.querySelectorAll('[contenteditable]').forEach((el) => {
      el.removeAttribute('contenteditable');
    });
    cloneRoot.querySelectorAll('[data-h2c-target]').forEach((el) => {
      el.removeAttribute('data-h2c-target');
    });
  }

  function cleanBodyClasses(rawClasses) {
    const drop = new Set([
      'print-preview-mode',
      'pdf-exporting',
      'printing-canva-preview',
      'embedded-mode',
      'edu-direct-printing'
    ]);
    return String(rawClasses || '')
      .split(/\s+/)
      .filter((cls) => cls && !drop.has(cls))
      .join(' ');
  }

  function triggerHtmlDownload(fullHtml, filenameBase) {
    const safeName = sanitizeFilename(filenameBase, 'exported_document');
    const filename = safeName.toLowerCase().endsWith('.html') ? safeName : `${safeName}.html`;
    const blob = new Blob(['\uFEFF' + fullHtml], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      try { a.remove(); } catch (_) {}
      try { URL.revokeObjectURL(url); } catch (_) {}
    }, 10000);
    return filename;
  }

  async function exportHtmlDocument(options) {
    const opts = options || {};
    if (document.activeElement && typeof document.activeElement.blur === 'function') {
      try { document.activeElement.blur(); } catch (_) {}
    }

    const baseTitle = opts.title || document.title || 'مستند_تعليمي';
    const canvaName = opts.canvaName ? String(opts.canvaName).trim() : '';
    const filenameBase = opts.filename || (canvaName ? `${baseTitle}_${canvaName}` : baseTitle);

    let cloneRoot = null;
    let sourceRoot = null;

    if (typeof opts.buildContainer === 'function') {
      const built = await opts.buildContainer();
      if (built && built.nodeType === 1) {
        cloneRoot = built;
      } else if (typeof built === 'string') {
        const tmp = document.createElement('div');
        tmp.innerHTML = built;
        cloneRoot = tmp;
      }
    } else if (opts.targetEl) {
      sourceRoot = typeof opts.targetEl === 'string' ? document.querySelector(opts.targetEl) : opts.targetEl;
      if (sourceRoot) {
        cloneRoot = sourceRoot.cloneNode(true);
        syncFormValuesToDom(sourceRoot, cloneRoot);
        replaceCanvasesWithImages(sourceRoot, cloneRoot);
      }
    } else if (typeof opts.targetHtml === 'string') {
      const tmp = document.createElement('div');
      tmp.innerHTML = opts.targetHtml;
      cloneRoot = tmp;
    }

    if (!cloneRoot) {
      throw new Error('لا يوجد محتوى للتصدير كملف HTML');
    }

    stripInteractiveChrome(cloneRoot, opts.removeSelectors);

    if (typeof opts.onAfterClone === 'function') {
      await opts.onAfterClone(cloneRoot, sourceRoot);
    }

    await inlineCloneAssets(cloneRoot);

    const bodyContentHtml = opts.useInnerHtml ? cloneRoot.innerHTML : cloneRoot.outerHTML;
    const fullBodyHtml = `${opts.wrapperOpen || ''}${bodyContentHtml}${opts.wrapperClose || ''}`;
    const headStylesHtml = collectHeadStylesHtml();
    const rootCssText = opts.rootCssText != null ? String(opts.rootCssText) : (document.documentElement.style.cssText || '');
    const bodyClass = opts.bodyClass != null ? cleanBodyClasses(opts.bodyClass) : cleanBodyClasses(document.body.className);
    const bodyStyleAttr = opts.bodyStyle ? ` style="${escHtml(opts.bodyStyle)}"` : '';
    const htmlStyleAttr = rootCssText.trim() ? ` style="${escHtml(rootCssText.trim())}"` : '';

    const exportCleanupCss = `
      * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
      .edit-only, .no-print, .toolbar, .preview-bar, .app-topbar, .sub-filter-bar,
      .fmt-bar, .table-controls, .row-del-btn, td.action-cell, .add-page-bar,
      .page-ctrl-bar, .witness-actions, .witness-box-tools, .witness-quick-bar,
      .witness-below-hint, .witness-del-btn, .card-meta-editbar, .custom-sheet-bar,
      .canva-embedded-controls, .preview-page-badge, .canva-hub-modal,
      #eduCanvaStudioInline, #ecsMergedStatusBanner, #ecsInlineDock,
      .pdf-overlay, .pdf-ready-banner { display: none !important; }
      [contenteditable] { outline: none !important; }
      ${opts.extraCss || ''}
    `;

    const fullDocumentHtml = `<!DOCTYPE html>
<html lang="ar" dir="rtl"${htmlStyleAttr}>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="generator" content="Edu Services Standalone HTML Export">
<title>${escHtml(baseTitle)}${canvaName ? ' - ' + escHtml(canvaName) : ''}</title>
${headStylesHtml}
<style id="eduExportedOverrideCss">
${exportCleanupCss}
</style>
</head>
<body${bodyClass ? ` class="${escHtml(bodyClass)}"` : ''}${bodyStyleAttr}>
${fullBodyHtml}
</body>
</html>`;

    const savedFile = triggerHtmlDownload(fullDocumentHtml, filenameBase);
    if (typeof opts.onStatus === 'function') {
      opts.onStatus(`🌐 تم تصدير وتحميل ملف HTML بنجاح: ${savedFile}`);
    }
    return savedFile;
  }

  global.EduHtmlExporter = {
    exportHtmlDocument,
    syncFormValuesToDom,
    replaceCanvasesWithImages,
    inlineCloneAssets,
    collectHeadStylesHtml,
    stripInteractiveChrome,
    sanitizeFilename,
    triggerHtmlDownload,
  };
})(window);
