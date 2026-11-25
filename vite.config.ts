import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import fs from 'fs';
import path from 'path';
import {defineConfig} from 'vite';

export default defineConfig(() => {
  return {
    plugins: [
      react(),
      tailwindcss(),
      // Copies src/assets into dist/ so root-relative URLs like
      // "/src/assets/images/..." (e.g. generate_image output) still resolve
      // in production builds. Do not remove.
      {
        name: 'copy-unbundled-src-assets',
        apply: 'build',
        closeBundle(error?: Error): void {
          if (error) {
            return;
          }
          const distDir = path.resolve(__dirname, 'dist');
          const srcAssets = path.resolve(__dirname, 'src/assets');
          const distAssets = path.resolve(distDir, 'src/assets');
          if (fs.existsSync(distDir) && fs.existsSync(srcAssets)) {
            fs.cpSync(srcAssets, distAssets, {recursive: true, force: true});
          }

          const templateRoutes: Array<[string, string]> = [
            ['performance_portfolio', 'edu_templates/performance_portfolio.html'],
            ['performance_witnesses', 'edu_templates/performance_witnesses.html'],
            ['school_reports', 'edu_templates/school_reports.html'],
            ['activities_portfolio', 'edu_templates/activities_portfolio.html'],
            ['school_activities_cards', 'edu_templates/activities_portfolio.html'],
            ['presentation_studio', 'edu_templates/presentation_studio.html'],
          ];
          for (const [routeSlug, relSource] of templateRoutes) {
            const srcFile = path.resolve(distDir, relSource);
            if (fs.existsSync(srcFile)) {
              const targetDir = path.resolve(distDir, routeSlug);
              fs.mkdirSync(targetDir, {recursive: true});
              fs.copyFileSync(srcFile, path.resolve(targetDir, 'index.html'));
            }
          }
        },
      },
    ],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modify—file watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
      // Disable file watching when DISABLE_HMR is true to save CPU during agent edits.
      watch: process.env.DISABLE_HMR === 'true' ? null : {},
    },
  };
});
