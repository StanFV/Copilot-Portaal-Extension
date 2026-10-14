import { resolve } from 'node:path';
import { defineConfig, type Plugin } from 'vite';

/**
 * voorbeeld.html: het paneel zonder office.js, voor het voorbeeld op de installatiepagina.
 * Office.js stuurt een iframe buiten Office namelijk naar about:blank ("not hosted in plain browser top window").
 */
function voorbeeld(): Plugin {
  return {
    name: 'voorbeeld',
    apply: 'build',
    enforce: 'post',
    generateBundle(_, bundle) {
      const paneel = bundle['taskpane.html'];
      if (paneel?.type !== 'asset') throw new Error('taskpane.html niet gevonden in de bundel');
      const html = String(paneel.source).replace(/\s*<!--[^>]*Office\.js[\s\S]*?-->\s*<script src="https:\/\/appsforoffice[^"]*"><\/script>/, '');
      if (html.includes('appsforoffice')) throw new Error('office.js niet uit voorbeeld.html gehaald');
      this.emitFile({ type: 'asset', fileName: 'voorbeeld.html', source: html });
    },
  };
}

// Twee pagina's: de installatiepagina (index.html) en het paneel in Outlook (taskpane.html),
// plus voorbeeld.html. Relatieve paden, zodat het onder https://stanfv.github.io/Copilot-Portaal-Extension/ werkt.
export default defineConfig({
  base: './',
  plugins: [voorbeeld()],
  build: {
    // GitHub Pages serveert de map docs/ op main (geen Actions nodig).
    outDir: 'docs',
    emptyOutDir: true,
    rollupOptions: {
      input: { index: resolve(__dirname, 'index.html'), taskpane: resolve(__dirname, 'taskpane.html') },
    },
  },
});
