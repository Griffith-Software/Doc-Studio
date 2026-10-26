// PDF bundle (built to renderer/vendor/pdf.js by `npm run bundle:pdf`).
// pdf.js (Apache-2.0) renders pages; pdf-lib (MIT) + fontkit (MIT) write changes. Exposes window.DDPdf.
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import * as pdfjsWorker from 'pdfjs-dist/legacy/build/pdf.worker.mjs';
import * as PDFLib from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';

// Run pdf.js's parser on the page's own thread (no separate worker file to load).
globalThis.pdfjsWorker = pdfjsWorker;
pdfjs.GlobalWorkerOptions.workerSrc = 'about:blank';

export { pdfjs, PDFLib, fontkit };
