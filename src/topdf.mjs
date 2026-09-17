import { chromium } from 'playwright';

const JOBS = [
  { src: 'onyx-medical-placement.html', out: 'pdf/Onyx_2026_Medical_Placement_Report.pdf',
    landscape: true,  scale: 0.62 },
  { src: 'trinet-questions.html',       out: 'pdf/Questions_for_TriNet.pdf',
    landscape: false, scale: 0.80 },
  { src: 'adp-questions.html',          out: 'pdf/Questions_for_ADP.pdf',
    landscape: false, scale: 0.80 },
];

const FIX = `
  html{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body{ background:#fff !important; }
  .scroll{ overflow:visible !important; border-radius:0 }
  .wrap{ max-width:none !important; padding:18px 22px 22px !important; gap:26px !important }
  section, .person, .panel, .verdict, .q, .five>div, .closer, table{ break-inside:avoid }
  thead{ display:table-header-group }
  tr{ break-inside:avoid }
  .ansline{ break-inside:avoid }
`;

const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
for (const j of JOBS) {
  const p = await b.newPage({ viewport: { width: 1400, height: 1000 } });
  await p.emulateMedia({ media: 'print', colorScheme: 'light' });
  await p.goto('file://' + process.cwd() + '/' + j.src, { waitUntil: 'networkidle' });
  await p.addStyleTag({ content: FIX });
  await p.waitForTimeout(1200);
  await p.pdf({
    path: j.out, format: 'Letter', landscape: j.landscape, scale: j.scale,
    printBackground: true,
    margin: { top: '0.4in', bottom: '0.45in', left: '0.4in', right: '0.4in' },
    displayHeaderFooter: true,
    headerTemplate: '<div></div>',
    footerTemplate: `<div style="width:100%;font:8px 'Helvetica';color:#888;padding:0 0.45in;display:flex;justify-content:space-between">
      <span>Onyx Accounting Group, LLC &middot; 2026 benefits placement</span>
      <span><span class="pageNumber"></span> of <span class="totalPages"></span></span></div>`,
  });
  await p.close();
  console.log('wrote', j.out);
}
await b.close();
