const puppeteer = require('puppeteer-core');

const CHROME_PATH = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';

async function buscarMatriculaSisa(matricula) {
  if (!matricula || !String(matricula).trim()) {
    return { ok: false, error: 'Matrícula requerida' };
  }

  const matClean = String(matricula).replace(/\D/g, '');
  if (!matClean) {
    return { ok: false, error: 'Matrícula inválida (debe contener dígitos)' };
  }

  const browser = await puppeteer.launch({
    executablePath: CHROME_PATH,
    headless: "new",
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--ignore-certificate-errors',
      '--window-size=1280,1000'
    ]
  });

  try {
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 1000 });
    const sleep = ms => new Promise(r => setTimeout(r, ms));

    async function clickCenter(fn) {
      const box = await page.evaluate(fn);
      if (box) {
        await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
        return true;
      }
      return false;
    }

    await page.goto('https://sisa.msal.gov.ar/sisa/#sisa', { waitUntil: 'networkidle2', timeout: 45000 });
    await sleep(2000);

    // Scroll to Buscador card
    await page.evaluate(() => window.scrollTo(0, 500));
    await sleep(800);

    // Click Buscador card
    await clickCenter(() => {
      const all = Array.from(document.querySelectorAll('*'));
      const el = all.find(e => e.textContent && e.textContent.includes('Buscador de Establecimientos y Profesionales') && e.children.length === 0);
      if (el) {
        const r = (el.closest('.util_cursor_hand') || el).getBoundingClientRect();
        return { x: r.x, y: r.y, width: r.width, height: r.height };
      }
      return null;
    });
    await sleep(2000);

    // Click REFEPS tab
    const clickedRefeps = await clickCenter(() => {
      const all = Array.from(document.querySelectorAll('*'));
      const el = all.find(e => e.textContent && e.textContent.includes('REFEPS') && e.textContent.includes('Profesionales') && e.children.length === 0);
      if (el) {
        const r = (el.closest('div') || el).getBoundingClientRect();
        return { x: r.x, y: r.y, width: r.width, height: r.height };
      }
      return null;
    });
    if (!clickedRefeps) {
      await page.mouse.click(450, 310);
    }
    await sleep(1500);

    // Click Matricula radio button
    await clickCenter(() => {
      const labels = Array.from(document.querySelectorAll('label'));
      const mat = labels.find(l => l.textContent && l.textContent.trim() === 'Matrícula');
      if (mat) {
        const r = mat.getBoundingClientRect();
        return { x: r.x, y: r.y, width: r.width, height: r.height };
      }
      return null;
    });
    await sleep(800);

    // Click Matricula input
    await clickCenter(() => {
      const inputs = Array.from(document.querySelectorAll('input[type="text"]'));
      const inp = inputs.find(i => i.placeholder && i.placeholder.toLowerCase().includes('matr') && i.style.display !== 'none');
      if (inp) {
        const r = inp.getBoundingClientRect();
        return { x: r.x, y: r.y, width: r.width, height: r.height };
      }
      return null;
    });
    await sleep(300);

    // Type the matricula
    await page.keyboard.down('Meta');
    await page.keyboard.press('KeyA');
    await page.keyboard.up('Meta');
    await page.keyboard.press('Backspace');
    await page.keyboard.type(matClean, { delay: 35 });
    await sleep(400);

    // Click Buscar button
    await clickCenter(() => {
      const btns = Array.from(document.querySelectorAll('div.boton, button, a.gwt-Anchor'));
      const btn = btns.find(b => b.textContent && b.textContent.includes('Buscar') && b.offsetParent !== null);
      if (btn) {
        const r = btn.getBoundingClientRect();
        return { x: r.x, y: r.y, width: r.width, height: r.height };
      }
      return null;
    });

    await sleep(3500);

    // Extract table rows
    const results = await page.evaluate(() => {
      const seen = new Set();
      const list = [];
      document.querySelectorAll('tr').forEach(tr => {
        const tds = Array.from(tr.querySelectorAll('td')).map(td => td.textContent.trim());
        const dniIdx = tds.findIndex(c => c === 'DNI');
        if (dniIdx !== -1 && tds[dniIdx + 1] && tds[dniIdx + 2]) {
          const item = {
            codigo: tds[dniIdx - 1] || '',
            tipoDoc: tds[dniIdx] || 'DNI',
            nroDoc: tds[dniIdx + 1] || '',
            nombre: tds[dniIdx + 2] || '',
            compartido: tds[dniIdx + 3] || ''
          };
          const key = item.codigo + '_' + item.nroDoc;
          if (!seen.has(key)) {
            seen.add(key);
            list.push(item);
          }
        }
      });
      return list;
    });

    return {
      ok: true,
      matricula: matClean,
      total: results.length,
      profesionales: results
    };
  } finally {
    await browser.close();
  }
}

if (require.main === module) {
  const query = process.argv[2] || '123456';
  buscarMatriculaSisa(query).then(res => {
    console.log(JSON.stringify(res, null, 2));
  }).catch(e => {
    console.error(e);
    process.exit(1);
  });
}

module.exports = { buscarMatriculaSisa };
