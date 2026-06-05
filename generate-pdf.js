const puppeteer = require('puppeteer');
const path = require('path');

(async () => {
    const browser = await puppeteer.launch({
        headless: 'new',
        executablePath: 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
        channel: 'msedge'
    });
    const page = await browser.newPage();
    
    const htmlPath = path.resolve(__dirname, 'PRD.html');
    await page.goto('file:///' + htmlPath.replace(/\\/g, '/'), { waitUntil: 'networkidle0' });
    
    await page.pdf({
        path: path.resolve(__dirname, 'PRD_中央站报警数据分析系统_产品需求文档.pdf'),
        format: 'A4',
        margin: { top: '20mm', right: '18mm', bottom: '20mm', left: '18mm' },
        printBackground: true,
        displayHeaderFooter: false
    });
    
    await browser.close();
    console.log('PDF generated successfully!');
})();
