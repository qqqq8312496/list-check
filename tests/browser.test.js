const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const {chromium}=require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const X=require('../vendor/xlsx.full.min.js');X.set_fs(fs);
(async()=>{
 const browser=await chromium.launch({channel:'msedge',headless:true});
 const context=await browser.newContext({viewport:{width:1440,height:1100},acceptDownloads:true});
 const page=await context.newPage(), errors=[], external=[];
 page.on('pageerror',e=>errors.push(e.message));page.on('request',r=>{if(/^https?:/.test(r.url())&&!r.url().startsWith('http://127.0.0.1:8878'))external.push(r.url());});
 await page.goto('http://127.0.0.1:8878/');
 await page.screenshot({path:path.join(__dirname,'desktop-empty.png'),fullPage:true});
 await page.click('#demoButton');await page.click('#runButton');
 await page.waitForSelector('.result-panel');assert.equal(await page.locator('#resultTable .results tbody tr').count(),3);
 assert.match(await page.locator('#resultTable').innerText(),/A 多 1 条/);
 await page.locator('[data-filter="duplicates"]').click();assert.equal(await page.locator('#resultTable .results tbody tr').count(),2);
 await page.locator('[data-detail]').first().click();assert.match(await page.locator('#detailContent').innerText(),/第|原始行号/);await page.click('#closeDetail');
 await page.locator('[data-filter="diff"]').click();await page.screenshot({path:path.join(__dirname,'desktop-results.png'),fullPage:true});
 await page.click('#supplementTab');await page.click('#runButton');assert.equal(await page.locator('#resultTable .results tbody tr').count(),7);assert.match(await page.locator('#resultTable').innerText(),/重名待确认/);
 await page.locator('[data-detail="1"]').click();await page.locator('[data-confirm]').first().click();assert.match(await page.locator('#resultTable').innerText(),/已手动确认/);
 if(!await page.locator('.aux-row').getAttribute('open')) await page.locator('.aux-row summary').click();await page.locator('[data-aux="0"]').selectOption('1');await page.locator('[data-aux="1"]').selectOption('1');await page.click('#runButton');
 const download=page.waitForEvent('download');await page.click('#exportButton');const file=await download;const out=path.join(__dirname,'export.xlsx');await file.saveAs(out);
 const wb=X.readFile(out,{raw:true});const rows=X.utils.sheet_to_json(wb.Sheets['补全后的主表'],{header:1,raw:true});assert.equal(rows.length,8);assert.equal(rows[2][4],'13800000001');assert.equal(rows[1][0],'王芳');assert.equal(rows[1][1],'A01');assert.equal(rows[4][4],'');assert.equal(rows[7][4],'13800000005');
 // Actual legacy binary .xls input and multi-row merged header detection.
 await page.setInputFiles('#file0',path.join(__dirname,'fixtures','master.xls'));await page.waitForFunction(()=>document.querySelector('#source0 .filename')?.textContent==='master.xls');
 await page.setInputFiles('#file1',path.join(__dirname,'fixtures','source.xlsx'));await page.waitForFunction(()=>document.querySelector('#source1 .filename')?.textContent==='source.xlsx');
 await page.locator('[data-settings="0"] summary').click();assert.equal(await page.locator('[data-range="0"][data-field="hs"]').inputValue(),'6');assert.equal(await page.locator('[data-range="0"][data-field="he"]').inputValue(),'7');assert.match(await page.locator('[data-key="0"]').innerText(),/基本资料 \/ 工号/);
 await page.locator('[data-fieldpick="2"]').check();if(!await page.locator('.aux-row').getAttribute('open')) await page.locator('.aux-row summary').click();await page.locator('[data-aux="0"]').selectOption('1');await page.locator('[data-aux="1"]').selectOption('1');await page.click('#runButton');assert.match(await page.locator('#resultTable').innerText(),/13800000002/);
 await page.screenshot({path:path.join(__dirname,'desktop-supplement.png'),fullPage:true});
 await page.click('#compareTab');await page.locator('[data-aux="0"]').selectOption('-1');await page.locator('[data-aux="1"]').selectOption('-1');await page.click('#runButton');assert.equal(await page.locator('#resultTable .results tbody tr').count(),2);
 await page.setInputFiles('#file1',path.join(__dirname,'fixtures','html-table.xls'));await page.waitForFunction(()=>document.querySelector('#source1 .filename')?.textContent==='html-table.xls');assert.match(await page.locator('#source1').innerText(),/陈晨/);
 const unicodeCsv=path.join(__dirname,'fixtures','utf16.csv');fs.writeFileSync(unicodeCsv,Buffer.from('\ufeff姓名,工号\r\n陈晨,0004\r\n王芳,0001','utf16le'));await page.setInputFiles('#file1',unicodeCsv);await page.waitForFunction(()=>document.querySelector('#source1 .filename')?.textContent==='utf16.csv');assert.match(await page.locator('#source1').innerText(),/陈晨/);assert.match(await page.locator('#source1').innerText(),/0004/);
 await page.locator('[data-range="0"][data-field="hs"]').fill('999999');await page.locator('[data-range="0"][data-field="hs"]').press('Tab');assert.equal(await page.locator('[data-range="0"][data-field="hs"]').inputValue(),'6');
 // Offline file:// usage and paste mode, including numeric text and formula-like content.
 const offline=await context.newPage();await offline.route('**/*',route=>/^https?:/.test(route.request().url())?route.abort():route.continue());await offline.goto('file:///'+path.resolve(__dirname,'../index.html').replace(/\\/g,'/'));
 for(const i of [0,1]){await offline.locator(`[data-paste="${i}"]`).click();await offline.fill('#pasteText','姓名\t工号\n张三\t0012\n=1+1\t0013');await offline.click('#pasteConfirm');}
 await offline.click('#runButton');assert.match(await offline.locator('#resultTable').innerText(),/次数一致/);await offline.locator('[data-filter="all"]').click();assert.match(await offline.locator('#resultTable').innerText(),/=1\+1/);
 const mobile=await context.newPage();await mobile.setViewportSize({width:390,height:844});await mobile.goto('http://127.0.0.1:8878/');await mobile.click('#demoButton');await mobile.click('#runButton');
 assert.equal(await mobile.evaluate(()=>document.documentElement.scrollWidth<=window.innerWidth),true);await mobile.screenshot({path:path.join(__dirname,'mobile.png'),fullPage:true});
 assert.deepEqual(errors,[]);assert.deepEqual(external,[]);console.log('PASS: demo comparison, duplicate detail, manual match, auxiliary match, XLSX export, legacy XLS, merged headers, HTML-XLS, offline paste and mobile layout. No page errors or external requests.');
 await browser.close();
})().catch(e=>{console.error(e);process.exit(1);});


