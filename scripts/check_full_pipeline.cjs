"use strict";
const fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict'),{spawnSync}=require('node:child_process');
const {chromium}=require(process.argv[2]),root=path.resolve(__dirname,'..'),url=process.argv[4],resume=process.argv[5];
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.argv[3]});const errors=[];
 try{
  const page=await browser.newPage({viewport:{width:1440,height:1100}});page.on('pageerror',e=>errors.push(String(e)));
  await page.goto(url);await page.waitForFunction(()=>document.querySelector('#runner-connection')?.textContent.includes('已登录'));
  let task;
  if(resume){
   await page.locator('#task-import').setInputFiles(path.join(root,'_qa/full-service/runs',resume,'selection.json'));
   await page.waitForFunction(()=>PPT_TASK.getTask().documents.length===1);task=await page.evaluate(()=>PPT_TASK.getTask());
   const existing=await page.evaluate(async id=>(await(await fetch('/api/jobs/latest?taskId='+encodeURIComponent(id))).json()).job,task.id);
   if(existing?.status!=='completed'){await page.waitForSelector('#task-resume:not([hidden])');await page.locator('#task-resume').click();}
  }else{
  await page.locator('#task-title').fill('通用PPT制作流程功能验证');await page.locator('#task-pageRange').fill('两页，不要封面和结束页');
  await page.locator('.task-extra summary').click();
  await page.locator('#task-mustKeep').fill('第一页使用三类并列与共同特征归纳的all-text；第二页使用总述与双栏图文要点。两图只验证绑定，不解释成科研证据。');
  await page.locator('#task-researchScope').fill('原稿和两幅图均为功能测试材料，内容足够，无需联网补充，不引入外部研究结论。');
  for(const [id,variant] of [['BW-TRIPLE-COMMON-01','all-text'],['BW-TWO-PERSPECTIVE-01','base']]){await page.locator(`[data-prefer-id="${id}"]`).check();await page.locator(`[data-variant-for="${id}"]`).selectOption(variant);}
  await page.locator('#task-word-upload').setInputFiles(path.join(root,'_qa/full-pipeline/full-pipeline-test.docx'));
  await page.waitForFunction(()=>PPT_TASK.getTask().documents.length===1);task=await page.evaluate(()=>PPT_TASK.getTask());
  await page.locator('#task-start-analysis').click();console.log('已从界面开始完整自动制作，等待模型、PowerPoint和实际预览检查。');
  }
  let job;const began=Date.now();const seen=new Set();
  while(Date.now()-began<2400000){
   try{job=await page.evaluate(async id=>(await(await fetch('/api/jobs/latest?taskId='+encodeURIComponent(id))).json()).job,task.id);}
   catch(error){console.log('进度连接暂时中断，继续等待原任务。');await new Promise(resolve=>setTimeout(resolve,3000));continue;}
   if(job){if(!seen.has(job.status)){seen.add(job.status);console.log('阶段：'+job.status);}if(['completed','needs-review','failed','cancelled'].includes(job.status))break;}
   await new Promise(resolve=>setTimeout(resolve,2000));
  }
  assert.equal(job?.status,'completed',JSON.stringify(job));assert.equal(job.presentationAvailable,true);assert.equal(job.renderPassed,true);
  await page.waitForSelector('#task-download-pptx');await page.locator('#task-presentation-previews img').first().scrollIntoViewIfNeeded();
  await page.locator('#task-presentation-previews img').evaluateAll(imgs=>Promise.all(imgs.map(img=>img.decode())));
  assert.equal(await page.locator('#task-presentation-previews img').count(),2);
  await page.reload();await page.waitForSelector('#task-download-pptx');assert.equal((await page.locator('#task-job-summary').textContent()).includes('2'),true);
  await page.locator('#task-result-links').scrollIntoViewIfNeeded();await page.screenshot({path:path.join(root,'docs/qa/full-pipeline-desktop.png')});
  await page.setViewportSize({width:390,height:844});await page.locator('#task-result-links').scrollIntoViewIfNeeded();
  assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),true);await page.screenshot({path:path.join(root,'docs/qa/full-pipeline-mobile.png')});
  assert.deepEqual(errors,[]);
  const verified=spawnSync('python',['scripts/check_full_pipeline.py','--job',job.id],{cwd:root,encoding:'utf8'});
  if(verified.status!==0)throw new Error(verified.stderr||verified.stdout);
  console.log(verified.stdout);
  fs.writeFileSync(path.join(root,'_qa/full-pipeline/last-job.json'),JSON.stringify(job,null,2));
 }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
