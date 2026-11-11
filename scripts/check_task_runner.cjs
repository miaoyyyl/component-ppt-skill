/* 真实模型端到端检查。需先用 --data-dir _qa/task-service --no-auto-commit 启动服务。 */
const fs=require("node:fs"),path=require("node:path"),crypto=require("node:crypto"),{spawnSync}=require("node:child_process"),assert=require("node:assert/strict");
const {chromium}=require(process.argv[2]),root=path.resolve(__dirname,".."),url=process.argv[4];
let checks=0;function check(v,m){assert.ok(v,m);checks++;}
(async()=>{
  const fixture=spawnSync("python",["scripts/check_task_service.py","--write-smoke"],{cwd:root,encoding:"utf8"});
  if(fixture.status!==0)throw new Error(fixture.stderr||fixture.stdout);
  const originalHashes=()=>Object.fromEntries(['by_human','materials'].flatMap(dir=>fs.readdirSync(path.join(root,dir)).filter(file=>/\.(docx|pptx)$/i.test(file)).map(file=>[dir+'/'+file,crypto.createHash('sha256').update(fs.readFileSync(path.join(root,dir,file))).digest('hex')])));
  const before=originalHashes();
  const browser=await chromium.launch({headless:true,executablePath:process.argv[3]});const errors=[],failed=[];
  try {
    const page=await browser.newPage({viewport:{width:1440,height:1100}});page.on("pageerror",e=>errors.push(String(e)));page.on("requestfailed",r=>failed.push(r.url()));
    await page.goto(url);await page.waitForFunction(()=>document.querySelector("#runner-connection")?.textContent.includes("已登录"));
    await page.locator("#task-title").fill("自动分析功能验证（生成测试原稿）");await page.locator("#task-pageRange").fill("2至3页，不需要封面和结束页");
    await page.locator('[data-prefer-id="BW-TRIPLE-COMMON-01"]').check();
    await page.locator('[data-variant-for="BW-TRIPLE-COMMON-01"]').selectOption("all-text");
    await page.locator("#task-word-upload").setInputFiles(path.join(root,"_qa/auto-analysis-smoke.docx"));
    await page.waitForFunction(()=>PPT_TASK.getTask().documents.some(d=>d.name==="auto-analysis-smoke.docx"));
    check(!(await page.locator("#task-start-analysis").isDisabled()),"原稿和选型完成后不能开始分析");
    const snapshot=await page.evaluate(()=>PPT_TASK.getTask());
    await page.locator("#task-start-analysis").click();
    await page.waitForFunction(()=>document.querySelector("#task-job-status").textContent.length>0);
    console.log("已从界面提交真实模型任务，等待内容分析、补充查询和逐页分配。");
    const start=Date.now();let job;
    while(Date.now()-start<950000){
      job=await page.evaluate(async taskId=>(await(await fetch('/api/jobs/latest?taskId='+encodeURIComponent(taskId))).json()).job,snapshot.id);
      if(job&&!['queued','extracting','model-running','validating','saving'].includes(job.status))break;
      await new Promise(resolve=>setTimeout(resolve,2000));
    }
    check(job?.status==="completed","真实模型任务未完成："+JSON.stringify(job));
    await page.waitForFunction(()=>!document.querySelector("#task-job-result").hidden);
    const result=await page.evaluate(async id=>await(await fetch('/api/jobs/'+id+'/analysis.json')).json(),job.id);
    check(result.contentCards.length>0&&result.pages.length>0,"没有真实的内容卡片和逐页方案");
    check(result.pages.some(p=>p.componentId==="BW-TRIPLE-COMMON-01"&&p.variantId==="all-text"),"模型没有采用人工优先组件和指定变体");
    check(result.sources.some(s=>/\.(nasa\.gov|usgs\.gov)$/.test(new URL(s.url).hostname)||['nasa.gov','usgs.gov'].includes(new URL(s.url).hostname)),"模型没有返回测试所需的官方来源");
    check(result.researchLog.length>0,"没有自动资料检索记录");
    await page.reload();await page.waitForFunction(()=>!document.querySelector("#task-job-result").hidden);
    check(await page.locator("#task-title").inputValue()===snapshot.brief.title,"刷新未恢复任务");
    check((await page.locator("#task-job-summary").textContent()).includes(String(job.pageCount)),"刷新未恢复模型结果");
    await page.locator("#task-runner").scrollIntoViewIfNeeded();check(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),"自动任务区桌面横向溢出");
    await page.screenshot({path:path.join(root,"docs/qa/automation-desktop.png")});
    await page.setViewportSize({width:390,height:844});await page.locator("#task-runner").scrollIntoViewIfNeeded();
    check(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),"自动任务区手机横向溢出");await page.screenshot({path:path.join(root,"docs/qa/automation-mobile.png")});
    check(!errors.length,"界面错误："+errors.join(";"));check(!failed.length,"资源失败："+failed.join(";"));
    check(JSON.stringify(originalHashes())===JSON.stringify(before),"自动分析改变了用户原稿或原始PPT");
    fs.writeFileSync(path.join(root,"docs/automation-check.json"),JSON.stringify({passed:true,checksPassed:checks,executionMode:job.executionMode,inputKind:"generated-test-docx",jobId:job.id,contentCards:result.contentCards.length,pages:result.pages.length,
      externalSources:result.sources.map(s=>({url:s.url,readScope:s.readScope})),researchQueries:result.researchLog.length,sourceOriginalsUnchanged:true,pageErrors:errors,failedResources:failed},null,2)+"\n");
    console.log(`真实模型端到端检查通过：${checks}项，${result.contentCards.length}条内容、${result.pages.length}页方案、${result.sources.length}项补充来源。`);
  }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
