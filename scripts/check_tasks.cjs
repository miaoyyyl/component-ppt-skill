/* 制作任务的选型、草稿恢复与导出检查：参数同 check_catalog.cjs。 */
const fs=require("node:fs"), path=require("node:path"), assert=require("node:assert/strict"), {pathToFileURL}=require("node:url"),crypto=require('node:crypto');
const {chromium}=require(process.argv[2]), root=path.resolve(__dirname,".."), state=require("../library/task-state.js");
const catalog=JSON.parse(fs.readFileSync(path.join(root,"library/components.json"),"utf8"));
let checks=0;
function check(value,message){assert.ok(value,message);checks++;}
const ids=catalog.components.map(c=>c.id), original=state.create(catalog);
const progressId="BW-TRIPLE-PROGRESS-01", centerId="BW-BILATERAL-CENTER-01";
const storageKey="ppt-component-production-tasks.v1", title="测试：本次人工优先选型";
let task=state.setPreferred(original,progressId,true,catalog,"all-text");
task=state.setPreferred(task,centerId,true,catalog,"bottom-text");task=state.move(task,centerId,-1,catalog);
check(original.preferredComponents.length===0,"操作修改了原始任务");
check(task.preferredComponents[0].componentId===centerId&&task.preferredComponents[0].priority===1,"优先级没有跟随顺序");
check(task.fallback.componentIds.length===ids.length-2&&!task.fallback.componentIds.includes(progressId),"备用清单没有排除优先组件");
check(JSON.stringify(state.validate(JSON.parse(JSON.stringify(task)),catalog))===JSON.stringify(task),"任务文件不能无损读取");
for(const version of catalog.compatibleTaskVersions||[]){const legacy=structuredClone(task);legacy.libraryVersion=version;const upgraded=state.validate(legacy,catalog);check(upgraded.libraryVersion===catalog.version&&JSON.stringify(upgraded.preferredComponents)===JSON.stringify(task.preferredComponents),'新版页框丢失旧任务选型');}
for(const modify of [t=>t.preferredComponents.push(t.preferredComponents[0]),t=>t.preferredComponents[0].variantId="unknown",t=>t.workflow.pageAllocation="manual",t=>t.libraryVersion="future",t=>t.brief.durationMinutes=-1,t=>t.schema="invalid",t=>t.paletteId="unknown"]){
  const invalid=structuredClone(task);modify(invalid);assert.throws(()=>state.validate(invalid,catalog));checks++;
}
(async()=>{
  const browser=await chromium.launch({headless:true,executablePath:process.argv[3]});const errors=[],failed=[],apiRequests=[];
  const downloadedTemplates = new Set();
  try {
    const page=await browser.newPage({viewport:{width:1440,height:1000}});page.on("pageerror",e=>errors.push(String(e)));page.on("requestfailed",r=>failed.push([r.url(),r.failure()?.errorText]));
    page.on('request',r=>{if(new URL(r.url()).pathname.startsWith('/api/'))apiRequests.push(r.url());});
    await page.goto(process.argv[4]||pathToFileURL(path.join(root,"library/index.html")).href);
    const componentGroup = id => page.locator(`[data-group="${catalog.components.find(c=>c.id===id).groupId}"]`);
    const expandComponentGroup = async id => {
      const group = componentGroup(id);
      if (await group.getAttribute("open")===null) await group.locator(":scope > summary").click();
    };
    const listHint=page.locator("#list-hint");
    check(await listHint.isHidden()&&await page.getByText("点击名称查看版式",{exact:true}).count()===0,"组件展示仍含点击提示");
    await page.locator('.workspace-links a[href="#selection"]').click();
    check(await listHint.isVisible()&&(await listHint.textContent()).includes("悬停查看版式"),"任务侧栏缺少悬停提示");
    const taskText=await page.locator("#production-task").innerText();
    check(!taskText.includes("在这里确定排版偏好")&&!taskText.includes("JSON 包含要求"),"任务页仍含已删说明");
    check(await page.locator(".task-saved,#task-save,#task-load,#task-import").count()===0,"任务页仍含手动保存或读取区块");
    const handoffHeading=page.locator(".task-handoff h3");
    const headingStyle=await handoffHeading.evaluate(el=>{
      const {fontSize,textAlign}=getComputedStyle(el);
      return {fontSize,textAlign};
    });
    check(headingStyle.fontSize==="20px"&&headingStyle.textAlign==="left","导出标题的字号或对齐不正确");
    const headingBox=await handoffHeading.boundingBox(), exportBox=await page.locator("#task-export").boundingBox();
    check(headingBox.x+headingBox.width<=exportBox.x&&Math.abs(headingBox.y+headingBox.height/2-exportBox.y-exportBox.height/2)<1,"导出标题未在按钮左侧垂直居中");
    const firstEntry=page.locator('.component-nav-entry').first();
    await firstEntry.hover();
    await page.locator('#layout-peek').waitFor({state:'visible'});
    check(await page.locator('#layout-peek strong').textContent()===catalog.components[0].name,'悬停预览组件错误');
    check(await page.locator('#nav-count').textContent()==='0','悬停改变了本次选择');
    await page.keyboard.press('Escape');
    await page.locator('#layout-peek').waitFor({state:'hidden'});
    const suiteToc = catalog.components.find(c=>c.groupId==='suite-blue-wave' && c.pageRole==='toc');
    await expandComponentGroup(suiteToc.id);
    const suiteEntry = page.locator("[data-group='suite-blue-wave'] [data-component='" + suiteToc.id + "']");
    await suiteEntry.hover();
    await page.locator('#layout-peek').waitFor({state:'visible'});
    check(await page.locator('#layout-peek strong').textContent()===suiteToc.name,'新套系目录悬停预览错误');
    await suiteEntry.locator('input').check();
    check(await page.locator("#task-preferred-count").textContent()==="1","新套系组件不能单独选择");
    await suiteEntry.locator('input').uncheck();
    check(await page.locator('#task-runner,#task-word-upload,#task-start-analysis,.model-workflow').count()===0,'组件库仍有自动制作入口');
    await page.locator("#task-title").fill(title);await page.locator("#task-audience").fill("课题组");
    await expandComponentGroup(progressId);
    await expandComponentGroup(centerId);
    await page.locator(`[data-prefer-id="${progressId}"]`).check();await page.locator(`[data-prefer-id="${centerId}"]`).check();
    check(await page.locator("#task-preferred-count").textContent()==="2","多选数量错误");
    check(await page.locator("#task-fallback-count").textContent()===String(ids.length-2),"备用数量错误");
    await page.locator(`[data-component-id="${centerId}"] [data-move="-1"]`).click();
    check(await page.locator("#task-preferred-list>li").first().getAttribute("data-component-id")===centerId,"上移后顺序错误");
    await page.locator(`[data-variant-for="${centerId}"]`).selectOption("bottom-text");
    await page.locator(`[data-component="${centerId}"]`).hover();
    await page.waitForFunction(()=>!document.querySelector('#layout-peek').hidden&&document.querySelector('#layout-peek img').src.endsWith('__bottom-text.png'));
    await componentGroup(centerId).locator(":scope > summary").click();
    await page.locator("#layout-peek").waitFor({state:"hidden"});
    check(await page.locator(`[data-prefer-id="${centerId}"]`).isHidden(), "选型侧栏无法收起正文分组");
    await page.locator(`[data-component-id="${centerId}"] .priority-preview`).click();
    await page.locator('#catalog-detail').waitFor({state:'visible'});
    check(await page.locator('#production-task').isHidden(),'组件预览仍被任务表单遮挡');
    check(await listHint.isHidden(),"从任务切回组件展示后仍显示侧栏提示");
    check(await page.locator(`[data-id="${centerId}"]`).isVisible(),"预览按钮未展开所选组件的分组");
    await page.locator('.workspace-links a[href="#selection"]').click();
    check((await page.locator("#slide-image").getAttribute("src")).endsWith("__bottom-text.png"),"任务变体与预览未同步");
    await page.waitForFunction(({storageKey,title,centerId})=>{
      const store=JSON.parse(localStorage.getItem(storageKey));
      const draft=store?.tasks.find(item=>item.id===store.activeTaskId);
      const first=draft?.preferredComponents[0];
      return draft?.brief.title===title&&first?.componentId===centerId&&first.variantId==="bottom-text";
    },{storageKey,title,centerId});
    await page.reload();
    check(await page.locator("#task-title").inputValue()===title,"刷新丢失制作要求");
    check(await page.locator("#task-preferred-list>li").first().getAttribute("data-component-id")===centerId,"刷新丢失优先级");
    check(await page.locator(`[data-variant-for="${centerId}"]`).inputValue()==="bottom-text","刷新丢失变体");
    const oldId=await page.evaluate(()=>PPT_TASK.getTask().id);
    const downloadPromise=page.waitForEvent("download");await page.locator("#task-export").click();const download=await downloadPromise;
    const exported=JSON.parse(fs.readFileSync(await download.path(),"utf8"));
    check(exported.paletteId==="blue"&&exported.handoff.theme.palette.primary==="0046A5","默认配色应为蓝白");
    check(exported.workflow.continueAutomatically===true&&exported.workflow.contentAnalysis==="model-auto","导出没有模型自动流程");
    check(exported.handoff.execution==='agent-after-user-handoff'&&!exported.handoff.materials.filesEmbedded,'交接范围不清楚');
    check(exported.handoff.materials.wordRequired===false&&exported.documents.length===0,'无Word任务导出仍要求原稿');
    check(exported.handoff.researchWorkflow.mode==='agent-led-topic-and-pdf'&&exported.handoff.researchWorkflow.wordRequired===false,'交接缺少PDF优先流程');
    check(fs.existsSync(path.join(root,exported.handoff.researchWorkflow.runbookPath))&&fs.existsSync(path.join(root,exported.handoff.researchWorkflow.entryScript)),'PDF流程入口不存在');
    check(Object.values(exported.handoff.researchWorkflow.skills).length===5,'科研skill分工不完整');
    check(exported.handoff.components.length===catalog.components.length,'备用组件说明缺失');
    for(const entry of exported.handoff.components){
      const selected=exported.preferredComponents.find(c=>c.componentId===entry.id);
      const component=catalog.components.find(c=>c.id===entry.id);
      check(entry.groupId===component.groupId&&entry.pageRole===component.pageRole&&entry.groupName===catalog.groups.find(g=>g.id===component.groupId).name,'交接丢失组件分组或套系角色');
      check(entry.selectionTier===(selected?'preferred':'fallback'),'交接选型级别错误');
      check(entry.regions.length===component.regions.length,'组件区域功能说明缺失');
      if(selected&&selected.variantId!=='auto')check(entry.forms.length===1&&entry.forms[0].id===selected.variantId,'交接放开了人工指定变体');
      for(const form of entry.forms){
        if(entry.pageRole!=='toc')check(form.slots.some(s=>s.role==='page-title')&&!form.slots.some(s=>s.role==='page-number'),'交接缺少页标题或仍带页码');
        if(entry.id==='BW-TWO-HIERARCHY-01')check(form.slots.some(s=>s.key==='numbers'),'内容编号不应被当成页码删除');
        const file=path.resolve(root,form.templatePath);check(file.startsWith(root+path.sep)&&fs.existsSync(file),'交接母版路径无效');
        check(crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex')===form.templateSha256,'交接母版哈希错误');
        check(form.usesXPlaceholder&&form.textCapacities.length>0&&form.slots.length>0,'交接缺少占位容量或槽位说明');
      }
    }
    check(fs.statSync(await download.path()).size<2*1024*1024,'导出的JSON超过交接容量');
    const paletteSelect = page.locator("#palette-select");
    const preview = page.locator("#slide-image");
    const link = page.locator("#template-link");
    const selectedKey = centerId + "__bottom-text";
    await expandComponentGroup(centerId);
    for (const scheme of catalog.palettes.filter(item => item.id !== "blue")) {
      const prefix = "palettes/" + scheme.id + "/";
      await paletteSelect.selectOption(scheme.id);
      check((await page.locator(`[data-component-id="${centerId}"] .priority-preview img`).getAttribute("src"))
        === prefix + "placeholder/previews/" + selectedKey + ".png", "任务卡片配色或变体不匹配");
      await page.locator(`[data-component="${centerId}"]`).hover();
      await page.locator("#layout-peek").waitFor({state:"visible"});
      check((await page.locator("#layout-peek img").getAttribute("src")).startsWith(prefix), "悬停预览配色错误");
      await page.locator(`[data-component-id="${centerId}"] .priority-preview`).click();
      await page.locator("#show-case").click();
      check((await preview.getAttribute("src")) === prefix + "case/previews/" + selectedKey + ".png", "原案例配色或变体不匹配");
      await preview.evaluate(el => el.decode());
      await page.locator("#show-reference").click();
      check((await preview.getAttribute("src")).startsWith("references/"), "原始参考图路径被主题替换");
      await page.locator("#show-normalized").click();
      check((await preview.getAttribute("src")) === prefix + "placeholder/previews/" + selectedKey + ".png", "占位预览配色或变体不匹配");
      await preview.evaluate(el => el.decode());
      const templateDownload = page.waitForEvent("download");
      await link.click();
      const template = await templateDownload;
      check(template.suggestedFilename().endsWith(`--${scheme.id}-X.pptx`), "母版下载文件名未标明配色");
      await page.locator('.workspace-links a[href="#selection"]').click();
      const themedDownload = page.waitForEvent("download");
      await page.locator("#task-export").click();
      const themed = JSON.parse(fs.readFileSync(await (await themedDownload).path(),"utf8"));
      check(themed.paletteId===scheme.id&&themed.handoff.theme.palette.primary===scheme.primary,"JSON配色设置错误");
      check(themed.handoff.library.themePath==="library/"+prefix+"theme.json","JSON主题路径错误");
      check(JSON.stringify(themed.preferredComponents)===JSON.stringify(exported.preferredComponents),"切换配色后丢失优先级或变体");
      const forms = themed.handoff.components.flatMap(component => component.forms);
      check(forms.every(form => form.templatePath.startsWith("library/"+prefix+"placeholder/templates/")), "交接混入其他配色母版");
      const selected = themed.handoff.components.find(component => component.id === centerId).forms[0];
      check(crypto.createHash("sha256").update(fs.readFileSync(await template.path())).digest("hex")===selected.templateSha256,"下载母版和JSON哈希不一致");
      downloadedTemplates.add(template.url());
    }
    const currentPalette = await paletteSelect.inputValue();
    await page.reload();
    check(await paletteSelect.inputValue()===currentPalette,"刷新丢失配色");
    await paletteSelect.selectOption("blue");
    await page.locator("#task-new").click();check(await page.locator("#task-preferred-count").textContent()==="0","新任务没有清空优先组件");
    await page.waitForFunction(key=>JSON.parse(localStorage.getItem(key))?.activeTaskId===PPT_TASK.getTask().id,storageKey);
    const legacy=await page.evaluate(({storageKey,oldId})=>{
      const store=JSON.parse(localStorage.getItem(storageKey));
      return store.tasks.find(item=>item.id===oldId);
    },{storageKey,oldId});
    check(legacy?.brief.title===title&&legacy.preferredComponents[0]?.variantId==="bottom-text","新建任务丢失旧草稿");
    legacy.libraryVersion="0.8.0";
    delete legacy.paletteId;
    legacy.documents=[{sourceId:'upload:'+'a'.repeat(64),name:'旧任务原稿.docx',sha256:'a'.repeat(64)}];
    legacy.handoff={execution:'untrusted',components:[]};
    await page.addInitScript(({storageKey,legacy})=>{
      const store=JSON.parse(localStorage.getItem(storageKey));
      store.activeTaskId=legacy.id;
      store.tasks=store.tasks.map(item=>item.id===legacy.id?legacy:item);
      localStorage.setItem(storageKey,JSON.stringify(store));
    },{storageKey,legacy});
    await page.reload();
    check(await paletteSelect.inputValue()==="blue","旧版草稿未默认使用蓝白");
    check(await page.locator("#task-preferred-count").textContent()==="2","旧版草稿恢复时丢失选择");
    const restored=await page.evaluate(()=>PPT_TASK.getTask());
    check(restored.preferredComponents[0].variantId==="bottom-text"&&restored.brief.audience==="课题组","旧版草稿恢复时丢失变体或要求");
    const exportAgain=page.waitForEvent('download');await page.locator('#task-export').click();
    const regenerated=JSON.parse(fs.readFileSync(await(await exportAgain).path(),'utf8'));
    check(regenerated.documents[0].name==='旧任务原稿.docx','旧任务材料引用丢失');
    check(regenerated.handoff.execution==='agent-after-user-handoff'&&regenerated.handoff.components.length===catalog.components.length,'重新导出信任了过期交接信息');
    await expandComponentGroup(progressId);
    await page.locator(`[data-prefer-id="${progressId}"]`).uncheck();check(await page.locator("#task-fallback-count").textContent()===String(ids.length-1),"取消优先没有回到备用");
    await page.locator(`[data-prefer-id="${progressId}"]`).check();await page.locator(`[data-variant-for="${progressId}"]`).selectOption("all-text");
    check((await page.locator(`[data-component-id="${progressId}"] .priority-preview img`).getAttribute("src")).endsWith("__all-text.png"),"优先卡片缩略图与所选变体不一致");
    await page.locator(".priority-preview img").evaluateAll(imgs=>Promise.all(imgs.map(img=>img.decode())));
    await page.locator("#production-task").scrollIntoViewIfNeeded();
    check(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),"桌面任务面板横向溢出");
    await page.screenshot({path:path.join(root,"docs/qa/task-desktop.png"),fullPage:false});
    await page.setViewportSize({width:390,height:844});await page.locator("#task-selection-heading").count();
    await page.locator(".task-selection-heading").scrollIntoViewIfNeeded();
    check(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),"手机任务面板横向溢出");
    check(await paletteSelect.isVisible(),"窄屏配色入口不可见");
    await page.screenshot({path:path.join(root,"docs/qa/task-mobile.png")});
    await page.setViewportSize({width:1440,height:1100});
    for(const id of ids.slice(0,10)){
      const checkbox=page.locator(`[data-prefer-id="${id}"]`);
      if(!await checkbox.isChecked()){
        await expandComponentGroup(id);
        await checkbox.check();
      }
    }
    const finalExport=page.waitForEvent('download');await page.locator('#task-export').click();await finalExport;
    await page.locator('#production-task').scrollIntoViewIfNeeded();
    check(await page.locator('#task-export').isVisible(),'导出操作不清楚');
    await page.locator('#production-task').screenshot({path:path.join(root,'docs/qa/agent-handoff-desktop.png')});
    await page.setViewportSize({width:390,height:844});await page.locator('.task-handoff').scrollIntoViewIfNeeded();
    check(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),'多选后的手机页面横向溢出');
    await page.screenshot({path:path.join(root,'docs/qa/agent-handoff-mobile.png')});
    // Chromium 会中止转交给下载管理器的导航。仅排除已收到且通过哈希检查的下载。
    const failedResources = failed.filter(([url,error])=>error!=="net::ERR_ABORTED"||!downloadedTemplates.has(url)).map(([url])=>url);
    check(!errors.length,"页面错误："+errors.join(";"));check(!failedResources.length,"资源失败："+failedResources.join(";"));
    check(apiRequests.length===0,'组件库仍请求模型或任务接口：'+apiRequests.join(';'));
    fs.writeFileSync(path.join(root,"docs/task-check.json"),JSON.stringify({passed:true,checksPassed:checks,mode:'json-agent-handoff',pageErrors:errors,failedResources,apiRequests,exportBytes:fs.statSync(await download.path()).size,viewports:["1440x1000","1440x1100","390x844"]},null,2)+"\n");
    console.log(`制作任务检查通过：${checks}项，涵盖配色切换、母版下载、草稿恢复、多选排序、变体同步和JSON导出。`);
  }finally{await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
