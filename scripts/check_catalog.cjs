/* 可选的本地浏览检查：node check_catalog.cjs <playwright模块目录> <Edge或Chromium路径> */
const fs = require("node:fs");
const path = require("node:path");
const {pathToFileURL} = require("node:url");
const {chromium} = require(process.argv[2]);
const root = path.resolve(__dirname, "..");
const catalog = JSON.parse(fs.readFileSync(path.join(root,"library/components.json"),"utf8"));
const inventory = JSON.parse(fs.readFileSync(path.join(root,"library/inventory.json"),"utf8"));
const placeholders=JSON.parse(fs.readFileSync(path.join(root,'library/placeholders/manifest.json'),'utf8')).items;
let checks = 0;
function check(value, message) {if (!value) throw new Error(message); checks++;}
(async () => {
  const browser = await chromium.launch({headless:true, executablePath:process.argv[3]});
  const errors = [], failed = [], external = [];
  try {
    const page = await browser.newPage({viewport:{width:1920,height:910}});
    page.on("pageerror", e => errors.push(String(e)));
    page.on("requestfailed", r => failed.push(r.url()));
    page.on("request", r => {if (/^https?:/.test(r.url())) external.push(r.url());});
    await page.goto(pathToFileURL(path.join(root,"library/index.html")).href);
    check(await page.locator("#library-version,#component-task,#preview-note,#source-text,section.source").count()===0,
      "组件展示仍包含已删除的版本、摘要、预览说明或来源区");
    check(!(await page.locator("#catalog-detail").innerText()).includes("点击某一项，可以在统一样式预览上定位该区域"),
      "区域定位补充说明仍在展示");
    check(await page.locator("#component-nav button").count() === catalog.components.length, "组件导航数量错误");
    const reportPages = page.locator("#component-nav > .report-pages");
    const reportSummary = reportPages.locator(":scope > summary");
    const reportStyles = reportPages.locator(".report-styles > details");
    check((await reportSummary.textContent()).trim()==="汇报页", "五套导航页缺少统一入口");
    check(JSON.stringify((await reportStyles.locator(":scope > summary").allTextContents()).map(text=>text.trim()))===
      JSON.stringify(["汇报样式1","汇报样式2","汇报样式3","汇报样式4","汇报样式5"]), "汇报样式命名或顺序不正确");
    check(await reportStyles.locator("[data-component]").count()===15, "汇报页未收纳五套首页、目录与结束页");
    check(await page.locator("#component-nav [data-group][open]").count()===1&&
      await reportStyles.first().getAttribute("open")!==null, "初始侧栏应只展开首套样式");
    const icons = page.locator("#component-nav summary > svg.component-group-icon");
    check(await icons.count()===catalog.groups.filter(group=>group.kind!=="navigation").length+1&&await icons.evaluateAll(elements=>elements.every(icon=>
      icon.getAttribute("aria-hidden")==="true"&&icon.getAttribute("fill")==="none"&&getComputedStyle(icon).stroke==="rgb(51, 51, 51)")),
      "一级分组未使用黑白线条图标");
    check(await reportStyles.locator("summary > svg").count()===0, "汇报样式仍显示重复图标");
    const backSheet = reportSummary.locator(".component-group-icon-back");
    check(await backSheet.isVisible(), "展开态缺少图标的叠页提示");
    await reportSummary.click();
    check(await reportStyles.first().isHidden(), "汇报页无法收起全部套系");
    check(await backSheet.isHidden(), "收起态仍显示叠页提示");
    await reportSummary.press("Enter");
    check(await reportStyles.first().isVisible(), "键盘无法展开汇报页");
    const firstEntry = page.locator("#component-nav button[aria-current='true']");
    check(await firstEntry.textContent()===catalog.components[0].name&&await firstEntry.getAttribute("title")===catalog.components[0].id,
      "组件行未紧凑展示名称和悬停编号");
    check(await firstEntry.evaluate(el=>getComputedStyle(el).backgroundColor)==="rgb(231, 235, 237)", "当前组件缺少浅灰选中底色");
    const parallel = page.locator('#component-nav > details[data-group="parallel-two"]');
    const parallelSummary = parallel.locator(":scope > summary");
    check(await parallelSummary.evaluate(el=>getComputedStyle(el).fontSize)==="16px", "正文分类标题字号错误");
    await parallelSummary.click();
    check(await parallel.locator("button").first().isVisible(), "正文分类无法展开");
    await parallelSummary.click();
    check(await parallel.locator("button").first().isHidden(), "正文分类无法收起");
    const wideLayout = await page.evaluate(()=>{
      const bounds = selector => document.querySelector(selector).getBoundingClientRect();
      const {width,height} = bounds("#preview");
      return {
        fullWidth:Math.abs(bounds("main").right-innerWidth)<1,
        previewFills:Math.abs(width-bounds(".preview-toolbar").width)<1&&Math.abs(height-width*9/16)<1,
        lightScrollbars:[".sidebar","main"].every(selector=>{
          const style=getComputedStyle(document.querySelector(selector));
          return style.scrollbarWidth==="thin"&&style.scrollbarColor==="rgb(222, 223, 225) rgb(250, 251, 252)";
        }),
      };
    });
    check(wideLayout.fullWidth, "宽屏详情区右侧仍有页面留白");
    check(wideLayout.previewFills, "宽屏预览未等比占满内容区");
    check(wideLayout.lightScrollbars, "侧栏或详情区未采用浅灰细滚动条");
    await page.setViewportSize({width:1440,height:1000});
    const groupOrder = await page.locator("#component-nav [data-group]").evaluateAll(groups => groups.map(group => group.dataset.group));
    check(JSON.stringify(groupOrder) === JSON.stringify(catalog.groups.map(group => group.id)), "组件库分组次序错误");
    for (const group of catalog.groups) {
      const groupIds = await page.locator("[data-group='" + group.id + "'] [data-component]").evaluateAll(entries => entries.map(entry => entry.dataset.component));
      check(JSON.stringify(groupIds) === JSON.stringify(catalog.components.filter(c=>c.groupId===group.id).map(c=>c.id)), "组内组件归属错误");
    }
    for (const c of catalog.components) {
      const group = page.locator(`[data-group="${c.groupId}"]`);
      if (await group.getAttribute("open")===null) await group.locator(":scope > summary").click();
      await page.locator(`[data-id="${c.id}"]`).click();
      if (c.variants?.length) {
        check(await page.locator("#variant-select").inputValue()===(c.defaultVariant||""),"推荐的默认形式未生效");
        await page.locator("#variant-select").selectOption("");
      }
      await page.locator("#show-normalized").click();
      await page.locator("#slide-image").evaluate(img => img.decode());
      check(await page.locator("#component-name").textContent() === c.name, "组件切换标题错误");
      check(await page.locator("#slide-image").evaluate(img => img.naturalWidth === 1280 && img.naturalHeight === 720), "统一样式预览未加载");
      check(await page.locator("#preview-status").isHidden(), "正常组件预览仍显示补充说明");
      check(await page.locator("#regions .region-item").count() === c.regions.length, "区域拆解丢失");
      check(await page.locator("#slots tr").count() === c.slots.length, "替换槽位丢失");
      const href = await page.locator("#template-link").getAttribute("href");
      check(fs.existsSync(path.join(root,"library",href)), "原生母版下载链接失效");
      if (c.pageRole && c.referenceNumber >= 34) {
        const placeholder = placeholders.find(p => p.id === c.id);
        check(href === placeholder.templatePath && await page.locator("#slide-image").getAttribute("src") === placeholder.previewPath, "套系页未展示对应X占位母版");
      }
      if (c.referenceFile) {
        await page.locator("#show-reference").click();
        await page.locator("#slide-image").evaluate(img => img.decode());
        check((await page.locator("#slide-image").getAttribute("src")).includes(c.referenceFile), "原图切换错误");
      } else {
        check(await page.locator("#show-reference").isHidden(), "归档原图仍有展示入口");
      }
      await page.locator("#regions .region-item").first().click();
      check(await page.locator("#show-regions").isChecked(), "区域定位未开启标注");
      check(await page.locator(".region-box.selected").count() === 1, "区域定位标注错误");
      for (const variant of c.variants||[]) {
        await page.locator("#variant-select").selectOption(variant.id);
        await page.locator("#show-normalized").click();
        await page.locator("#slide-image").evaluate(img=>img.decode());
        const detail=inventory.components.find(x=>x.id===c.id).variants.find(v=>v.id===variant.id);
        const ph=placeholders.find(p=>p.id===c.id+'__'+variant.id);
        check((await page.locator("#slide-image").getAttribute("src"))===ph.previewPath,"变体没有切换对应占位预览");
        check((await page.locator("#template-link").getAttribute("href"))===ph.templatePath,"占位预览与下载文件不一致");
        check(await page.locator("#slots tr").count()===c.slots.filter(s=>!detail.replacedSlotKeys.includes(s.key)).length+detail.addedSlots.length,"变体替换槽位没有同步更新");
        check(await page.locator("#slide-image").evaluate(img=>img.naturalWidth===1280&&img.naturalHeight===720),"变体预览加载失败");
        await page.locator("#regions .region-item").first().click();
        check(await page.locator(".region-box.selected").count()===1,"变体分区定位失效");
        if (variant.id==="parallel") check(await page.locator("#regions .region-item:disabled").count()===1,"隐藏连线的区域仍可错误定位");
        await page.locator('#show-case').click();await page.locator('#slide-image').evaluate(img=>img.decode());
        check((await page.locator('#slide-image').getAttribute('src'))===detail.previewPath,'原案例对照丢失');
        await page.locator('#show-normalized').click();
      }
    }
    await page.locator('[data-id="BW-TOC-01"]').click();
    await page.locator("#toc-demo summary").click();
    await page.locator("#toc-count").selectOption("10");
    await page.locator("#toc-active").selectOption("10");
    check(await page.locator("#toc-active option").count() === 10, "十项目录当前章选项错误");
    check(await page.locator("#toc-diagram svg text").count() === 20, "十项目录条目或编号丢失");
    check(await page.locator("#toc-metrics").textContent().then(t=>t.includes("52.20") && t.includes("8.00")), "十项目录间距计算错误");
    await page.locator("#toc-count").selectOption("3");
    check(await page.locator("#toc-active option").count() === 3, "变更数量后当前章状态未重置");
    check(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth), "桌面视口存在横向溢出");
    await page.locator('[data-id="BW-BILATERAL-CENTER-01"]').click();
    await page.locator("#show-regions").uncheck();
    await page.locator("#preview").scrollIntoViewIfNeeded();
    await page.screenshot({path:path.join(root,"docs/qa/catalog-desktop.png")});
    await page.setViewportSize({width:390,height:844});
    await page.locator(".component-heading").scrollIntoViewIfNeeded();
    check(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth), "手机视口存在横向溢出");
    await page.screenshot({path:path.join(root,"docs/qa/catalog-mobile.png")});
    check(errors.length===0, "页面脚本错误："+errors.join("; "));
    check(failed.length===0, "本地资源加载失败："+failed.join("; "));
    check(external.length===0, "离线目录发起了外部请求");
    fs.writeFileSync(path.join(root,"docs/catalog-check.json"), JSON.stringify({passed:true, checksPassed:checks,
      browsers:"Headless Microsoft Edge", viewports:["1920x910","1440x1000","390x844"], pageErrors:errors, failedResources:failed, externalRequests:external},null,2)+"\n");
    console.log(`离线目录检查通过：${checks}项，${catalog.components.length}类组件、全部变体、原图对照和分区定位`);
  } finally {await browser.close();}
})().catch(e=>{console.error(e); process.exitCode=1;});
