"use strict";
(() => {
  const library = window.COMPONENT_LIBRARY, {catalog, theme, inventory} = library;
  const state = window.PPTTaskState;
  const byId = id => document.getElementById(id);
  let current = catalog.components[0], referenceMode = false, caseMode=false, selectedRegion = null;
  let paletteId = window.PPT_TASK.getTask().paletteId, previewAsset = null;
  const variantChoices = new Map();
  const componentDetail = () => inventory.components.find(c=>c.id===current.id);
  function activeVariant() {
    const key=variantChoices.has(current.id)?variantChoices.get(current.id):current.defaultVariant;
    return (componentDetail().variants||[]).find(v=>v.id===key)||null;
  }
  function effectiveSlots() {
    const variant=activeVariant();
    return variant ? current.slots.filter(s=>!variant.replacedSlotKeys.includes(s.key)).concat(variant.addedSlots) : current.slots;
  }
  const nav = byId("component-nav");
  const groupIcon = byId("component-group-icon").content.firstElementChild;
  function createNavGroup(name, showIcon = true) {
    const group = document.createElement("details");
    group.className = "component-nav-group";
    const summary = document.createElement("summary");
    const label = document.createElement("span");
    label.textContent = name;
    if (showIcon) summary.append(groupIcon.cloneNode(true));
    summary.append(label);
    group.append(summary);
    return group;
  }
  const reportPages = createNavGroup("汇报页");
  reportPages.classList.add("report-pages");
  reportPages.open = true;
  const reportStyles = document.createElement("div");
  reportStyles.className = "report-styles";
  reportPages.append(reportStyles);
  nav.append(reportPages);
  let reportStyleNumber = 0;
  for (const group of catalog.groups) {
    const isReportStyle = group.kind === "navigation";
    const section = createNavGroup(isReportStyle ? `汇报样式${++reportStyleNumber}` : group.name, !isReportStyle);
    section.dataset.group = group.id;
    section.open = group.id === current.groupId;
    const entries = document.createElement("div");
    entries.className = "component-nav-group-entries";
    for (const c of catalog.components.filter(component => component.groupId === group.id)) {
      const navName = c.name.startsWith(group.name + " · ") ? c.name.slice(group.name.length + 3) : c.name;
      const entry=document.createElement("div");entry.className="component-nav-entry";
      const checkbox=document.createElement("input");checkbox.type="checkbox";checkbox.dataset.preferId=c.id;checkbox.setAttribute("aria-label",`优先使用${c.name}`);
      checkbox.checked=window.PPT_TASK.isPreferred(c.id);
      checkbox.addEventListener("change",()=>window.PPT_TASK.toggle(c.id,checkbox.checked));
      const button = document.createElement("button");
      button.textContent = navName;
      button.title = c.id;
      button.type = "button";
      button.dataset.id = c.id;
      button.addEventListener("click", () => {current = c; selectedRegion = null; render(); document.querySelector("main").scrollTop=0;});
      checkbox.id="select-"+c.id;
      const label=document.createElement("label");label.htmlFor=checkbox.id;label.textContent=navName;
      entry.dataset.component=c.id;entry.append(checkbox,label,button);entries.append(entry);
    }
    section.append(entries);
    (isReportStyle ? reportStyles : nav).append(section);
  }
  function renderPalette() {
    const palette = (state.paletteTheme(library, paletteId) || theme).palette;
    byId("palette").replaceChildren(...["primary", "white", "inactive", "emphasisRed", "emphasisYellow"].map(key => {
      const swatch = document.createElement("span"); swatch.className = "swatch";
      swatch.style.background = `#${palette[key]}`; swatch.title = `${key} #${palette[key]}`;
      return swatch;
    }));
  }
  function list(id, values) {
    byId(id).replaceChildren(...values.map(text => {const li = document.createElement("li"); li.textContent = text; return li;}));
  }
  function regionBounds(region) {
    const variant=activeVariant(), detail=variant||componentDetail();
    const override=variant?.regionShapeIds[region.id];
    const ids = new Set(override || region.shapeIds.concat(region.generatedShapeIds||[]));
    const elements = detail.templateElements.filter(s => current.pageRole === "toc" && region.id === "rows" ? s.id >= 20000 : ids.has(s.id));
    const boxes = elements.map(s => s.boxPx).filter(Boolean);
    if (!boxes.length) return null;
    const left = Math.max(0, Math.min(...boxes.map(b => b[0]))), top = Math.max(0, Math.min(...boxes.map(b => b[1])));
    return [left, top, Math.min(1280, Math.max(...boxes.map(b => b[0] + b[2]))) - left,
      Math.min(720, Math.max(...boxes.map(b => b[1] + b[3]))) - top];
  }
  function overlay() {
    const container = byId("region-overlay"); container.replaceChildren();
    if (referenceMode || !previewAsset?.rendered || !byId("show-regions").checked) return;
    current.regions.forEach((r, index) => {
      const b = regionBounds(r); if (!b) return;
      const box = document.createElement("div"); box.className = "region-box" + (selectedRegion === r.id ? " selected" : "");
      Object.assign(box.style, {left:`${b[0]/12.8}%`, top:`${b[1]/7.2}%`, width:`${b[2]/12.8}%`, height:`${b[3]/7.2}%`});
      const label = document.createElement("span"); label.textContent = `${index+1}. ${r.name}`; box.append(label);
      box.addEventListener("click", () => {selectedRegion = r.id; renderRegions(); overlay();});
      container.append(box);
    });
  }
  function renderRegions() {
    byId("regions").replaceChildren(...current.regions.map((r, index) => {
      const button = document.createElement("button"); button.type = "button";
      button.className = "region-item" + (selectedRegion === r.id ? " active" : "");
      const title = document.createElement("strong"); title.textContent = `${index+1}. ${r.name}`;
      const fn = document.createElement("p"); fn.textContent = r.function;
      const hidden=activeVariant()?.regionShapeIds[r.id]?.length===0;
      const replace = document.createElement("p"); replace.className = "replacement"; replace.textContent = hidden ? "当前变体不显示此区域。" : `替换：${r.replace}`;
      button.disabled=hidden;
      button.append(title, fn, replace);
      button.addEventListener("click", () => {selectedRegion = r.id; referenceMode = false; byId("show-regions").checked = true; updatePreview(); renderRegions();});
      return button;
    }));
  }
  function updatePreview() {
    const image = byId("slide-image");
    byId("show-reference").hidden = !current.referenceFile;
    if (!current.referenceFile) referenceMode = false;
    const variant = activeVariant();
    previewAsset = state.formAsset(library, paletteId, current.id, variant?.id || "base", referenceMode || caseMode ? "case" : "placeholder");
    const showingReference = current.referenceFile && (referenceMode || (paletteId === "blue" && !previewAsset?.rendered));
    const previewPath = showingReference ? `references/${current.referenceFile}` : previewAsset?.rendered ? previewAsset.previewPath : null;
    image.hidden = !previewPath;
    if (previewPath) image.src = previewPath;
    else image.removeAttribute("src");
    image.alt = `${current.name}：${showingReference ? "用户原始参考" : (variant?variant.name:"统一样式后的原生PPT渲染")}`;
    byId("show-normalized").setAttribute("aria-pressed", String(!referenceMode&&!caseMode));
    byId("show-case").setAttribute("aria-pressed",String(!referenceMode&&caseMode));
    byId("show-reference").setAttribute("aria-pressed", String(referenceMode));
    const status = byId("preview-status");
    const pendingRender = !referenceMode && !previewAsset?.rendered;
    status.hidden = !pendingRender;
    status.textContent = pendingRender
      ? (paletteId !== "blue" ? "当前配色母版待更新" : showingReference ? "当前显示原始参考图，母版预览待重新渲染" : "母版预览待重新渲染")
      : "";
    const download = byId("template-link");
    download.hidden = !previewAsset;
    if (previewAsset) {
      download.href = previewAsset.templatePath;
      download.textContent = previewAsset.usesXPlaceholder ? "下载 X 占位母版" : "下载原生 PPT 母版";
    }
    overlay();
  }
  function render() {
    for (const button of nav.querySelectorAll("button[data-id]")) button.setAttribute("aria-current", String(button.dataset.id === current.id));
    byId("component-id").textContent = `${current.id} / ${window.PPT_TASK.isPreferred(current.id)?"本次优先":"本次备用"}`;
    byId("component-name").textContent = current.name;
    byId("argument-text").textContent = current.argument;
    const variant=activeVariant();
    byId("variant-toolbar").hidden=!(current.variants||[]).length;
    byId("variant-select").replaceChildren();
    for (const [key,name] of [["","参考结构"],...(current.variants||[]).map(v=>[v.id,v.name])]) {
      const option=document.createElement("option");option.value=key;option.textContent=name;byId("variant-select").append(option);
    }
    byId("variant-select").value=variant?variant.id:"";
    byId("adaptation-section").hidden=!(current.adaptationRules||[]).length;
    list("adaptation-rules",current.adaptationRules||[]);
    list("fits", current.fits); list("fixed", current.fixed);
    renderRegions();
    byId("slots").replaceChildren(...effectiveSlots().map(s => {
      const tr = document.createElement("tr");
      const region = current.regions.find(r => r.id===s.regionId || r.shapeIds.some(id => s.shapeIds.includes(id)));
      for (const text of [s.key, region ? region.name : "目录状态", s.replace]) {const td = document.createElement("td"); td.textContent = text; tr.append(td);}
      return tr;
    }));
    byId("toc-demo").hidden = current.id !== "BW-TOC-01";
    updatePreview();
  }
  byId('show-normalized').addEventListener('click',()=>{referenceMode=false;caseMode=false;updatePreview();});
  byId('show-case').addEventListener('click',()=>{referenceMode=false;caseMode=true;updatePreview();});
  byId('show-reference').addEventListener('click',()=>{referenceMode=true;updatePreview();});
  byId("show-regions").addEventListener("change", overlay);
  byId("variant-select").addEventListener("change",()=>{
    const value=byId("variant-select").value;variantChoices.set(current.id,value);
    selectedRegion=null;render();
  });
  document.addEventListener("ppt-task-change",event=>{
    if (paletteId !== event.detail.task.paletteId) {
      paletteId = event.detail.task.paletteId;
      referenceMode = false;
      renderPalette();
      tocDemo(false);
    }
    for(const input of nav.querySelectorAll("input[data-prefer-id]"))input.checked=window.PPT_TASK.isPreferred(input.dataset.preferId);
    for(const choice of event.detail.task.preferredComponents)variantChoices.set(choice.componentId,choice.variantId==="auto"?catalog.components.find(c=>c.id===choice.componentId).defaultVariant||"":choice.variantId==="base"?"":choice.variantId);
    render();
  });
  document.addEventListener("ppt-task-preview",event=>{
    const choice=event.detail;current=catalog.components.find(c=>c.id===choice.componentId);
    variantChoices.set(current.id,choice.variantId==="auto"?current.defaultVariant||"":choice.variantId==="base"?"":choice.variantId);
    selectedRegion=null;referenceMode=false;render();
    if (choice.revealInNav) {
      const entry = nav.querySelector(`[data-component="${current.id}"]`);
      for (let group = entry.closest("details"); group; group = group.parentElement.closest("details")) {
        group.open = true;
      }
      entry.scrollIntoView({block:"nearest"});
    }
  });
  for(const choice of window.PPT_TASK.getTask().preferredComponents)variantChoices.set(choice.componentId,choice.variantId==="auto"?catalog.components.find(c=>c.id===choice.componentId).defaultVariant||"":choice.variantId==="base"?"":choice.variantId);
  for (let n=1;n<=10;n++) {const option = document.createElement("option"); option.value = n; option.textContent = `${n} 项`; byId("toc-count").append(option);}
  byId("toc-count").value = "6";
  function tocDemo(reset) {
    const n = Number(byId("toc-count").value), rule = theme.toc;
    const palette = (state.paletteTheme(library, paletteId) || theme).palette;
    if (reset) {
      byId("toc-active").replaceChildren();
      for (let i=1;i<=n;i++) {const option = document.createElement("option"); option.value = i; option.textContent = `第 ${i} 章`; byId("toc-active").append(option);}
    }
    const active = Number(byId("toc-active").value || 1);
    const h = Math.min(rule.preferredRowHeight, (rule.panelHeight-(n+1)*rule.minimumGap)/n);
    const gap = (rule.panelHeight-n*h)/(n+1);
    byId("toc-metrics").textContent = `行高 ${h.toFixed(2)} px；顶部、底部和行间净间距均为 ${gap.toFixed(2)} px`;
    const svgNS = "http://www.w3.org/2000/svg";
    const svg = document.createElementNS(svgNS,"svg"); svg.setAttribute("viewBox","0 0 1280 720");
    const add = (type, attrs, text) => {const el = document.createElementNS(svgNS,type); for (const [k,v] of Object.entries(attrs)) el.setAttribute(k,v); if(text !== undefined) el.textContent=text; svg.append(el); return el;};
    const asset = state.formAsset(library, paletteId, "BW-TOC-01", "base", "case");
    if (asset?.rendered) add("image",{href:asset.previewPath,x:0,y:0,width:1280,height:720});
    add("rect",{x:490,y:55,width:625,height:610,fill:"white"});
    for (let i=0;i<n;i++) {
      const y=rule.panelY+gap+i*(h+gap), color="#" + (i+1===active?palette.primary:palette.inactive);
      add("rect",{x:500.27,y,width:595.74,height:h,rx:h/2,fill:i+1===active?"#"+palette.surface:"white",stroke:color,"stroke-width":.7});
      add("line",{x1:627.63,y1:y+8,x2:627.63,y2:y+h-8,stroke:color,"stroke-width":.5});
      add("text",{x:574.3,y:y+h/2,fill:color,"text-anchor":"middle","dominant-baseline":"central","font-size":42.6667,"font-weight":700,"font-family":"Microsoft YaHei"},String(i+1).padStart(2,"0"));
      add("text",{x:641.95,y:y+h/2,fill:color,"dominant-baseline":"central","font-size":37.3333,"font-weight":700,"font-family":"Microsoft YaHei"},`第${i+1}部分的章节名称`);
    }
    byId("toc-diagram").replaceChildren(svg);
  }
  byId("toc-count").addEventListener("change",()=>tocDemo(true));
  byId("toc-active").addEventListener("change",()=>tocDemo(false));
  renderPalette(); tocDemo(true); render();
})();
