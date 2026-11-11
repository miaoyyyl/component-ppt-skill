"use strict";
(() => {
  const library = window.COMPONENT_LIBRARY, catalog = library.catalog, state = window.PPTTaskState;
  const byId = id=>document.getElementById(id), storageKey="ppt-component-production-tasks.v1";
  const fields = ["title","audience","durationMinutes","pageRange","mustKeep","sourceNotes","researchScope"];
  const paletteSelect = byId("palette-select");
  for (const scheme of catalog.palettes) paletteSelect.add(new Option(scheme.name, scheme.id));
  let task = state.create(catalog), tasks = [], timer, storageIssue = "";
  function message(value, error=false) {
    byId("task-message").textContent = value;
    byId("task-message").classList.toggle("error",error);
  }
  try {
    const raw = localStorage.getItem(storageKey);
    if (raw) {
      const store=JSON.parse(raw);
      if (store.version!==1 || !Array.isArray(store.tasks)) throw new Error("本机任务列表格式不正确");
      tasks=store.tasks.map(t=>state.validate(t,catalog));
      task=tasks.find(t=>t.id===store.activeTaskId)||tasks[0]||task;
    }
  } catch (error) {storageIssue=`本机草稿未能读取：${error.message}。可继续选型并导出 JSON。`;}
  function remember() {
    const index=tasks.findIndex(t=>t.id===task.id);
    if (index<0) tasks.push(structuredClone(task)); else tasks[index]=structuredClone(task);
  }
  function persist(explicit=false) {
    clearTimeout(timer);
    const oldSavedAt=task.savedAt;
    if (explicit) task.savedAt=new Date().toISOString();
    remember();
    try {
      localStorage.setItem(storageKey,JSON.stringify({version:1,activeTaskId:task.id,tasks}));
      if (explicit) message("任务已保存到本机");
    } catch (error) {
      task.savedAt=oldSavedAt;
      remember();
      message("本机存储不可用或空间不足，请导出 JSON 保存任务。",true);
    }
  }
  function changed(render=true) {
    task.updatedAt=new Date().toISOString();
    task=state.validate(task,catalog);
    if(byId('task-message').classList.contains('error'))message('');
    remember();
    clearTimeout(timer);timer=setTimeout(()=>persist(),300);
    if (render) renderSelection();
    document.dispatchEvent(new CustomEvent("ppt-task-change",{detail:{task:structuredClone(task)}}));
  }
  function preview(componentId,variantId) {
    document.dispatchEvent(new CustomEvent("ppt-task-preview",{detail:{componentId,variantId,revealInNav:true}}));
    location.hash="library";
  }
  function action(text,label,callback) {
    const button=document.createElement("button");button.type="button";button.className="quiet-button";button.textContent=text;button.setAttribute("aria-label",label);button.addEventListener("click",callback);return button;
  }
  function renderPalette() {
    paletteSelect.value = task.paletteId;
    const primary = catalog.palettes.find(scheme => scheme.id === task.paletteId).primary;
    paletteSelect.closest(".palette-control").style.setProperty("--presentation-primary", "#" + primary);
  }
  function renderSelection() {
    byId("task-preferred-count").textContent=String(task.preferredComponents.length);
    byId("task-fallback-count").textContent=String(task.fallback.componentIds.length);
    const list=byId("task-preferred-list");list.replaceChildren();
    byId("task-empty").hidden=task.preferredComponents.length>0;
    task.preferredComponents.forEach((choice,index)=>{
      const component=catalog.components.find(c=>c.id===choice.componentId), item=document.createElement("li");item.dataset.componentId=component.id;
      const head=document.createElement("div");head.className="priority-card-heading";
      const rank=document.createElement("span");rank.className="priority-rank";rank.textContent=String(index+1).padStart(2,"0");
      const name=document.createElement("strong");name.textContent=component.name;head.append(rank,name);
      const image=document.createElement("img");
      const variantId = (choice.variantId === "auto" ? component.defaultVariant : choice.variantId) || "base";
      const asset = state.formAsset(library, task.paletteId, component.id, variantId);
      if (asset?.rendered) image.src = asset.previewPath;
      else if (task.paletteId === "blue" && component.referenceFile) image.src = `references/${component.referenceFile}`;
      else image.hidden = true;
      image.alt=component.name+"的所选形式预览";image.loading="lazy";
      const previewButton=action("预览",`预览${component.name}`,()=>preview(component.id,choice.variantId));previewButton.classList.add("priority-preview");previewButton.prepend(image);
      const label=document.createElement("label");label.textContent="本次形式";
      const select=document.createElement("select");select.dataset.variantFor=component.id;select.setAttribute("aria-label",component.name+"的本次变体");
      for (const [id,text] of [["auto","由模型按内容选择"],["base","参考结构"],...(component.variants||[]).map(v=>[v.id,v.name])]) {
        const option=document.createElement("option");option.value=id;option.textContent=text;select.append(option);
      }
      select.value=choice.variantId;
      select.addEventListener("change",()=>{task=state.setVariant(task,component.id,select.value,catalog);changed();document.dispatchEvent(new CustomEvent("ppt-task-preview",{detail:{componentId:component.id,variantId:select.value}}));byId("task-preferred-list").querySelector(`[data-variant-for="${component.id}"]`).focus();});
      label.append(select);
      const actions=document.createElement("div");actions.className="priority-actions";
      for (const [offset,text] of [[-1,"上移"],[1,"下移"]]) {
        const button=action(text,`${text}${component.name}`,()=>{
          task=state.move(task,component.id,offset,catalog);changed();
          const card=list.querySelector(`[data-component-id="${component.id}"]`), target=card.querySelector(`button[data-move="${offset}"]`);
          (target.disabled?card.querySelector("select"):target).focus();
        });button.dataset.move=String(offset);button.disabled=index+offset<0||index+offset>=task.preferredComponents.length;actions.append(button);
      }
      actions.append(action("移至备用",`将${component.name}移至备用`,()=>{task=state.setPreferred(task,component.id,false,catalog);changed();}));
      item.append(head,previewButton,label,actions);list.append(item);
    });
    byId("task-fallback-names").textContent=task.fallback.componentIds.map(id=>catalog.components.find(c=>c.id===id).name).join("、")||"全部入库组件均已列为优先。";
  }
  function renderForm() {
    for (const key of fields) byId(`task-${key}`).value=task.brief[key]??"";
    renderPalette();
    renderSelection();
    document.dispatchEvent(new CustomEvent("ppt-task-change",{detail:{task:structuredClone(task)}}));
  }
  function adopt(value) {task=state.validate(value,catalog);changed();renderForm();}
  window.PPT_TASK={
    getTask:()=>structuredClone(task),
    isPreferred:id=>task.preferredComponents.some(c=>c.componentId===id),
    toggle(id,enabled,variantId="auto") {task=state.setPreferred(task,id,enabled,catalog,variantId);changed();},
    setVariant(id,variantId) {task=state.setVariant(task,id,variantId,catalog);changed();},
    setDocuments(documents) {task.documents=documents;changed(false);},
    message, persist, adopt
  };
  paletteSelect.addEventListener("change", () => {
    task.paletteId = paletteSelect.value;
    renderPalette();
    changed();
  });
  for (const key of fields) byId(`task-${key}`).addEventListener("input",event=>{
    const value=event.target.value;
    if (key==="durationMinutes" && (value!=="" && (!Number.isFinite(Number(value)) || Number(value)<1 || Number(value)>1440))) {message("汇报时长应为1至1440分钟，或留空。",true);return;}
    task.brief[key]=key==="durationMinutes"?(value===""?null:Number(value)):value;
    changed(false);
  });
  byId("task-new").addEventListener("click",()=>{
    persist();
    task = {...state.create(catalog), paletteId:task.paletteId};
    changed();renderForm();message("已新建任务");
  });
  byId("task-export").addEventListener("click",()=>{
    try {
      task=state.validate(task,catalog);persist();
      const output=state.forAgent(task,library);
      const blob=new Blob([JSON.stringify(output,null,2)+"\n"],{type:"application/json;charset=utf-8"}), url=URL.createObjectURL(blob), link=document.createElement("a");
      link.href=url;link.download=(task.brief.title.trim()||"制作任务").replace(/[<>:"/\\|?*\u0000-\u001f]/g,"_").slice(0,80)+".ppt-task.json";
      link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);message("已导出任务 JSON");
    }catch(error){message('任务未能导出：'+error.message,true);}
  });
  window.addEventListener("pagehide",()=>persist());
  renderForm();
  if(storageIssue)message(storageIssue,true);
})();
