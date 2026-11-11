"use strict";
(() => {
  const library = window.COMPONENT_LIBRARY;
  const get = id => document.getElementById(id);
  const peek = get("layout-peek");
  const main = document.querySelector("main");
  const nav = get("component-nav");
  let timer;
  function closePeek() {
    clearTimeout(timer);
    peek.hidden = true;
  }
  function showWorkspace() {
    const selecting = location.hash === "#selection";
    closePeek();
    document.body.classList.toggle("selecting", selecting);
    get("production-task").hidden = !selecting;
    get("catalog-detail").hidden = selecting;
    get("list-title").textContent = selecting ? "选择本次组件" : "全部组件";
    get("list-hint").hidden = !selecting;
    for (const link of document.querySelectorAll(".workspace-links a")) {
      if (link.hash === (selecting ? "#selection" : "#library")) link.setAttribute("aria-current", "page");
      else link.removeAttribute("aria-current");
    }
    main.scrollTop = 0;
  }
  function schedulePeek(entry) {
    closePeek();
    if (get("production-task").hidden) return;
    timer = setTimeout(() => {
      const id = entry.dataset.component;
      const component = library.catalog.components.find(c => c.id === id);
      const task = PPT_TASK.getTask();
      const choice = task.preferredComponents.find(c => c.componentId === id);
      const selected = (choice?.variantId || "auto") === "auto" ? component.defaultVariant : choice.variantId;
      const variant = component.variants?.find(v => v.id === selected);
      const asset = PPTTaskState.formAsset(library, task.paletteId, id, variant?.id || "base");
      const previewPath = asset?.rendered ? asset.previewPath
        : task.paletteId === "blue" && component.referenceFile ? "references/" + component.referenceFile : null;
      if (!previewPath) return;
      peek.querySelector("strong").textContent = component.name + (variant ? " · " + variant.name : "");
      peek.querySelector("img").src = previewPath;
      peek.hidden = false;
      const box = entry.getBoundingClientRect();
      const beside = box.right + 12 + peek.offsetWidth <= innerWidth;
      peek.style.left = (beside ? box.right + 12 : Math.max(8, Math.min(box.left, innerWidth - peek.offsetWidth - 8))) + "px";
      peek.style.top = Math.max(8, Math.min(beside ? box.top : box.bottom + 6, innerHeight - peek.offsetHeight - 8)) + "px";
    }, 180);
  }
  for (const entry of nav.querySelectorAll(".component-nav-entry")) {
    entry.addEventListener("pointerenter", e => {if (e.pointerType !== "touch") schedulePeek(entry)});
    entry.addEventListener("pointerleave", closePeek);
    entry.addEventListener("focusin", e => {if (e.target.matches("input:focus-visible")) schedulePeek(entry)});
    entry.addEventListener("focusout", closePeek);
  }
  function updateCount() { get("nav-count").textContent = PPT_TASK.getTask().preferredComponents.length }
  nav.addEventListener("toggle", e => {if (!e.target.open) closePeek();}, true);
  document.addEventListener("ppt-task-change", () => {updateCount(); closePeek();});
  document.addEventListener("keydown", e => {if (e.key === "Escape") closePeek()});
  document.addEventListener("scroll", () => {if (!peek.hidden) closePeek()}, true);
  window.addEventListener("resize", closePeek);
  window.addEventListener("hashchange", showWorkspace);
  updateCount();
  showWorkspace();
})();
