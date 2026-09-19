/* ppt-studio 本地查看器前端逻辑 */
const $ = (id) => document.getElementById(id);

const state = {
  pageCount: 0,
  current: 1,
  hasQA: false,
  exporting: false,
  saving: new Set(),
};

function toast(msg, isError = false, ms = 2600) {
  const el = $("toast");
  el.textContent = msg;
  el.classList.toggle("error", isError);
  el.hidden = false;
  clearTimeout(el._t);
  el._t = setTimeout(() => (el.hidden = true), ms);
}

function setHint(msg) { $("hint").textContent = msg || ""; }

async function api(path, init) {
  const r = await fetch(path, init);
  const data = await r.json().catch(() => ({ error: "响应解析失败" }));
  if (!r.ok) throw new Error(data.error || `HTTP ${r.status}`);
  return data;
}

/* ---------- 初始化 ---------- */
async function init() {
  try {
    const info = await api("/api/project");
    state.pageCount = info.pageCount;
    state.hasQA = info.hasQA;
    $("deck-title").textContent = info.title || info.deck;
    if (!state.hasQA) setHint("还没有预览图，点「重新导出预览」生成");
    renderThumbs();
    await loadPage(1);
  } catch (e) {
    toast(`项目加载失败: ${e.message}`, true, 6000);
  }
}

/* ---------- 缩略图 ---------- */
function renderThumbs() {
  const box = $("thumbs");
  box.replaceChildren();
  for (let n = 1; n <= state.pageCount; n++) {
    const fig = document.createElement("figure");
    fig.className = "thumb";
    fig.dataset.page = n;
    if (state.hasQA) {
      const img = document.createElement("img");
      img.src = `/api/image/${n}.png`;
      img.loading = "lazy";
      img.alt = `第 ${n} 页`;
      fig.append(img);
    } else {
      const miss = document.createElement("div");
      miss.className = "thumb-miss";
      miss.textContent = `${n} 未导出`;
      fig.append(miss);
    }
    const cap = document.createElement("figcaption");
    cap.textContent = n;
    fig.append(cap);
    fig.addEventListener("click", () => loadPage(n));
    box.append(fig);
  }
}

/* ---------- 单页 ---------- */
async function loadPage(n) {
  n = Math.min(Math.max(1, n), state.pageCount || 1);
  state.current = n;
  $("page-indicator").textContent = `${n} / ${state.pageCount || "–"}`;
  $("btn-prev").disabled = n <= 1;
  $("btn-next").disabled = n >= state.pageCount;

  const img = $("page-image");
  if (state.hasQA) {
    img.src = `/api/image/${n}.png?t=${Date.now()}`; // 防缓存，配合重新导出
    img.hidden = false;
    $("canvas-empty").hidden = true;
  } else {
    img.hidden = true;
    $("canvas-empty").hidden = false;
  }

  // 高亮缩略图
  document.querySelectorAll(".thumb").forEach((el) =>
    el.classList.toggle("active", Number(el.dataset.page) === n));

  await loadTexts(n);
}

/* ---------- 文本面板 ---------- */
async function loadTexts(page) {
  const box = $("text-list");
  box.replaceChildren();
  $("page-type").textContent = "";
  let data;
  try {
    data = await api(`/api/texts?page=${page}`);
  } catch (e) {
    toast(`读取文本失败: ${e.message}`, true);
    return;
  }
  const pageInfo = data.pages[0];
  if (!pageInfo) return;
  $("page-type").textContent = pageInfo.pageType ? `类型: ${pageInfo.pageType}` : "";
  const texts = pageInfo.texts || [];
  if (texts.length === 0) {
    const p = document.createElement("p");
    p.className = "panel-empty";
    p.textContent = "本页没有文本元素";
    box.append(p);
    return;
  }
  for (const t of texts) box.append(buildTextItem(page, t));
}

function buildTextItem(page, t) {
  const item = document.createElement("div");
  item.className = "text-item";

  const head = document.createElement("button");
  head.className = "text-item-head";
  const eid = document.createElement("span");
  eid.className = "eid";
  eid.textContent = t.elementId;
  const preview = document.createElement("span");
  preview.className = "preview";
  preview.textContent = t.preview || "（空文本）";
  head.append(eid, preview);

  const body = document.createElement("div");
  body.className = "text-item-body";
  body.hidden = true;
  const ta = document.createElement("textarea");
  ta.value = t.rawText ?? "";
  if (t.rich) {
    const note = document.createElement("div");
    note.className = "rich-note";
    note.textContent = "该元素含富文本标签（<p>/<span> 等），编辑时保留标签结构";
    body.append(note);
  }
  body.append(ta);
  const actions = document.createElement("div");
  actions.className = "text-item-actions";
  const save = document.createElement("button");
  save.className = "btn";
  save.textContent = "保存";
  const cancel = document.createElement("button");
  cancel.className = "btn";
  cancel.textContent = "收起";
  actions.append(save, cancel);
  body.append(actions);

  head.addEventListener("click", () => {
    body.hidden = !body.hidden;
    item.classList.toggle("open", !body.hidden);
    if (!body.hidden) ta.focus();
  });

  save.addEventListener("click", async () => {
    if (state.saving.has(t.elementId)) return;
    state.saving.add(t.elementId);
    save.disabled = true;
    save.textContent = "保存中…";
    try {
      await api("/api/update-text", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ page, elementId: t.elementId, text: ta.value }),
      });
      toast("已写入 .page，重新导出后可见");
      preview.textContent = ta.value.replace(/<[^>]+>/g, "").trim().slice(0, 100) || "（空文本）";
    } catch (e) {
      toast(`保存失败: ${e.message}`, true);
    } finally {
      save.disabled = false;
      save.textContent = "保存";
      state.saving.delete(t.elementId);
    }
  });
  cancel.addEventListener("click", () => {
    body.hidden = true;
    item.classList.remove("open");
  });

  item.append(head, body);
  return item;
}

/* ---------- 导出 ---------- */
async function exportPreview() {
  if (state.exporting) return;
  state.exporting = true;
  const btn = $("btn-export");
  btn.disabled = true;
  setHint("正在调用本机 PowerPoint 渲染…");
  try {
    const r = await api("/api/export", { method: "POST" });
    if (r.stale) toast(r.hint, true, 8000);
    else toast("导出完成");
    const info = await api("/api/project");
    const wasQA = state.hasQA;
    state.hasQA = info.hasQA;
    if (!wasQA && state.hasQA) renderThumbs();
    await loadPage(state.current);
  } catch (e) {
    toast(`导出失败: ${e.message}`, true, 6000);
  } finally {
    state.exporting = false;
    btn.disabled = false;
    setHint("");
  }
}

/* ---------- 事件 ---------- */
$("btn-prev").addEventListener("click", () => loadPage(state.current - 1));
$("btn-next").addEventListener("click", () => loadPage(state.current + 1));
$("btn-export").addEventListener("click", exportPreview);
document.addEventListener("keydown", (e) => {
  if (e.target.matches("textarea, input")) return;
  if (e.key === "ArrowLeft") loadPage(state.current - 1);
  if (e.key === "ArrowRight") loadPage(state.current + 1);
});

init();
