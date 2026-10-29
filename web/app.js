const picker = document.getElementById("picker");
const drop = document.getElementById("drop");
const list = document.getElementById("list");
const runBtn = document.getElementById("run");
const statusEl = document.getElementById("status");
const result = document.getElementById("result");
const download = document.getElementById("download");
const downloadPdf = document.getElementById("download-pdf");
const reportBody = document.getElementById("report");
const logEl = document.getElementById("log");

const FONT_FILES = [
  "Arimo-Regular.ttf",
  "Arimo-Bold.ttf",
  "Carlito-Regular.ttf",
  "Carlito-Bold.ttf",
  "Tinos-Regular.ttf",
  "Tinos-Bold.ttf",
  "DejaVuSans.ttf",
  "DejaVuSans-Bold.ttf",
  "NotoSansGujarati-Regular.ttf",
  "NotoSansGujarati-Bold.ttf",
  "NotoSansDevanagari-Regular.ttf",
  "NotoSansDevanagari-Bold.ttf",
];

// Family names must match FONTS in pdf_report.py
const INDIC_FONTS = {
  "NotoSansGujarati-Regular.ttf": { family: "NotoGujarati", weight: "400" },
  "NotoSansGujarati-Bold.ttf": { family: "NotoGujarati", weight: "700" },
  "NotoSansDevanagari-Regular.ttf": { family: "NotoDevanagari", weight: "400" },
  "NotoSansDevanagari-Bold.ttf": { family: "NotoDevanagari", weight: "700" },
};

let files = [];
let pyodide = null;

// ---------------------------------------------------------
// GUJARATI / HINDI TEXT FOR THE PDF
// ---------------------------------------------------------
// Pyodide cannot shape Indic scripts (vowel signs, conjuncts), but the
// browser can. pdf_report.py calls this to draw such text as a PNG.
// All sizes in and out are PDF points.
const RENDER_SCALE = 8; // pixels per point (~576 dpi)
const indicCanvas = document.createElement("canvas");

window.renderIndic = (text, family, bold, sizePt, r, g, b) => {
  const ctx = indicCanvas.getContext("2d");
  const px = sizePt * RENDER_SCALE;
  const font = `${bold ? "700" : "400"} ${px}px "${family}"`;

  ctx.font = font;
  const m = ctx.measureText(text);
  const ascent = Math.ceil(m.fontBoundingBoxAscent || px * 1.1);
  const descent = Math.ceil(m.fontBoundingBoxDescent || px * 0.5);

  indicCanvas.width = Math.max(1, Math.ceil(m.width));
  indicCanvas.height = ascent + descent;

  ctx.font = font; // resizing the canvas resets its state
  ctx.fillStyle = `rgb(${r}, ${g}, ${b})`;
  ctx.textBaseline = "alphabetic";
  ctx.fillText(text, 0, ascent);

  const base64 = indicCanvas.toDataURL("image/png").split(",")[1];
  const png = Uint8Array.from(atob(base64), (ch) => ch.charCodeAt(0));

  return {
    png,
    width: m.width / RENDER_SCALE, // text advance
    imageWidth: indicCanvas.width / RENDER_SCALE,
    height: indicCanvas.height / RENDER_SCALE,
    ascent: ascent / RENDER_SCALE,
  };
};

function setStatus(text, isError = false) {
  statusEl.textContent = text;
  statusEl.classList.toggle("err", isError);
}

function log(line) {
  logEl.textContent += line + "\n";
}

function refreshButton() {
  runBtn.disabled = !pyodide || files.length === 0;
}

// ---------------------------------------------------------
// LOAD PYTHON ENGINE
// ---------------------------------------------------------
async function loadEngine() {
  try {
    pyodide = await loadPyodide();
    await pyodide.loadPackage("micropip");
    await pyodide.pyimport("micropip").install(["openpyxl", "xlrd==2.0.1", "fpdf2==2.8.1"]);

    for (const name of ["gazette.py", "pdf_report.py", "contents.py"]) {
      const code = await (await fetch(name, { cache: "no-cache" })).text();
      pyodide.FS.writeFile(name, code);
    }

    const fontDir = pyodide.FS.cwd() + "/fonts";
    pyodide.FS.mkdirTree(fontDir);
    for (const name of FONT_FILES) {
      const res = await fetch("fonts/" + name);
      if (!res.ok) throw new Error("Missing font: " + name);
      const data = await res.arrayBuffer();
      pyodide.FS.writeFile(fontDir + "/" + name, new Uint8Array(data));

      // Gujarati/Hindi fonts are also registered with the page for renderIndic
      const indic = INDIC_FONTS[name];
      if (indic) {
        const face = new FontFace(indic.family, data, { weight: indic.weight });
        document.fonts.add(await face.load());
      }
    }

    pyodide.pyimport("gazette");

    setStatus("Ready.");
  } catch (e) {
    pyodide = null;
    setStatus("Could not load engine. Check your internet connection and reload.", true);
    log(String(e));
  }
  refreshButton();
}

// ---------------------------------------------------------
// FILE LIST
// ---------------------------------------------------------
function addFiles(fileList) {
  const incoming = [...fileList].filter(
    (f) => f.name.toLowerCase().endsWith(".xls") && !f.name.startsWith("~$")
  );
  const known = new Set(files.map((f) => f.name));
  for (const f of incoming) {
    if (!known.has(f.name)) files.push(f);
  }
  files.sort((a, b) => a.name.localeCompare(b.name));
  result.hidden = true;
  renderList();
}

function move(i, delta) {
  const j = i + delta;
  [files[i], files[j]] = [files[j], files[i]];
  result.hidden = true;
  renderList();
}

function removeAt(i) {
  files.splice(i, 1);
  result.hidden = true;
  renderList();
}

function renderList() {
  list.innerHTML = "";

  if (files.length === 0) {
    list.innerHTML = '<li class="empty">No files selected yet.</li>';
  }

  files.forEach((f, i) => {
    const li = document.createElement("li");

    const name = document.createElement("span");
    name.className = "name";
    name.textContent = f.name;
    li.append(name);

    const buttons = [
      ["↑", "Move up", i === 0, () => move(i, -1)],
      ["↓", "Move down", i === files.length - 1, () => move(i, 1)],
      ["✕", "Remove", false, () => removeAt(i)],
    ];
    for (const [label, title, disabled, onClick] of buttons) {
      const b = document.createElement("button");
      b.textContent = label;
      b.title = title;
      b.disabled = disabled;
      b.onclick = onClick;
      li.append(b);
    }

    list.append(li);
  });

  refreshButton();
}

picker.addEventListener("change", () => {
  addFiles(picker.files);
  picker.value = "";
});

drop.addEventListener("dragover", (e) => {
  e.preventDefault();
  drop.classList.add("over");
});
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", (e) => {
  e.preventDefault();
  drop.classList.remove("over");
  addFiles(e.dataTransfer.files);
});

function setLink(a, bytes, name, type) {
  if (a.href) URL.revokeObjectURL(a.href);
  a.href = URL.createObjectURL(new Blob([bytes], { type }));
  a.download = name;
}

// ---------------------------------------------------------
// GENERATE
// ---------------------------------------------------------
runBtn.addEventListener("click", async () => {
  runBtn.disabled = true;
  result.hidden = true;
  logEl.textContent = "";
  setStatus("Processing…");

  let pyFiles = null;
  let out = null;

  try {
    const payload = [];
    for (const f of files) {
      payload.push([f.name, new Uint8Array(await f.arrayBuffer())]);
    }

    const gazette = pyodide.pyimport("gazette");
    pyFiles = pyodide.toPy(payload);
    out = gazette.build(pyFiles);

    const bytes = out.get(0).toJs();
    const pdfBytes = out.get(1).toJs();
    const report = out.get(2).toJs({ dict_converter: Object.fromEntries });
    const lines = out.get(3).toJs();
    lines.forEach(log);

    setLink(download, bytes, gazette.OUTPUT_NAME,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    setLink(downloadPdf, pdfBytes, gazette.PDF_NAME, "application/pdf");

    reportBody.innerHTML = "";
    for (const r of report) {
      const tr = document.createElement("tr");
      if (r.all_zero) tr.className = "zero";
      for (const v of [r.sr, r.sheet, r.page, r.file, r.all_zero ? "All values are 0" : ""]) {
        const td = document.createElement("td");
        td.textContent = v;
        tr.append(td);
      }
      reportBody.append(tr);
    }

    const zeros = report.filter((r) => r.all_zero).length;
    setStatus(
      `Done. ${report.length} sheet(s) merged` +
        (zeros ? `, ${zeros} with all values 0.` : ".")
    );
    result.hidden = false;
  } catch (e) {
    setStatus("Failed. See log for details.", true);
    log(String(e));
  } finally {
    if (pyFiles) pyFiles.destroy();
    if (out) out.destroy();
    refreshButton();
  }
});

loadEngine();
