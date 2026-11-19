// pptx.js.org demos page (uses the PPTXjs v2 UMD bundle: window.PPTXjs)
(function () {
  "use strict";
  var P = window.PPTXjs;
  var $ = function (s) { return document.querySelector(s); };
  if (!P) {
    $("#viewer").textContent = "PPTXjs failed to load.";
    return;
  }

  var source = "Sample_demo1.pptx";   // URL string or File
  var viewer = null;
  var editor = null;
  var mode = "list";

  // ------------------------------------------------------------ hash <-> tab (old demo URLs redirect here)
  var tabByHash = { viewer: "t-viewer", slideshow: "t-slideshow", upload: "t-viewer", editor: "t-editor", convert: "t-convert", code: "t-code", rtl: "t-viewer" };
  var hash = location.hash.slice(1);
  if (hash === "rtl") { source = "../samples/hebrew-rtl.pptx"; $("#sample").value = source; $("#o-textdir").value = "auto"; }
  if (hash === "slideshow") mode = "slideshow";

  // ------------------------------------------------------------ event log
  function log(msg, kind) {
    var li = document.createElement("li");
    li.className = kind || "";
    li.textContent = new Date().toLocaleTimeString() + "  " + msg;
    var list = $("#log");
    list.insertBefore(li, list.firstChild);
    while (list.children.length > 50) list.removeChild(list.lastChild);
  }

  // ------------------------------------------------------------ viewer
  function options() {
    var scale = $("#o-scale").value;
    var theme = $("#o-theme").value;
    return {
      mode: mode,
      scale: scale === "fit" || scale === "contain" ? scale : Number(scale),
      themeProcess: theme === "true" ? true : theme === "false" ? false : "colorsAndImageOnly",
      skipHiddenSlides: $("#o-hidden").checked,
      textDirection: $("#o-textdir").value,
      dir: $("#o-dir").value,
      slideshow: {
        transition: $("#o-transition").value,
        autoplay: Number($("#o-autoplay").value) || false,
        loop: $("#o-loop").checked
      }
    };
  }

  function createViewer() {
    if (viewer) viewer.destroy();
    var host = $("#viewer");
    host.style.height = mode === "slideshow" && $("#o-scale").value === "contain" ? "70vh" : "";
    viewer = new P.PptxViewer(host, options());
    viewer.on("loadstart", function () { log("loading…"); });
    viewer.on("load", function (e) {
      var p = e.presentation;
      log("loaded " + p.slides.length + " slides (" + p.width + "×" + p.height + " px) in " + p.renderTime + " ms");
    });
    viewer.on("warning", function (e) { log(e.warning.message, "warn"); });
    viewer.on("error", function (e) { log(e.error.code + ": " + e.error.message, "error"); });
    viewer.on("slidechange", function (e) { log("slide " + (e.index + 1)); });
    viewer.load(source).catch(function () {});
    if (mode === "slideshow" && Number($("#o-autoplay").value) > 0) {
      viewer.on("load", function () { viewer.play(Number($("#o-autoplay").value)); });
    }
  }

  ["#o-scale", "#o-theme", "#o-hidden", "#o-textdir", "#o-autoplay"].forEach(function (id) {
    $(id).addEventListener("change", createViewer);
  });
  ["#o-transition", "#o-loop"].forEach(function (id) {
    $(id).addEventListener("change", function () { viewer && viewer.setOptions({ slideshow: options().slideshow }); });
  });
  $("#o-dir").addEventListener("change", function () { viewer && viewer.setOptions({ dir: options().dir }); });
  $("#b-prev").addEventListener("click", function () { viewer && viewer.prev(); });
  $("#b-next").addEventListener("click", function () { viewer && viewer.next(); });
  $("#b-fs").addEventListener("click", function () { viewer && viewer.toggleFullscreen(); });

  // ------------------------------------------------------------ tabs
  document.addEventListener("tabselect", function (e) {
    var id = e.detail.id;
    if (id === "t-viewer" || id === "t-slideshow") {
      var wanted = id === "t-slideshow" ? "slideshow" : "list";
      // both tabs share one panel; keep it visible
      $("#p-viewer").hidden = false;
      if (wanted !== mode) {
        mode = wanted;
        if (viewer) viewer.setMode(mode); else createViewer();
      }
    }
    if (id === "t-editor") ensureEditor();
    var name = { "t-viewer": "viewer", "t-slideshow": "slideshow", "t-editor": "editor", "t-convert": "convert", "t-code": "code" }[id];
    if (name && history.replaceState) history.replaceState(null, "", "#" + name);
  });

  // ------------------------------------------------------------ editor (lazy)
  function ensureEditor() {
    if (editor) return;
    editor = new P.PptxEditor("#editor", { fileName: "edited-presentation", textDirection: $("#o-textdir").value });
    editor.on("error", function (e) { log("editor: " + e.error.message, "error"); });
    editor.load(source).catch(function () {});
  }

  // ------------------------------------------------------------ open files
  function open(src, label) {
    source = src;
    $("#current-file").textContent = label || "";
    createViewer();
    if (editor) editor.load(src).catch(function () {});
    $("#c-stats").textContent = "";
  }
  $("#file-input").addEventListener("change", function (e) {
    var f = e.target.files[0];
    if (f) open(f, f.name);
    e.target.value = "";
  });
  $("#sample").addEventListener("change", function (e) {
    if (e.target.value.indexOf("hebrew") !== -1 && $("#o-textdir").value === "file") $("#o-textdir").value = "auto";
    open(e.target.value, "");
  });

  var overlay = $("#drop-overlay");
  var depth = 0;
  document.addEventListener("dragenter", function (e) {
    if (e.dataTransfer && Array.prototype.indexOf.call(e.dataTransfer.types, "Files") !== -1) { depth++; overlay.hidden = false; }
  });
  document.addEventListener("dragleave", function () { if (--depth <= 0) { depth = 0; overlay.hidden = true; } });
  document.addEventListener("dragover", function (e) { e.preventDefault(); });
  document.addEventListener("drop", function (e) {
    e.preventDefault();
    depth = 0;
    overlay.hidden = true;
    var f = e.dataTransfer && e.dataTransfer.files[0];
    if (!f) return;
    if (!/\.pptx$/i.test(f.name)) { log(f.name + " is not a .pptx file", "error"); return; }
    open(f, f.name);
  });

  // ------------------------------------------------------------ convert & export
  function download(blob, name) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }
  function baseName() {
    var n = typeof source === "string" ? source.split("/").pop() : source.name;
    return n.replace(/\.pptx$/i, "");
  }
  $("#c-html").addEventListener("click", function () {
    var t0 = performance.now();
    P.renderToHtml(source, { textDirection: $("#o-textdir").value }).then(function (res) {
      var ms = Math.round(performance.now() - t0);
      $("#c-stats").textContent = res.slides.length + " slides converted in " + ms + " ms.";
      var notes = $("#c-notes");
      notes.textContent = "";
      res.slides.forEach(function (s) {
        var li = document.createElement("li");
        li.textContent = (s.hidden ? "[hidden] " : "") + (s.notes || "(no speaker notes)");
        notes.appendChild(li);
      });
      var page = "<!doctype html><meta charset=\"utf-8\"><title>" + baseName() + "</title><style>" + P.BASE_CSS + "\n" + res.css +
        "\nbody{background:#e9e7e3;margin:0;padding:24px}.pptxjs .slide{margin:0 auto 24px;box-shadow:0 2px 10px #0003}</style>" +
        "<div class=\"pptxjs\">" + res.slides.map(function (s) { return s.html; }).join("\n") + "</div>";
      download(new Blob([page], { type: "text/html" }), baseName() + ".html");
    }).catch(function (e) { $("#c-stats").textContent = e.message; });
  });
  $("#c-odp").addEventListener("click", function () {
    P.loadPptx(source).then(function (pkg) {
      download(P.exportOdp(pkg, { textDirection: $("#o-textdir").value }), baseName() + ".odp");
    }).catch(function (e) { $("#c-stats").textContent = e.message; });
  });
  $("#c-pptx").addEventListener("click", function () {
    P.loadPptx(source).then(function (pkg) { download(P.exportPptx(pkg), baseName() + "-copy.pptx"); })
      .catch(function (e) { $("#c-stats").textContent = e.message; });
  });

  // ------------------------------------------------------------ start
  createViewer();
  var tab = document.getElementById(tabByHash[hash] || "t-viewer");
  if (tab && tab.id !== "t-viewer") window.selectTab(tab, false);
})();
