// pptx.js.org: shared site behaviour (theme, menu, tabs, copy buttons, code colouring, docs TOC)
(function () {
  "use strict";
  var root = document.documentElement;

  // ---- theme: system default, user choice remembered
  var stored = null;
  try { stored = localStorage.getItem("theme"); } catch (e) { /* storage blocked */ }
  if (stored === "light" || stored === "dark") root.setAttribute("data-theme", stored);
  function currentTheme() {
    var t = root.getAttribute("data-theme");
    if (t) return t;
    return window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  }
  document.addEventListener("click", function (e) {
    var btn = e.target.closest && e.target.closest("[data-theme-toggle]");
    if (!btn) return;
    var next = currentTheme() === "dark" ? "light" : "dark";
    root.setAttribute("data-theme", next);
    try { localStorage.setItem("theme", next); } catch (err) { /* ignore */ }
    btn.setAttribute("aria-label", next === "dark" ? "Switch to light theme" : "Switch to dark theme");
  });

  // ---- mobile menu
  document.addEventListener("click", function (e) {
    var btn = e.target.closest && e.target.closest("[data-menu-toggle]");
    if (!btn) return;
    var nav = document.getElementById(btn.getAttribute("aria-controls"));
    var open = nav.classList.toggle("open");
    btn.setAttribute("aria-expanded", String(open));
  });

  // ---- tabs (ARIA tablist; keyboard arrows)
  function selectTab(tab, focus) {
    var list = tab.closest('[role="tablist"]');
    var tabs = list.querySelectorAll('[role="tab"]');
    for (var i = 0; i < tabs.length; i++) {
      var on = tabs[i] === tab;
      tabs[i].setAttribute("aria-selected", String(on));
      tabs[i].tabIndex = on ? 0 : -1;
      var panel = document.getElementById(tabs[i].getAttribute("aria-controls"));
      if (panel) panel.hidden = !on;
    }
    if (focus) tab.focus();
    tab.dispatchEvent(new CustomEvent("tabselect", { bubbles: true, detail: { id: tab.id } }));
  }
  window.selectTab = selectTab;
  document.addEventListener("click", function (e) {
    var tab = e.target.closest && e.target.closest('[role="tab"]');
    if (tab) selectTab(tab, false);
  });
  document.addEventListener("keydown", function (e) {
    var tab = e.target.closest && e.target.closest('[role="tab"]');
    if (!tab || (e.key !== "ArrowRight" && e.key !== "ArrowLeft")) return;
    var tabs = Array.prototype.slice.call(tab.closest('[role="tablist"]').querySelectorAll('[role="tab"]'));
    var dir = (e.key === "ArrowRight") !== (getComputedStyle(tab).direction === "rtl") ? 1 : -1;
    selectTab(tabs[(tabs.indexOf(tab) + dir + tabs.length) % tabs.length], true);
    e.preventDefault();
  });

  // ---- minimal syntax colouring for JS/TS/HTML snippets (no dependency)
  function colour(code) {
    var src = code.textContent;
    var out = "";
    var re = /(\/\/[^\n]*|\/\*[\s\S]*?\*\/|<!--[\s\S]*?-->)|("(?:[^"\\\n]|\\.)*"|'(?:[^'\\\n]|\\.)*'|`(?:[^`\\]|\\.)*`)|\b(import|from|export|const|let|var|new|await|async|function|return|if|else|for|of|in|true|false|null|class|extends|type|interface)\b/g;
    var last = 0, m;
    var esc = function (s) { return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); };
    while ((m = re.exec(src))) {
      out += esc(src.slice(last, m.index));
      var cls = m[1] ? "tok-c" : m[2] ? "tok-s" : "tok-k";
      out += '<span class="' + cls + '">' + esc(m[0]) + "</span>";
      last = re.lastIndex;
    }
    code.innerHTML = out + esc(src.slice(last));
  }

  // ---- wrap <pre> in .code with a copy button
  function enhanceCode() {
    var pres = document.querySelectorAll("pre > code");
    for (var i = 0; i < pres.length; i++) {
      var code = pres[i];
      var pre = code.parentElement;
      if (pre.parentElement.classList.contains("code")) continue;
      if (!code.hasAttribute("data-plain")) colour(code);
      var wrap = document.createElement("div");
      wrap.className = "code";
      pre.parentNode.insertBefore(wrap, pre);
      wrap.appendChild(pre);
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "copy";
      btn.textContent = "Copy";
      btn.setAttribute("aria-label", "Copy code");
      wrap.appendChild(btn);
    }
  }
  document.addEventListener("click", function (e) {
    var btn = e.target.closest && e.target.closest(".copy");
    if (!btn) return;
    var text = btn.getAttribute("data-copy") || (btn.parentElement.querySelector("code") || btn.parentElement).textContent.trim();
    var done = function () {
      var old = btn.textContent;
      btn.textContent = "Copied";
      setTimeout(function () { btn.textContent = old; }, 1400);
    };
    if (navigator.clipboard) navigator.clipboard.writeText(text).then(done, function () {});
  });

  // ---- docs: highlight the current section in the TOC
  function trackToc() {
    var links = document.querySelectorAll(".toc a[href^='#']");
    if (!links.length || !("IntersectionObserver" in window)) return;
    var map = {};
    for (var i = 0; i < links.length; i++) map[links[i].getAttribute("href").slice(1)] = links[i];
    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (en) {
        if (!en.isIntersecting) return;
        for (var k in map) map[k].classList.remove("active");
        if (map[en.target.id]) map[en.target.id].classList.add("active");
      });
    }, { rootMargin: "-80px 0px -70% 0px" });
    for (var id in map) {
      var el = document.getElementById(id);
      if (el) io.observe(el);
    }
  }

  // ---- current year in footer
  function year() {
    var y = document.querySelectorAll("[data-year]");
    for (var i = 0; i < y.length; i++) y[i].textContent = String(new Date().getFullYear());
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
  function init() { enhanceCode(); trackToc(); year(); }
})();
