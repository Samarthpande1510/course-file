(() => {
  const $ = (id) => document.getElementById(id);
  const MAX = 15 * 1024 * 1024;
  let file = null, data = null;
  const go = $("go"), drop = $("drop");

  const size = (n) => (n > 1048576 ? (n / 1048576).toFixed(1) + " MB" : Math.ceil(n / 1024) + " KB");
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const showErr = (m) => { $("errmsg").textContent = " " + m; $("err").classList.add("show"); };
  const hideErr = () => { $("err").classList.remove("show"); $("warn").classList.remove("show"); };

  function reset() {
    file = data = null;
    $("file").value = "";
    $("chip").classList.remove("show");
    $("done").classList.remove("show");
    go.disabled = true;
    $("status").textContent = "";
    $("pv").innerHTML = '<div class="empty"><i></i>Upload a workbook to see its details here.</div>';
    hideErr();
  }

  const MARK = { ok: "✓", placeholder: "!", missing: "×" };
  const NOTE = { placeholder: "Template text, edit it", missing: "Not found in the sheet" };

  function render(s) {
    $("warnlist").innerHTML = (s.warnings || []).map((w) => `<li>${esc(w)}</li>`).join("");
    $("warn").classList.toggle("show", !!(s.warnings || []).length);

    const meta = [s.program, s.semester && "Semester " + s.semester, s.year].filter(Boolean).map(esc).join(", ");
    const checks = s.fields.map((f) => `
      <li><span class="mk ${f.status}" aria-hidden="true">${MARK[f.status]}</span>
        <span class="k">${esc(f.label)}</span>
        <span class="v">${f.value ? esc(f.value) : ""}${NOTE[f.status] ? `<small class="${f.status}">${NOTE[f.status]}</small>` : ""}</span></li>`).join("");

    const outcomes = s.outcomes.length ? `
      <h4>Course outcomes</h4>
      <ul class="outc">${s.outcomes.map((o) => `
        <li><div class="top"><span>${esc(o.id)}</span><span>${o.value === null ? "no data" : o.value + "%"}</span></div>
          <p title="${esc(o.text)}">${esc(o.text)}</p>
          <i><em data-w="${o.value === null ? 0 : o.value}"></em></i></li>`).join("")}</ul>` : "";

    const avgCell = (a) => a.avg !== null ? `${a.avg}%` : a.mean !== null ? `${a.mean} marks` : `<span class="na">n/a</span>`;
    const assess = s.assessments.length ? `
      <h4>Assessments found</h4>
      <table class="asm"><thead><tr><th scope="col">Name</th><th scope="col">Students</th><th scope="col">Average</th></tr></thead>
        <tbody>${s.assessments.map((a) => `<tr><td>${esc(a.name)}</td><td>${a.students}</td><td>${avgCell(a)}</td></tr>`).join("")}</tbody></table>` : "";

    const hours = s.hours.length ? `<p class="hours">${s.hours.map(([k, v]) => `${esc(k)} ${esc(v)}`).join(", ")} hours per week</p>` : "";

    $("pv").innerHTML = `
      <h3 class="sum-title">${esc(s.course) || "Course name not found"}</h3>
      <p class="sum-meta">${meta}</p>
      <div class="stats">
        <div class="stat"><b>${s.students}</b><span>Students</span></div>
        <div class="stat"><b>${s.outcomes.length}</b><span>Outcomes</span></div>
        <div class="stat"><b>${s.assessments.length}</b><span>Assessments</span></div>
        <div class="stat"><b>${s.sheets}</b><span>Sheets</span></div>
      </div>
      <h4>What goes on the cover page</h4>
      <ul class="checks">${checks}</ul>
      ${outcomes}${assess}${hours}`;
    requestAnimationFrame(() => requestAnimationFrame(() =>
      document.querySelectorAll("#pv em").forEach((e) => (e.style.width = Math.min(100, e.dataset.w) + "%"))));
  }

  async function take(f) {
    hideErr();
    $("done").classList.remove("show");
    if (!f) return;
    if (!/\.xlsx$/i.test(f.name)) return showErr("Please choose an Excel .xlsx file.");
    if (f.size > MAX) return showErr("That file is larger than 15 MB.");
    file = f; data = null;
    $("fname").textContent = f.name;
    $("fsize").textContent = size(f.size);
    $("chip").classList.add("show");
    go.disabled = true;
    $("status").textContent = "Reading workbook...";
    try {
      data = await Engine.readWorkbook(await f.arrayBuffer());
      render(Engine.summarize(data));
      go.disabled = false;
      $("status").textContent = "Ready to generate.";
    } catch (e) {
      file = data = null;
      $("chip").classList.remove("show");
      $("status").textContent = "";
      showErr(e instanceof Engine.WorkbookError ? e.message : "Something went wrong while reading the workbook.");
    }
  }

  $("file").addEventListener("change", (e) => take(e.target.files[0]));
  $("clear").addEventListener("click", reset);
  ["dragenter", "dragover"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", (e) => take(e.dataTransfer.files[0]));

  go.addEventListener("click", async () => {
    if (!data) return;
    $("err").classList.remove("show");
    $("done").classList.remove("show");
    go.disabled = true;
    go.classList.add("spin");
    $("status").textContent = "Building your course file...";
    try {
      const tpl = await fetch("template.docx");
      if (!tpl.ok) throw new Error("The course file template could not be loaded.");
      const blob = await Engine.buildCourseFile(data, {
        course_code: $("code").value.trim(),
        program_short: $("prog").value.trim(),
        department: $("dept").value.trim(),
      }, await tpl.arrayBuffer());
      const d = data.details;
      const name = `Course File - ${d["course name"] || "Course"} ${d["academic year"] || ""}.docx`.replace(/[\\/:*?"<>|]/g, "-");
      const url = URL.createObjectURL(blob);
      const a = Object.assign(document.createElement("a"), { href: url, download: name });
      document.body.append(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 2000);
      $("done").classList.add("show");
      $("status").textContent = "";
    } catch (e) {
      showErr(e.message || "Generation failed.");
      $("status").textContent = "";
    } finally {
      go.disabled = false;
      go.classList.remove("spin");
    }
  });

  // Highlight the nav link for the section in view.
  const links = [...document.querySelectorAll('.nav ul a')];
  const byId = Object.fromEntries(links.map((l) => [l.getAttribute("href").slice(1), l]));
  if ("IntersectionObserver" in window) {
    const io = new IntersectionObserver((entries) => {
      entries.forEach((en) => {
        if (!en.isIntersecting) return;
        links.forEach((l) => l.removeAttribute("aria-current"));
        byId[en.target.id]?.setAttribute("aria-current", "true");
      });
    }, { rootMargin: "-35% 0px -55% 0px" });
    ["build", "requirements", "privacy"].forEach((id) => { const el = document.getElementById(id); if (el) io.observe(el); });
  }
})();
