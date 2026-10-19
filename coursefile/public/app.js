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

  function render(s) {
    $("warnlist").innerHTML = (s.warnings || []).map((w) => `<li>${esc(w)}</li>`).join("");
    $("warn").classList.toggle("show", !!(s.warnings || []).length);
    $("pv").innerHTML = `
      <div class="stats">
        <div class="stat"><b>${s.cos}</b><span>Course outcomes</span></div>
        <div class="stat"><b>${s.students}</b><span>Students</span></div>
        <div class="stat"><b>${s.tools}</b><span>Assessments</span></div>
      </div>
      <dl>
        <dt>Course</dt><dd>${esc(s.course) || "Not found"}</dd>
        <dt>Instructor</dt><dd>${esc(s.faculty)}</dd>
        <dt>Program</dt><dd>${esc(s.program)}</dd>
        <dt>Semester</dt><dd>${esc(s.semester)}, ${esc(s.year)}</dd>
      </dl>
      <h4>Average CO attainment</h4>
      ${s.attainment.map((a) => `<div class="att"><span>${esc(a.co)}</span><i><em data-w="${a.value}"></em></i><span>${a.value}%</span></div>`).join("")}`;
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
})();
