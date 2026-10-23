/* Excel attainment workbook -> course file (.docx). Runs entirely in the browser.
   Uses globals: ExcelJS, JSZip, DOMParser, XMLSerializer (injected in tests). */
(function (root) {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const FONT = "Book Antiqua";
  const TEXT_WIDTH = 9800; // dxa, usable page width
  const FULL_SECTIONS = false; // true appends the body sections after the index

  class WorkbookError extends Error {}

  // ------------------------------------------------------------ reading
  const s = (v) => (v === null || v === undefined ? "" : String(v).replace(/ /g, " ").trim());
  const num = (v) => (typeof v === "number" && isFinite(v) ? v : null);
  const g = (n) => String(parseFloat(Number(n).toFixed(6)));
  const pct = (v) => (num(v) === null ? "" : num(v).toFixed(1) + "%");

  function norm(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return v;
    if (typeof v === "object") {
      if ("result" in v) return norm(v.result);
      if (v.richText) return v.richText.map((t) => t.text).join("");
      if (v.text !== undefined) return v.text;
      return null;
    }
    return v;
  }

  function grid(ws) {
    const out = [];
    const cols = Math.max(ws.columnCount, 1);
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const arr = [];
      for (let c = 1; c <= cols; c++) {
        const cell = row.getCell(c);
        arr.push(cell.isMerged && cell.master && cell.master.address !== cell.address ? null : norm(cell.value));
      }
      out.push(arr);
    }
    return out;
  }

  // Map whatever label the sheet uses ("Program:", "Programme Name", "Faculty"...) onto our canonical keys.
  const CANON = [
    [/program/, "program name"], [/course\s*name|^course$|subject/, "course name"], [/sem/, "semester"],
    [/year/, "academic year"], [/instructor|faculty|teacher|professor/, "course instructor"],
    [/theory/, "theory hours per week"], [/lab|practical/, "lab hours per week"], [/tutorial/, "tutorial hours per week"],
  ];

  async function readWorkbook(buffer) {
    const wb = new ExcelJS.Workbook();
    try {
      await wb.xlsx.load(buffer);
    } catch (e) {
      throw new WorkbookError("This does not look like an Excel .xlsx workbook. It may be a different file type renamed to .xlsx, damaged, or password protected.");
    }
    const names = wb.worksheets.map((w) => w.name);
    if (!names.length) throw new WorkbookError("This workbook has no sheets.");
    const detailsName = names.find((n) => n.trim().toLowerCase() === "course details");
    if (!detailsName) {
      throw new WorkbookError('This does not look like the attainment workbook: there is no sheet named "Course Details". Sheets found: ' + names.slice(0, 6).join(", ") + (names.length > 6 ? ", ..." : "") + ".");
    }
    const warnings = [];
    // Teachers rename sheets ("CO Attainment", "Course-Exit-Survey"...), so match by meaning, not exact name.
    const find = (re) => names.find((n) => re.test(n.trim()));
    const alias = {
      "COs and Mapping": find(/co.*mapping/i),
      "Assessment Tools(List)": find(/assessment\s*tools/i),
      "Overall CO attainment": find(/overall\s*co|^co\s*-?\s*attainment/i),
      "Course End Survey Results": find(/survey|course.?(end|exit)|exit/i),
      "SO Attainment": find(/^so\s*-?\s*attainment/i),
    };
    const actual = (n) => alias[n] || n;
    const G = (n) => grid(wb.getWorksheet(actual(n)));
    const has = (n) => names.includes(actual(n));
    // Run one section; if the sheet is absent or malformed, carry on (and note it when the body sections are on).
    const part = (sheet, fn, fallback) => {
      if (!has(sheet)) { if (FULL_SECTIONS) warnings.push(`Sheet "${sheet}" was not found, so it was skipped.`); return fallback; }
      try { return fn(); } catch (e) { if (FULL_SECTIONS) warnings.push(`Sheet "${sheet}" could not be read fully and was skipped.`); return fallback; }
    };

    const details = {};
    for (const r of G(detailsName)) {
      const val = r.slice(1).find((x) => s(x) !== "");
      if (!s(r[0]) || val === undefined) continue;
      const k = s(r[0]).replace(/:+$/, "").toLowerCase();
      const canon = (CANON.find(([re]) => re.test(k)) || [null, k])[1];
      if (!(canon in details)) details[canon] = s(val);
    }
    const missing = ["course name", "course instructor", "semester", "academic year", "program name"].filter((k) => !details[k]);
    const looksUnfilled = (v) => /^20x+/i.test(v) || /^abcd\b/i.test(v) || /^(xxx+|tbd|na)$/i.test(v);
    const unfilled = [["academic year", "academic year"], ["course instructor", "instructor name"], ["course name", "course name"]].filter(([k]) => details[k] && looksUnfilled(details[k])).map(([, l]) => l);
    if (unfilled.length) warnings.push("The Course Details sheet still has template placeholder text for: " + unfilled.join(", ") + ". Fix it in Excel, or edit it in the Word file.");
    if (missing.length) warnings.push("These details were not found on the Course Details sheet and will be blank: " + missing.join(", ") + ".");

    let cos = [], sos = [], mapping = {}, soHead = [];
    part("COs and Mapping", () => {
      for (const row of G("COs and Mapping")) {
        const a = s(row[0]), b = s(row[1]);
        if (/^CO-?\d+$/i.test(a)) cos.push([a.toUpperCase(), b]);
        else if (/^SO-?\d+$/i.test(a)) sos.push([a.toUpperCase(), b]);
        const d = s(row[3]);
        if (d === "Mapping") soHead = row.slice(4).map(s).filter((h) => h.startsWith("SO"));
        else if (/^CO\d+$/.test(d) && soHead.length) mapping[d] = row.slice(4, 4 + soHead.length).map((c) => (s(c) ? "√" : ""));
      }
    });

    const tools = part("Assessment Tools(List)", () => {
      const out = [];
      G("Assessment Tools(List)").slice(1).forEach((r) => { if (r[1] && num(r[2]) !== null) out.push([s(r[1]), g(r[2])]); });
      return out;
    }, []);

    let blocks = [], average = null, remarks = { co: "", pedagogy: "" };
    part("Overall CO attainment", () => {
      const ov = G("Overall CO attainment");
      ov.forEach((row, ri) => {
        row.forEach((v, ci) => {
          if (ci + 1 < row.length && s(row[ci + 1]).startsWith("CO") && s(v) && !s(v).startsWith("CO") && !s(v).toLowerCase().startsWith("average")) {
            const heads = [];
            let cj = ci + 1;
            while (cj < row.length && s(row[cj]).startsWith("CO")) { heads.push([cj, s(row[cj])]); cj++; }
            const rows = [];
            for (let rr = ri + 1; rr < Math.min(ri + 4, ov.length); rr++) {
              const lab = s(ov[rr][ci]);
              if (!lab) break;
              rows.push([lab, heads.map(([c]) => ov[rr][c])]);
            }
            if (rows.length) blocks.push([s(v), heads.map(([, h]) => h), rows]);
          }
        });
      });
      ov.forEach((row, ri) => row.forEach((v, ci) => {
        if (s(v).toLowerCase().startsWith("average attainment") && ri + 1 < ov.length) {
          const heads = [];
          for (let cj = ci + 1; cj < row.length; cj++) if (s(row[cj]).startsWith("CO")) heads.push([cj, s(row[cj])]);
          average = [heads.map(([, h]) => h), heads.map(([c]) => ov[ri + 1][c])];
        }
      }));
      ov.forEach((row) => row.forEach((v, ci) => {
        const t = s(v).toLowerCase();
        const next = () => s(row.slice(ci + 1).find((x) => s(x)) ?? "");
        if (t.startsWith("remarks for high")) remarks.co = next();
        else if (t.startsWith("remarks for changes")) remarks.pedagogy = next();
      }));
    });

    const survey = part("Course End Survey Results", () => {
      const sv = G("Course End Survey Results");
      const qs = {};
      sv.forEach((r) => { if (/^Q\d+$/.test(s(r[7]))) qs[s(r[7])] = s(r[8]); });
      const rows = [];
      let responded = "", total = "";
      sv.slice(10).forEach((r) => {
        if (/^Q\d+$/.test(s(r[0])) && num(r[1]) !== null) {
          const counts = r.slice(1, 5).map((x) => Math.trunc(num(x) || 0));
          const n = counts.reduce((a, b) => a + b, 0);
          const score = n ? counts.reduce((a, c, i) => a + c * (i + 1), 0) / n : 0;
          rows.push([s(r[0]), qs[s(r[0])] || "", ...counts, s(r[5]), score.toFixed(2)]);
        } else if (s(r[0]).startsWith("Total Number of students who")) responded = s(r[5]);
        else if (s(r[0]).startsWith("Total Number of students in")) total = s(r[5]);
      });
      return { rows, responded, total };
    }, null);

    const soAtt = part("SO Attainment", () => {
      let head = [];
      const rows = [];
      G("SO Attainment").forEach((r) => {
        if (s(r[3]) === "Component") head = r.slice(3, 12).map(s);
        else if (head.length && s(r[3])) rows.push([s(r[3]), s(r[4]), ...r.slice(5, 12).map(pct)]);
      });
      return head.length && rows.length ? { head, rows } : null;
    }, null);

    // Overall attainment: if the sheet is laid out differently, look for a header row of CO1..COn and an "overall" row.
    if (!average && has("Overall CO attainment")) {
      try {
        const gr = G("Overall CO attainment");
        const coRe = /^c[o0]\s*-?\s*(\d+)$/i;
        const hr = gr.findIndex((r) => r.filter((c) => coRe.test(s(c))).length >= 2);
        if (hr >= 0) {
          const cols = gr[hr].map((c, i) => [i, s(c).match(coRe)]).filter(([, m]) => m).map(([i, m]) => [i, "CO" + m[1]]);
          const row = gr.slice(hr + 1).find((r) => r.slice(0, 3).some((c) => /^overall+\b/i.test(s(c))));
          if (row) average = [cols.map(([, h]) => h), cols.map(([i]) => row[i])];
        }
      } catch (e) { /* leave average empty */ }
    }
    if (!average && blocks.length) {
      // average the "Attainment" row of every block, per CO
      const acc = {};
      for (const [, heads, rows] of blocks) {
        const r = rows.find(([lab]) => /ttainment/i.test(lab));
        if (r) heads.forEach((h, i) => { if (num(r[1][i]) !== null) (acc[h] = acc[h] || []).push(num(r[1][i])); });
      }
      const hs = Object.keys(acc).sort();
      if (hs.length) average = [hs, hs.map((h) => acc[h].reduce((a, b) => a + b, 0) / acc[h].length)];
    }

    // Per-student marks: any sheet that has a "Roll No" header row counts as an assessment, whatever it is called.
    const known = new Set(Object.values(alias).filter(Boolean).concat([detailsName]));
    const short = [[/^class\s*test\s*-?\s*1/i, "CT-1"], [/^class\s*test\s*-?\s*2/i, "CT-2"], [/^lab\s*exp/i, "Lab"], [/^case\s*study/i, "Case"],
      [/^present/i, "Pres."], [/^class\s*par/i, "Part."], [/^assign/i, "Assign."], [/^quiz/i, "Quiz"], [/^viva/i, "Viva"], [/^mini\s*project/i, "Mini Proj."], [/^project/i, "Project"]];
    const icaCols = [], ica = {}, assessments = [];
    const rollHead = /^roll\s*\.?\s*no\.?$/i;
    const notStudent = /max|total|attain|average|appeared|^co\b|marks|out of/i;
    for (const ws of wb.worksheets) {
      if (known.has(ws.name) || /attain|survey|exit|mapping|details$/i.test(ws.name.trim()) && !/\(details & grad/i.test(ws.name)) continue;
      try {
        const gr = grid(ws);
        const hdr = gr.slice(0, 60).findIndex((r) => rollHead.test(s(r[0])));
        if (hdr < 0) continue;
        const head = gr[hdr];
        const use = head.map((h, i) => i).filter((i) => i >= 2 && s(head[i]) !== "" && s(head[i]) !== "Total");
        const mx = gr[hdr + 1] || [];
        const maxTotal = use.reduce((a, i) => a + (num(mx[i]) || 0), 0);
        const nm = ws.name.replace(/\(.*$/, "").replace(/\s*details\s*&\s*grad.*$/i, "").replace(/^class\s*par\w*$/i, "Class Participation").trim();
        const label = (short.find(([re]) => re.test(nm)) || [null, nm.slice(0, 12)])[1];
        const col = icaCols.length;
        const rows = gr.slice(hdr + 2).filter((row) => s(row[0]) && !notStudent.test(s(row[0])) && (s(row[1]) || use.some((i) => num(row[i]) !== null)));
        if (!rows.length) continue;
        // Many teachers only fill the Total column. Use it when it is really filled in; otherwise add up the question columns.
        const tc = head.findIndex((h, i) => i >= 2 && /^total$/i.test(s(h)));
        const fromTotal = tc >= 0 && rows.filter((r) => num(r[tc]) !== null).length >= rows.length / 2;
        const maxT = fromTotal ? num(mx[tc]) || 0 : maxTotal;
        icaCols.push(maxT ? `${label} (/${g(maxT)})` : label);
        const totals = [];
        for (const row of rows) {
          const roll = s(row[0]);
          const tot = fromTotal ? num(row[tc]) || 0 : use.reduce((a, i) => a + (num(row[i]) || 0), 0);
          totals.push(tot);
          (ica[roll] = ica[roll] || { name: s(row[1]), m: {} }).m[col] = tot;
        }
        // average score as % of the maximum; if the layout fooled the sum, show nothing rather than a wrong number
        const mean = totals.reduce((a, b) => a + b, 0) / totals.length;
        let avg = maxT ? (mean / maxT) * 100 : null;
        if (avg !== null && (avg > 100 || (!fromTotal && avg < 1))) avg = null;
        assessments.push({ name: nm || ws.name, students: rows.length, max: maxT || null, avg, mean: avg === null && (fromTotal || mean > 0) ? mean : null });
      } catch (e) {
        if (FULL_SECTIONS) warnings.push(`Sheet "${ws.name}" could not be read fully and was skipped.`);
      }
    }
    return { details, cos, sos, soHead, mapping, tools, blocks, average, remarks, survey, soAtt, icaCols, ica, assessments, sheetNames: names, warnings };
  }

  const isPlaceholder = (v) => /^20x+/i.test(v) || /^abcd\b/i.test(v) || /^(xxx+|tbd|na)$/i.test(v);

  function summarize(d) {
    const x = d.details;
    const digits = (t) => (String(t).match(/\d+/) || [""])[0];
    const att = {};
    if (d.average) d.average[0].forEach((h, i) => { if (num(d.average[1][i]) !== null) att[digits(h)] = num(d.average[1][i]); });
    const field = (label, key) => {
      const v = x[key] || "";
      return { label, value: v, status: !v ? "missing" : isPlaceholder(v) ? "placeholder" : "ok" };
    };
    return {
      course: x["course name"] || "", faculty: x["course instructor"] || "", semester: x["semester"] || "",
      year: x["academic year"] || "", program: x["program name"] || "",
      warnings: d.warnings || [],
      fields: [field("Course name", "course name"), field("Program", "program name"), field("Semester", "semester"),
        field("Academic year", "academic year"), field("Instructor", "course instructor")],
      sheets: d.sheetNames.length,
      students: Object.keys(d.ica).length,
      outcomes: d.cos.map(([id, text]) => ({ id, text, value: digits(id) in att ? Math.round(att[digits(id)] * 10) / 10 : null })),
      assessments: d.assessments.map((a) => ({ ...a, avg: a.avg === null ? null : Math.round(a.avg * 10) / 10, mean: a.mean === null || a.mean === undefined ? null : Math.round(a.mean * 10) / 10 })),
      hours: [["Theory", x["theory hours per week"]], ["Lab", x["lab hours per week"]], ["Tutorial", x["tutorial hours per week"]]].filter(([, v]) => v),
    };
  }

  // ------------------------------------------------------------ writing: DOM helpers
  const kids = (el, name) => Array.from(el.childNodes).filter((n) => n.nodeType === 1 && n.localName === name);
  const esc = (t) => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  function runText(r) {
    let t = "";
    for (const n of Array.from(r.childNodes)) {
      if (n.nodeType !== 1) continue;
      if (n.localName === "t") t += n.textContent;
      else if (n.localName === "tab") t += "\t";
    }
    return t;
  }
  function setRunText(r, text) {
    const doc = r.ownerDocument;
    for (const n of Array.from(r.childNodes)) if (!(n.nodeType === 1 && n.localName === "rPr")) r.removeChild(n);
    text.split("\t").forEach((part, i) => {
      if (i > 0) r.appendChild(doc.createElementNS(W, "w:tab"));
      if (part) {
        const t = doc.createElementNS(W, "w:t");
        t.setAttribute("xml:space", "preserve");
        t.appendChild(doc.createTextNode(part));
        r.appendChild(t);
      }
    });
  }
  const paraText = (p) => kids(p, "r").map(runText).join("");

  function rewriteCoverLine(p, value) {
    const runs = kids(p, "r");
    const idx = runs.findIndex((r) => runText(r).includes(":"));
    if (idx < 0) return;
    const t = runText(runs[idx]);
    const head = t.slice(0, t.indexOf(":") + 1) + " ";
    setRunText(runs[idx], head);
    runs.slice(idx + 1).forEach((r) => setRunText(r, ""));
    if (idx + 1 < runs.length) setRunText(runs[idx + 1], value);
    else setRunText(runs[idx], head + value);
  }

  function setCell(tc, text, rowFallback) {
    const doc = tc.ownerDocument;
    const ps = kids(tc, "p");
    ps.slice(1).forEach((p) => kids(p, "r").forEach((r) => setRunText(r, "")));
    const runs = kids(ps[0], "r");
    if (runs.length) {
      setRunText(runs[0], text);
      runs.slice(1).forEach((r) => setRunText(r, ""));
    } else {
      const r = doc.createElementNS(W, "w:r");
      if (rowFallback) r.appendChild(rowFallback.cloneNode(true));
      ps[0].appendChild(r);
      setRunText(r, text);
    }
  }

  // ------------------------------------------------------------ writing: XML builders
  function rPr(size, bold, italic) {
    return `<w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:eastAsia="${FONT}" w:cs="${FONT}"/>${bold ? "<w:b/>" : ""}${italic ? "<w:i/>" : ""}<w:sz w:val="${size * 2}"/><w:szCs w:val="${size * 2}"/></w:rPr>`;
  }
  function para(text = "", o = {}) {
    const { size = 11, bold = false, italic = false, after = 6, keep = false, center = false } = o;
    return `<w:p><w:pPr>${keep ? "<w:keepNext/>" : ""}<w:spacing w:before="0" w:after="${after * 20}"/>${center ? '<w:jc w:val="center"/>' : ""}</w:pPr>` +
      (text ? `<w:r>${rPr(size, bold, italic)}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>` : "") + "</w:p>";
  }
  const pageBreak = () => '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  const heading = (num_, text, level = 1) => para(`${num_}  ${text}`, { size: { 1: 14, 2: 12, 3: 11 }[level], bold: true, after: 6, keep: true });
  const placeholder = (t) => para(`[ ${t} ]`, { size: 10, italic: true, after: 10 });

  function table(head, rows, o = {}) {
    const size = o.size || 9;
    const n = head.length;
    const raw = o.widths || Array(n).fill(1);
    const sum = raw.reduce((a, b) => a + b, 0);
    const w = raw.map((x) => Math.round((TEXT_WIDTH * x) / sum));
    const cf = o.centerFrom ?? 999;
    const cell = (text, i, isHead) =>
      `<w:tc><w:tcPr><w:tcW w:w="${w[i]}" w:type="dxa"/>${isHead ? '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>' : ""}<w:vAlign w:val="center"/></w:tcPr>` +
      `<w:p><w:pPr><w:spacing w:before="30" w:after="30"/>${i >= cf ? '<w:jc w:val="center"/>' : ""}</w:pPr>` +
      (text !== "" && text != null ? `<w:r>${rPr(size, isHead, false)}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>` : "") + "</w:p></w:tc>";
    const b = '<w:{k} w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
    const borders = ["top", "left", "bottom", "right", "insideH", "insideV"].map((k) => b.replace("{k}", k)).join("");
    return `<w:tbl><w:tblPr><w:tblW w:w="${TEXT_WIDTH}" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="70" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
      `<w:tblGrid>${w.map((x) => `<w:gridCol w:w="${x}"/>`).join("")}</w:tblGrid>` +
      `<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>${head.map((h, i) => cell(h, i, true)).join("")}</w:tr>` +
      rows.map((r) => `<w:tr><w:trPr><w:cantSplit/></w:trPr>${r.map((v, i) => cell(v, i, false)).join("")}</w:tr>`).join("") +
      "</w:tbl>" + para("", { after: 4 });
  }

  // ------------------------------------------------------------ build the course file
  async function buildCourseFile(data, extra, templateBuffer) {
    extra = extra || {};
    const d = data.details;
    const zip = await JSZip.loadAsync(templateBuffer);
    const xml = await zip.file("word/document.xml").async("string");
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const body = Array.from(doc.documentElement.childNodes).find((n) => n.nodeType === 1 && n.localName === "body");
    const paras = kids(body, "p");
    const byStart = {};
    paras.forEach((p) => {
      const t = paraText(p);
      if (t.trim()) byStart[t.split("\t")[0].trim().split(":")[0].trim()] = p;
    });

    const year = d["academic year"] || "", program = d["program name"] || "", course = d["course name"] || "";
    const sem = d["semester"] || "", faculty = d["course instructor"] || "";
    const code = extra.course_code || "", short = extra.program_short || "";

    if (extra.department) {
      const dp = paras.find((p) => paraText(p).trim().toUpperCase().startsWith("DEPARTMENT"));
      if (dp) { const rs = kids(dp, "r"); setRunText(rs[0], extra.department.toUpperCase() + " "); rs.slice(1).forEach((r) => setRunText(r, "")); }
    }
    for (const k of ["Course Name", "Semester", "Program", "Academic Year", "Faculty Name"]) {
      if (!byStart[k]) throw new WorkbookError("The course file template is missing its \"" + k + "\" line, so it cannot be filled in.");
    }
    rewriteCoverLine(byStart["Course Name"], course);
    rewriteCoverLine(byStart["Semester"], sem);
    rewriteCoverLine(byStart["Program"], [program, short && `(${short})`].filter(Boolean).join(" "));
    rewriteCoverLine(byStart["Academic Year"], year);
    rewriteCoverLine(byStart["Faculty Name"], faculty);

    const t0 = kids(body, "tbl")[0];
    const trs = kids(t0, "tr");
    const cell = (r, c) => kids(trs[r], "tc")[c];
    const fb = (r) => { const run = kids(kids(cell(r, 0), "p")[0], "r")[0]; return run && kids(run, "rPr")[0]; };
    setCell(cell(0, 1), code, fb(0)); setCell(cell(0, 3), program, fb(0));
    setCell(cell(1, 1), faculty, fb(1)); setCell(cell(1, 3), short, fb(1));
    setCell(cell(2, 1), course, fb(2)); setCell(cell(2, 3), `${year}/${sem}`, fb(2));

    // drop empty trailing paragraphs after the TOC so our content starts cleanly
    const sect = kids(body, "sectPr")[0];
    const all = Array.from(body.childNodes).filter((n) => n.nodeType === 1);
    for (let i = all.length - 1; i >= 0; i--) {
      const el = all[i];
      if (el === sect) continue;
      if (el.localName === "p" && !paraText(el).trim() && !/<w:drawing|<w:pict|w:br/.test(new XMLSerializer().serializeToString(el))) body.removeChild(el);
      else break;
    }

    const o = [];
    const add = (x) => o.push(x);
    add(pageBreak());
    add(heading("1", "Academic Calendar")); add(placeholder("Attach or paste the institute's academic calendar here."));
    add(heading("2", "Course Policy"));
    add(table(["Course", "Value"], [["Course name", course], ["Semester", sem], ["Instructor", faculty],
      ["Theory hours per week", d["theory hours per week"] || ""], ["Lab hours per week", d["lab hours per week"] || ""], ["Tutorial hours per week", d["tutorial hours per week"] || ""]],
      { widths: [200, 250], size: 10 }));
    if (data.tools.length) {
      add(para("Assessment scheme", { bold: true, after: 4 }));
      add(table(["Assessment tool", "Marks"], [...data.tools, ["Total", g(data.tools.reduce((a, t) => a + parseFloat(t[1]), 0))]], { widths: [250, 100], size: 10 }));
    }
    add(para("Course outcomes", { bold: true, after: 4 }));
    add(table(["CO", "Statement"], data.cos, { widths: [50, 400], size: 10 }));
    if (Object.keys(data.mapping).length) {
      add(para("CO – SO mapping", { bold: true, after: 4 }));
      add(table(["CO", ...data.soHead], Object.entries(data.mapping).map(([co, m]) => [co, ...m]), { size: 10, centerFrom: 1 }));
      add(para("Student outcomes", { bold: true, after: 4 }));
      add(table(["SO", "Statement"], data.sos, { widths: [50, 400], size: 9 }));
    }

    add(pageBreak());
    add(heading("3", "Assessment and Evaluation"));
    add(heading("3.1", "Direct Assessment Tools", 2));
    add(heading("3.1.1", "CO attainment – Lab Experiments / Tutorials / Assignments / Presentations / Case studies", 3));
    const isMid = (b) => /^M\d/.test(b[0]);
    const blockTable = ([title, heads, rows]) => {
      add(para(title, { bold: true, size: 10, after: 2, keep: true }));
      add(table(["", ...heads], rows.map(([lab, vals]) => [lab, ...vals.map((v) => (/ttainment/.test(lab) ? pct(v) : num(v) !== null ? g(v) : ""))]), { widths: [3, ...heads.map(() => 1)], centerFrom: 1 }));
    };
    data.blocks.filter((b) => !isMid(b)).forEach(blockTable);
    add(heading("3.1.2", "CO attainment – Mid Term Test I and II, with Question Papers and marks", 3));
    data.blocks.filter(isMid).forEach(blockTable);
    add(placeholder("Attach Mid Term Test I and II question papers here."));
    if (data.average) {
      add(para("Average CO attainment (all direct tools)", { bold: true, size: 10, after: 2, keep: true }));
      add(table(data.average[0], [data.average[1].map(pct)], { size: 10, centerFrom: 0 }));
    }
    add(heading("3.1.3", "Sample answer sheets of Mid Term Test I and II (Good, Average, Below Average – 2 each)", 3));
    add(placeholder("Attach scanned sample answer sheets here."));
    add(heading("3.1.4", "CO attainment – Term End Examination", 3));
    add(placeholder("Term End Examination CO attainment is not part of the workbook; add it here."));
    add(heading("3.1.5", "Term End Examination Question Papers", 3));
    add(placeholder("Attach Term End Examination question papers here."));
    add(heading("3.2", "Indirect Assessment Tools", 2));
    add(heading("3.2.1", "CO attainment – Course End Survey", 3));
    if (data.survey && data.survey.rows.length) {
      add(para("Rating: 1 = Do not agree; 2 = Partially agree; 3 = Agree; 4 = Strongly agree. Entries are number of students.", { size: 9, italic: true, after: 4 }));
      add(table(["Q", "Question", "1", "2", "3", "4", "CO", "Avg."], data.survey.rows, { widths: [28, 250, 24, 24, 24, 24, 34, 36], centerFrom: 2 }));
      if (data.survey.responded) add(para(`Responses received: ${data.survey.responded} of ${data.survey.total} students.`, { size: 10 }));
    } else add(placeholder("No survey sheet found in the workbook."));

    add(pageBreak());
    add(heading("4", "Record of Internal Continuous Assessment (ICA) marks"));
    const rolls = Object.keys(data.ica).sort();
    if (rolls.length) {
      add(table(["Roll no", "Name", ...data.icaCols], rolls.map((r) => [r, data.ica[r].name, ...data.icaCols.map((_, i) => (i in data.ica[r].m ? g(data.ica[r].m[i]) : ""))]),
        { size: 8, widths: [55, 150, ...data.icaCols.map(() => 50)], centerFrom: 2 }));
    } else add(placeholder("No student marks found."));

    add(pageBreak());
    add(heading("5", "Corrective actions for Continuous Improvement (CI)"));
    add(heading("5.1", "Faculty Course Assessment Report (FCAR)", 2));
    add(para("Remarks for high / low CO attainment", { bold: true, size: 10, after: 2 }));
    add(data.remarks.co ? para(data.remarks.co, { size: 10 }) : placeholder("Enter remarks"));
    add(para("Remarks for changes required in syllabus / pedagogy as per CO attainment", { bold: true, size: 10, after: 2 }));
    add(data.remarks.pedagogy ? para(data.remarks.pedagogy, { size: 10 }) : placeholder("Enter remarks"));
    if (data.soAtt) {
      add(para("SO attainment", { bold: true, size: 10, after: 2, keep: true }));
      add(table(data.soAtt.head, data.soAtt.rows, { size: 8, widths: [90, 60, ...Array(7).fill(45)], centerFrom: 2 }));
    }
    add(heading("5.2", "Suggestion by Module Coordinators and Program Coordinator", 2));
    add(placeholder("Add suggestions here."));
    add(heading("6", "Study material (links or QR code can be added)"));
    add(placeholder("Add study-material links or QR codes here."));

    const frag = new DOMParser().parseFromString(`<w:body xmlns:w="${W}">${o.join("")}</w:body>`, "text/xml");
    if (FULL_SECTIONS) for (const n of Array.from(frag.documentElement.childNodes)) body.insertBefore(doc.importNode(n, true), sect);

    zip.file("word/document.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, ""));
    return zip.generateAsync({ type: root.__zipOutType || "blob", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  }

  const api = { readWorkbook, summarize, buildCourseFile, WorkbookError };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Engine = api;
})(typeof window !== "undefined" ? window : globalThis);
