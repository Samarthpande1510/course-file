<p align="center">
  <img src="public/hero-bg.jpg" alt="Course File Builder" width="100%">
</p>

<h1 align="center">Course File Builder</h1>

<p align="center">
  Upload your attainment workbook. Get your course file.<br>
  No reformatting in Word, no server, nothing leaves your computer.
</p>

<p align="center">
  <a href="#try-it">Try it</a> ·
  <a href="#how-it-works">How it works</a> ·
  <a href="#deploy-to-vercel">Deploy</a> ·
  <a href="#make-it-yours">Make it yours</a>
</p>

---

## Why this exists

Every semester, teachers fill in an Excel sheet with course outcomes, marks and attainment, and then copy pieces of it into a Word course file by hand, fixing fonts and tables as they go.

This tool does that copying. You drop in the Excel file, it fills the department's course file template, and you download the Word document. The formatting is already right because it starts from the department's own sample.

## What you get today

- The **cover page** with course name, semester, program, academic year and faculty.
- The **details table** (course code, degree, program, year and semester).
- The **table of contents**, exactly as in the sample.

The workbook is already fully read (outcomes, mapping, attainment, survey, student marks). The code that writes those into the document is built and switched off for now. See [Make it yours](#make-it-yours).

## Try it

You need [Node.js](https://nodejs.org) installed.

```bash
npm install
npm run dev
```

Open <http://localhost:3000>, choose your `.xlsx` file, check the summary on the right, and press **Generate course file**.

## How it works

```
 Excel workbook (.xlsx)                Word template (.docx)
        │                                      │
        ▼                                      ▼
   read in the browser  ───────►  fill cover page + details
      with ExcelJS                    edit with JSZip
                                               │
                                               ▼
                                     Course File.docx  (downloads)
```

Everything runs in the browser. There is no backend, so there is nothing to host except static files, and student marks are never uploaded anywhere.

| File | What it does |
| --- | --- |
| `public/index.html` | The page and its styling |
| `public/app.js` | Upload, summary and download handling |
| `public/engine.js` | Reads the workbook and fills the template |
| `public/template.docx` | The department's sample course file |
| `public/vendor/` | ExcelJS and JSZip, bundled so the app works offline |

## Odd files are handled

| You upload | What happens |
| --- | --- |
| A different file renamed to `.xlsx`, or a damaged one | A plain error explaining what is wrong |
| An Excel file that is not this workbook | An error listing the sheets it found |
| The right workbook with a sheet missing or scrambled | The file still generates, with a yellow note about what was skipped |
| Any program name (`B Tech`, `BTI`, `MBA Tech`...) | Copied through as written |

## Deploy to Vercel

1. Import this repository in Vercel.
2. Deploy. `vercel.json` already points Vercel at the `public` folder and no build step is needed, so leave the framework preset on **Other**.

Because the site is static, there are no servers to keep awake.

## Make it yours

- **Turn on the full document.** In `public/engine.js`, change `FULL_SECTIONS` from `false` to `true`. The file will then also include course policy, CO attainment tables, the course end survey, every student's ICA marks and the FCAR section. Items Excel cannot hold (academic calendar, question papers, study material) get a marked placeholder.
- **Change the template.** Replace `public/template.docx` with your own. Keep the cover lines named *Course Name*, *Semester*, *Program*, *Academic Year* and *Faculty Name*, and keep the details table as the first table.
- **Change the help number.** The footer in `index.html` has a `Call *****` line. Put your number there.
- **Change the background.** Swap `public/hero-bg.jpg` for any image you like.

## Good to know

- Built and tested against the *Computer Networks, B Tech Sem 5* workbook. Other workbooks should work if they use the same sheet names (`Course Details`, `COs and Mapping`, and so on).
- The title page uses a large font, so a very long course name will wrap onto a second line.
- The table of contents has no page numbers, by design.

## Built with

Plain HTML, CSS and JavaScript, with [ExcelJS](https://github.com/exceljs/exceljs) and [JSZip](https://stuk.github.io/jszip/).
