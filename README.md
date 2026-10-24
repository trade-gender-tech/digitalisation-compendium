# Digitalisation Compendium — web package

Database page for **Empowering Women to Trade through Digitalisation: Policy Insights**, a joint initiative of the Informal Working Group on Trade and Gender and the Informal Working Group on MSMEs, prepared by the WTO Trade and Gender Office for publication on wto.org.

The page presents every response to the questionnaire *Empowering Women to Trade through Digitalization*: 18 WTO Members and 4 other stakeholders (22 responses, 42 questions), received between June and October 2026. **All answers, explanations and question wording are shown verbatim, as submitted.** Only personal data is left out (see "Updating the data").

The WTO site header, footer and breadcrumbs are shown as on the GitHub preview pages, so the page can be reviewed as it will look on wto.org.

---

## Quick start

1. Open **`index.html`** in any current browser (Chrome, Edge, Firefox or Safari). Double-clicking it is enough: no web server and no internet connection are needed.
2. Click **Explore Compendium** to open the database.

---

## What is on the page

| Section | Content | Source |
|---|---|---|
| Hero | Title as in the joint proposal (*Empowering Women to Trade through Digitalisation: Policy Insights*), the two Informal Working Groups, a one-line description; Explore and Excel download buttons | text in `index.html` |
| About the Compendium | Launch at MC14 (Yaoundé, side event of 25 March 2026), the joint proposal and its Co-Chairs/Coordinator, the work plans it supports, four objective cards, the two lines of work; Contributions panel (the regions and other stakeholders that responded, with names; no counts) | text from the joint communication WT/MIN(26)/4 • INF/TGE/COM/13/Rev.1 • INF/MSME/W/68/Rev.1 (27 February 2026); panel from the data |
| Survey Themes | The six themes of the Members survey, with their question ranges | text in `index.html`; ranges from the data |
| What the Responses Show | Six messages in words, not numbers (strategies, programmes, instruments, sectors, skills, data), each in a wide card with a line icon (drawn inline in the script, `ICONS`) and chips quoting the most frequent answers verbatim. Words such as "most" or "some" are worked out from the data (`amount()` in the script), so the wording stays true if responses change | data (`FINDINGS` in the script) |
| Insights from Other Stakeholders | Four verbatim quotes | data (`QUOTES` in the script) |
| How to Use, Database | See below | |

The joint communication itself is not part of the package (it is kept with the source files in `3. Digitalisation Compendium/COM13R1.docx`) and is not offered for download.

**Quotes are checked against the data.** Each quote in `QUOTES` is shown only if it is found word for word in that organization's answer, so a quote can never drift from the source.

**Colours.** The page uses the teal palette of the WTO trade and gender databases (deep teal, navy headings, gold call-to-action button). Every colour is a CSS variable defined once at the top of the stylesheet (`--teal`, `--navy`, `--gold`, …), so the theme can be changed in one place. The hero image is tinted teal by the stylesheet (`background-blend-mode`), so a replacement image takes the same colour. The dot colours were checked for colour-vision deficiency (dark teal for the finding, light teal for a partial answer, grey for the rest), and every figure is also written out, so nothing depends on colour or hover alone.

---

## The database

**Filters** (all combine):
- **Responses:** Respondent Type (WTO Members / Other stakeholders), Region / Category, specific Respondents.
- **Questions:** Themes, specific Questions.
- **Keyword search** across every answer and explanation; matches are highlighted.

**Three views:**
1. **Response view** (default): one full response at a time: the respondent's identification fields as submitted (Time, WTO Member/Observer Government, Ministry/Institution), then every question with the answer options and the explanation, grouped by theme. A list of respondents on the left (a dropdown on phones) and Previous / Next.
   **Download PDF** (Response view only): creates an A4 document of the response shown, in Verdana, with the WTO OMC | Trade for Women logo (`images/wto-trade-for-women-logo.png`), the respondent's details, every question shown with its answer and explanation verbatim, and "Page X of Y" in the footer. It contains only the questions currently shown: if the Themes or Questions filters are used, the PDF is an extract and says so. The button opens the browser's print dialog, where **Save as PDF** creates the file (suggested name: "Digitalisation Compendium – *respondent*"). This keeps the text selectable and the links clickable, and uses the Verdana font installed on the computer: the font itself is licensed by Microsoft and cannot be shipped with the page.
2. **Question view:** one card per question with a bar chart of the answers among the responses shown, and every respondent's answer and explanation (5 shown, "Show all" for the rest; long explanations have "Read more"). Clicking a name opens that response at that question.
3. **Matrix view:** all responses side by side (WTO Members 18 × 30, other stakeholders 4 × 12). Cells show the answer options; "¶ Explanation" marks an explanation. Hover shows a preview; click opens the full answer in a pop-up, with a button to open the full response.

---

## Folder structure

```
index.html                          The page: HTML, CSS (one <style> block) and the page script (one <script> block)
README.md                           This file
data/
  digitalization_data.json          Questionnaire responses, loaded by index.html
  digitalization_data.js            The same data as a script, used only when index.html is opened from disk
images/
  site-header.png, site-footer.png  WTO site header and footer (stand-ins for the wto.org template)
  favicon.png                       Browser tab icon
  wto-trade-for-women-logo.png      WTO OMC | Trade for Women logo, used on the PDF of a response
  hero.png, pillar-*.png, purpose-*.png   Hero and illustrations, reused from the IWG Compendium for now
downloads/
  digitalization-compendium.xlsx    The compendium in Excel; also the source of the data files
tools/
  questionnaire_to_excel.py         Builds the compendium Excel from the raw questionnaire exports
  excel_to_json.py                  Rebuilds data/digitalization_data.json and .js from the Excel file
  requirements.txt                  Python package needed by both tools
vendor/                             jquery/  bootstrap/  bootstrap-select/  open-sans/
```

All paths in `index.html` are relative, so the folder works wherever it is copied.

---

## Publishing on wto.org

1. **Header, footer and breadcrumbs.** `<header class="site-header">`, `<footer class="site-footer">` and `<nav id="breadcrumbsDiv">` are stand-ins copied from wto.org. Replace them with the standard WTO template and delete `images/site-header.png` and `images/site-footer.png`.
2. **Bootstrap.** The page is built on Bootstrap 3. Replace `vendor/bootstrap/css/bootstrap.min.css` with `wto.css`. Bootstrap's JavaScript is used for the matrix pop-up (`#detailModal`).
3. **Libraries already on the site.** If the WTO template already loads jQuery, Bootstrap JS or bootstrap-select, remove the matching `vendor/` tags. The page script is plain ES5 and works with jQuery 1.12 or later. No chart library is needed: the charts are HTML/CSS drawn by the script.
4. **Font.** If `wto.css` already provides Open Sans, remove the `vendor/open-sans/open-sans.css` link.
5. **CSS and script** can be moved into their own files unchanged. The script must load after the libraries. It loads the data itself (see **Data loading** below).
6. **Style scope.** Page styles are scoped to `main.digi-compendium` or to class names used only on this page. A few names are generic (`.tag`, `.tags`, `.page-wrap`, `.filter-panel`, `.stats-bar`); prefix them if `wto.css` uses the same names.
7. **Element IDs used by the script:** `reach`, `cs-theme-count-1`…`6`, `fd-grid`, `sh-quotes`, `btnExploreDb`, `btnExploreDbTop`, `dbCollapse`, `db-error`, `db-ui`, `keywordInput`, `f-group`, `f-area`, `f-respondent`, `f-theme`, `f-question`, `btnClear`, `resultsText`, `viewHint`, `view-response`, `view-question`, `view-matrix`, `detailModal`, `detailModalTitle`, `detailModalBody`, `detailOpenResponse`, `vizTip`, `pdfFrame` (created by the script for the PDF).

---

## Data loading

The page holds no data. When it opens, the script loads **`data/digitalization_data.json`** and builds everything from it. Browsers do not let a page opened straight from disk (by double-clicking `index.html`) read a JSON file; in that case only, the script loads **`data/digitalization_data.js`**, the same data written by the same converter as a script. On a web server only the JSON file is used. The file names are set in `DATA_JSON` and `DATA_SCRIPT` at the end of the page script.

---

## Updating the data

### 1. New responses: rebuild the Excel from the questionnaire exports

```
pip install -r tools/requirements.txt     (first time only)
python3 tools/questionnaire_to_excel.py --members "MEMBERS EXPORT.xlsx" --stakeholders "STAKEHOLDERS EXPORT.xlsx"
```

The inputs are the two Formstack exports. The script writes `downloads/digitalization-compendium.xlsx` (sheets **About**, **WTO Members**, **Other Stakeholders**, **Questions**). Every cell is copied **verbatim**. The only changes:

- **Personal data is left out:** respondent names, positions, e-mails, phone numbers, browser, IP address, Formstack ID, location and upload links (Formstack admin links). Contact details written inside an answer are replaced by `[contact details removed]` (currently one place: Nigeria, question 29).
- Two columns are added at the front: **Respondent** (standard WTO name, e.g. "Mexico" for "México", "Russian Federation" for "WTO MEMBER") and **Region** / **Category**. The respondent's own wording stays in the "WTO Member/Observer Governement:" column.
- Each question column header gets a code: `Q1 | 1. Do you explicitly…`; its explanation column `Q1.1 | 1.1. Please elaborate…` (other stakeholders: `S1`, `S1.a`, `S1.b`).

If a new Member responds, the script stops and names it: add it to `MEMBERS` at the top of the script (name and region). If Formstack adds a column it does not know, the script stops rather than silently dropping it. Use `--period "between … and …"` to update the period on the About sheet.

The script was checked cell by cell against the current exports: all 1,149 kept cells are identical, except the one redaction.

### 2. Rebuild the data files used by the page

```
python3 tools/excel_to_json.py
```

Then refresh `index.html`. The Excel file can also be edited by hand, followed by this step; keep the sheet names and the coded column headers.

### Written in `index.html`, to be edited by hand if they change
The About text (taken from the joint communication), the survey period, the theme names and descriptions, the sentence wording of the six findings (`FINDINGS`; their numbers are calculated) and the quotes (`QUOTES`). Everything else is calculated from the data.

### Notes on the data
- Explanations are reproduced as submitted, including placeholders such as "-" or "n.a." and Brazil's response in Portuguese.
- Question 9 has one answer spelled "digitalisation" and eight spelled "digitalization"; they appear as two answers because the data is verbatim.
- The six themes group the questions for navigation; they follow the order of the questionnaire.

---

## Libraries and requirements

| Library | Version | Used for | Licence |
|---|---|---|---|
| jQuery | 3.7.1 | DOM and events | MIT |
| Bootstrap | 3.4.1 | Base CSS, pop-up (`wto.css` replaces the CSS on wto.org) | MIT |
| bootstrap-select | 1.13.18 | Searchable, multi-select filter dropdowns | MIT |
| Open Sans | variable, 400/600/700 | Page font (self-hosted) | SIL Open Font License 1.1 |

To rebuild the data you also need **Python 3.8+** with **openpyxl** (`tools/requirements.txt`). Nothing is loaded from the internet; the page works offline.
