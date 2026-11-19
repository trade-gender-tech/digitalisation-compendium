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
| About the Compendium | Launch at MC14 (Yaoundé, side event of 25 March 2026), the joint proposal and its Co-Chairs/Coordinator, the work plans it supports, two objective cards | text from the joint communication WT/MIN(26)/4 • INF/TGE/COM/13/Rev.1 • INF/MSME/W/68/Rev.1 (27 February 2026) |
| Survey Themes | The six themes of the Members survey, as teal cards with an illustration | text in `index.html` |
| How to Use, Database | See below | |

The joint communication itself is not part of the package (it is kept with the source files in `3. Digitalisation Compendium/COM13R1.docx`) and is not offered for download.

**Colours.** The page uses the teal palette of the WTO trade and gender databases (deep teal, navy headings, gold call-to-action button). Every colour is a CSS variable defined once at the top of the stylesheet (`--teal`, `--navy`, `--gold`, …), so the theme can be changed in one place. The hero image is tinted teal by the stylesheet (`background-blend-mode`), so a replacement image takes the same colour. The dot colours were checked for colour-vision deficiency (dark teal for the finding, light teal for a partial answer, grey for the rest), and every figure is also written out, so nothing depends on colour or hover alone.

---

## The database

**Filters** (all combine):
- **Responses:** WTO Member (WTO Members, plus the countries and regions where WTO stakeholders report policies, e.g. Ghana or Cambodia; a stakeholder's response matches if it names a policy there), Respondent Type (WTO Members / WTO stakeholders), Region / Category, specific Respondents.
- **Questions:** Themes, specific Questions.
- **Keyword search** across every answer and explanation; matches are highlighted.

**Four views:**
0. **Policy view** (default): a catalogue of the policies, strategies, laws, trade agreement provisions, programmes, platforms, funds and networks that respondents *name* in their answers, so that a researcher can find measures without knowing the questionnaire. At the top, a **policy landscape**: one tile per policy area with the number of policies (click to filter; several can be combined). Below, a compact list: **one line per policy** (name, who reports it and, for WTO stakeholders, the country where it applies, and the instrument). Clicking a line opens it: the description quoted **word for word** from the response, any items the respondent lists under it, the links the respondent gave for it, and a button that opens the source answer in the Response view. "Expand all" opens every line shown. Each policy is labelled with an **instrument type**, one or more **policy areas**, a **scope** (domestic / international cooperation) and, where the response mentions one, a **target group**; these labels are added for browsing and are the only words on the cards that are not the respondent's. Instrument type, scope and target group filters on the left combine with each other, with the landscape and with the main filters above (respondent filters, themes and questions: a policy is kept if it is mentioned in one of the chosen questions; keyword: searches the policy names, descriptions and labels). Group by policy area (default; a policy with several areas appears under each), instrument or WTO Member; groups show ten lines with "Show all". Respondents whose answers name no specific policy are listed in a note under the cards. See **The policy catalogue** below.
1. **Response view**: one full response at a time: the respondent's identification fields as submitted (Time, WTO Member/Observer Government, Ministry/Institution), then every question with the answer options and the explanation, grouped by theme. A list of respondents on the left (a dropdown on phones) and Previous / Next.
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
  digitalization_policies.json      Policy view catalogue, written by tools/build_policies.py
  digitalization_policies.js        The same catalogue as a script, used only when index.html is opened from disk
images/
  site-header.png, site-footer.png  WTO site header and footer (stand-ins for the wto.org template)
  favicon.png                       Browser tab icon
  wto-trade-for-women-logo.png      WTO OMC | Trade for Women logo, used on the PDF of a response
  hero.png                          Hero background (a dotted world map, tinted teal by the stylesheet)
  pillar-1…4.png, purpose-*.png     Illustrations of the objective and Survey Theme cards
downloads/
  digitalization-compendium.xlsx    The compendium in Excel; also the source of the data files
tools/
  questionnaire_to_excel.py         Builds the compendium Excel from the raw questionnaire exports
  excel_to_json.py                  Rebuilds data/digitalization_data.json and .js from the Excel file
  build_policies.py                 Builds data/digitalization_policies.json and .js and checks every quote
  policy_catalogue/                 The policy catalogue, one file per respondent (curated; see below)
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
7. **Element IDs used by the script:** `btnExploreDb`, `btnExploreDbTop`, `dbCollapse`, `db-error`, `db-ui`, `keywordInput`, `f-country`, `f-group`, `f-area`, `f-respondent`, `f-theme`, `f-question`, `btnClear`, `resultsText`, `viewHint`, `view-policy`, `view-response`, `view-question`, `view-matrix`, `detailModal`, `detailModalTitle`, `detailModalBody`, `detailOpenResponse`, `vizTip`, `pdfFrame` (created by the script for the PDF).

---

## Data loading

The page holds no data. When it opens, the script loads **`data/digitalization_data.json`** and builds everything from it. Browsers do not let a page opened straight from disk (by double-clicking `index.html`) read a JSON file; in that case only, the script loads **`data/digitalization_data.js`**, the same data written by the same converter as a script. On a web server only the JSON file is used. The file names are set in `DATA_JSON` and `DATA_SCRIPT` at the end of the page script. The Policy view catalogue is loaded the same way (`POLICIES_JSON` / `POLICIES_SCRIPT`); if it is missing, the page still works and opens on the Response view.

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

### 3. Rebuild the policy catalogue

```
python3 tools/build_policies.py
```

Run it after every data rebuild. It stops with a list of errors if any catalogued name, description, listed item or link is no longer found word for word in the answers, so the Policy view can never show text that differs from the responses. New responses are **not** catalogued automatically: add a file for the new respondent in `tools/policy_catalogue/` (see below).

### The policy catalogue
Each file in `tools/policy_catalogue/` lists the policies named in one response. An entry points at the respondent's own text rather than copying it: the policy name as written, and the start and end of the passage that describes it (`x=(start, end)`) in a given answer (`q='Q9'`). The builder copies the passage between those two phrases from the data, so the description is always the respondent's wording. The labels (`t` instrument type, `a` policy areas, `g` target groups, `s` scope, `c` country for stakeholder answers) are the editorial part and use the codes listed at the top of `tools/build_policies.py`. The current catalogue (November 2026) covers 247 policies from 19 respondents; the responses of China, Mauritius and the Institute of International Studies (University of Chile) give links or general descriptions without naming specific policies.

### Written in `index.html`, to be edited by hand if they change
The About text (taken from the joint communication), the survey period, and the theme names and descriptions. Everything else is calculated from the data.

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
