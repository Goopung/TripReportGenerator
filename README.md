# TripReportGenerator

This is a standalone app for automatically generating conference business trip result reports.
https://tripreportgenerator.streamlit.app/

## Major Updates

1. The `Extracted Overview Text or Direct Input` field can now be entered and edited.
2. The `Purpose of Business Trip` is automatically generated and immediately reflected.
3. The `Relevance to This Research and Summary of Major Sessions` is automatically generated and immediately reflected.
4. The `Detailed Schedule` is automatically generated and immediately reflected.
5. In the submission document checklist UI, `Missing` is displayed in red and `Completed` is displayed in green.
6. The phrase `File name:` has been removed from the report body.
7. The phrase `This file is included in the ZIP package as an original attachment.` has been removed from the report body.
8. The submission document checklist is displayed only in the UI and is not included in the report.
9. When selected, a separate DOCX justification letter can be generated.

## Installation

```bash
cd TripReportGeneratorApp_v3
python -m venv .venv
````

Windows:

```bash
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
streamlit run app.py
```

macOS / Linux:

```bash
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
streamlit run app.py
```

## .env Configuration

```env
OPENAI_API_KEY=your_api_key_here
OPENAI_MODEL=gpt-5.5
```

If `gpt-5.5` is not available for your current account, change it to an available model name.

## When Korean Text Appears as Square Boxes in PDF

This is a PDF font issue. In v4, the app automatically searches for Korean fonts such as `Malgun Gothic` on Windows, `AppleGothic` on macOS, and `Noto/Nanum` fonts on Linux.

If Korean text still appears as square boxes on your machine, manually specify the Korean font path in the `.env` file:

Windows:

```env
KOREAN_FONT_PATH=C:/Windows/Fonts/malgun.ttf
```

macOS:

```env
KOREAN_FONT_PATH=/Library/Fonts/AppleGothic.ttf
```

Linux:

```env
KOREAN_FONT_PATH=/usr/share/fonts/truetype/nanum/NanumGothic.ttf
```

## v5 Justification Letter PDF Template

* The justification letter is filled directly onto the PDF based on `templates/사유서_template.pdf`.
* The final generated ZIP file includes the completed `사유서.pdf`.
* The `Year / Month / Day` section at the bottom of the justification letter is automatically filled based on the report creation date.
* The `Principal Investigator: (Seal)` line at the bottom of the justification letter is automatically filled with the name of the principal investigator.

## Font Configuration

When generating PDF files, Korean text prioritizes Batang-style fonts, while English text and numbers prioritize Times New Roman-style fonts.

Font files are not included in the package. On Windows, the following paths are usually detected automatically:

```env
KOREAN_FONT_PATH=C:/Windows/Fonts/batang.ttc
ENGLISH_FONT_PATH=C:/Windows/Fonts/times.ttf
```

If you want to use different fonts, specify the paths directly in the `.env` file.

## v6 Fixes

* Korean fonts are now forcibly applied inside PDF tables as well, preventing text such as `Traveler / Destination / Date / Trip Details` from appearing as black boxes.
* The domestic/international selection on the PDF cover page no longer displays `checked / unchecked`; it now uses square box symbols `□ / ■`.
* Korean fonts now prioritize Batang-style fonts by default, while English fonts prioritize Times New Roman by default.
* Font files are not included in the project. To specify fonts manually, set the following values in the `.env` file:

  * `KOREAN_FONT_PATH=C:/Windows/Fonts/batang.ttc`
  * `ENGLISH_FONT_PATH=C:/Windows/Fonts/times.ttf`

## v7 Style Update

* The report layout has been redesigned with reference to the AAAI 2026 international business trip result report style:

  * The cover page has been changed to a three-part structure: `Conference Attendance / International Business Trip Result Report / Date of Preparation`.
  * Main section headings in the body now use the format `□ Purpose of Business Trip`, `□ Trip Period`, `□ Traveler`, and `□ Detailed Schedule`.
  * The traveler information and detailed schedule sections have been changed into compact tables.
  * Supporting documents, receipts, and trip photos are now arranged in a grid-style layout.
  * A `□ Expected Outcomes` section is automatically added on the final page.
  * PDF body pages now include page numbers in the footer using the `- page -` format.

## v8 Fixes

* Fixed the `name '_register_pdf_fonts' is not defined` error.
* Restored the PDF Korean font registration function and continued using the Batang / Times New Roman priority rule.
