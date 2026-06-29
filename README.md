# Deutsch Heimkehr

Deutsch Heimkehr is a workbook-driven German learning app hosted with GitHub Pages.

## Structure

```text
index.html
css/app.css
js/app.js
source/course.json
source/xlsx/A1/unit-1/*.xlsx
```

## Lesson loading

The app reads `source/course.json`, then downloads the listed `.xlsx` lesson workbooks directly from the repository.

The `.xlsx` files remain the authoring format.

## Adding a lesson

1. Add the workbook to the correct folder, for example:

```text
source/xlsx/A1/unit-1/Deutsch_Heimkehr_A1_U1_L6_v1.0.xlsx
```

2. Add that filename to `source/course.json`.

3. Commit and push.

## Current release

Deutsch Heimkehr v3.1
