# Deutsch Heimkehr

Workbook-driven German course app.

## Current structure

```text
Deutsch_Heimkehr/
  index.html
  css/
    app.css
  js/
    app.js
  lessons/
    A1/
      unit-1/
  source/
    xlsx/
      A1/
        unit-1/
```

## Current lesson format

Lesson workbooks are still authored as `.xlsx` files. Put editable workbook sources in:

```text
source/xlsx/A1/unit-1/
```

The current app can still load lesson workbooks using the browser folder picker.

## Recommended next step

Keep `.xlsx` as the authoring format for now. Later, add an export process:

```text
.xlsx -> .json
```

Then the app can load lessons directly from GitHub Pages without needing the folder picker.

## GitHub Pages

For GitHub Pages, place this project at the root of a repository and enable Pages from the repository settings.
