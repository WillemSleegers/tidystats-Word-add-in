# News

## 1.1.1

### Changes

* Updated all packages to their latest versions, including React 19 and Vite 7.
* Replaced references to tidystats.io with a support page hosted on Vercel.
* Fixed a bug where `context.sync` was not being called correctly in the insert statistics functions.
* Fixed a redundant lookup in the `findStatistic` method.
* Fixed a bug where the confidence interval level (e.g., "95") would be overwritten with the lower bound value when updating statistics.
* Renamed the "Analyses" heading to "Statistics" in the Statistics tab.
* Refactored the codebase to improve code organization and maintainability.

## 1.1.0

### New features

* The uploaded file now stays active after closing the add-in, so it's not needed to re-upload the same statistics file after closing the add-in
* The uploaded file can be removed by clicking on 'Remove file' in the dropdown menu next to the upload button.
* Reported statistics can be replaced with a fixed value, and highlighted.
* Help messages have been added to explain the functionality of the add-in.
* Expanding statistics is now a bit smarter. Some statistics are automatically expanded or hidden depending on the number of groups of statistics.
* Certain groups of statistics (e.g., regression coefficients, ANOVA terms) can be inserted as tables.

### Changes

* Complete overhaul of the UI to create a more modern look and to improve the user-friendliness of the add-in.
* Removed citation buttons and replaced them with a link to the website about how to cite tidystats.
* Statistics are no longer indented to preserve screen space.

## 1.0.0

First release.
