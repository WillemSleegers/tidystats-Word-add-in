# How To

Project creation date: 2022-01-08

## Setup

### First time setup

- Run `npx create-react-app tidystats-word-add-in --template typescript` to create the app
- Remove redundant files (e.g., `setupTests.js`)
- Install `npm install office-addin-debugging`
- Copy a manuscript from a different project or get one created using `yo office`
- In `package.json`, replace `"react-scripts start"` with `"HTTPS=true BROWSER=none react-scripts start & office-addin-debugging start manifest.xml"` to create a HTTPS server, disable automatically opening the browser, and start Word with the default manifest file.
- Optional: In `package.json` add `"word": "office-addin-debugging start manifest.xml"` under "scripts" to have a script that launches the Word add-in
- Run `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true` to enable the web inspector in the Word add-in
- Install `npm install @types/office-js --save-dev`
- Install `npm i @fluentui/react`
- Install `npm install --save styled-components`
- Install `npm i --save-dev @types/styled-components`

### Later setup

- Run `npm install`

## Usage

- Start the dev server with `npm start`
