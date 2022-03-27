# How To

Project creation date: 2022-01-08

## Setup

- Run `npx create-react-app tidystats-word-add-in --template typescript` to create the app
- Remove redundant files (e.g., `setupTests.js`)
- Install `npm install office-addin-debugging`
- Copy a manuscript from a different project or get one created using `yo office`
- In `package.json`, replace `"react-scripts start"` with `"HTTPS=true BROWSER=none react-scripts start"` to create a HTTPS server and disable automatically opening the browser
- In `package.json` add `"word": "office-addin-debugging start manifest.xml"` under "scripts" to launch the Word add-in
- Run `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true` to enable the web inspector in the Word add-in

- Install `npm install @types/office-js --save-dev`
- Install `npm i @fluentui/react`
- Install `npm install --save styled-components`
- Install `npm i --save-dev @types/styled-components`

## Usage

- Start the dev server with `npm start`
- Open the Word add-in with `npm run word`
