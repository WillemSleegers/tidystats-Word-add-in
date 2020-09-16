# How to Create Word Add-ins

## Step 1

- Install node via brew: `brew install node`
- Install yo office: `npm install -g yo generator-office`

## Step 2

- Create the add-in: `yo office`
- Run `npm audit` to see if there are any vulnerabilities

### Errors

- "found 1 high severity vulnerability"
	- Run: `npm audit fix --force`

## Step 3

- cd into the folder
- Install Fluent UI react: `npm install @fluentui/react`

### Errors:

- "Error: Cannot find module 'nan'"
	- Run `npm install` manually

- "You must install peer dependencies yourself."
	- Run `npm install --save-dev <name of package>` for each package

## Step 3

- Start the dev server: `npm run dev-server`

### Errors:

- "ValidationError: Invalid options object. Copy Plugin has been initialized using an options object that does not match the API schema."

Change:

```
new CopyWebpackPlugin([
        {
          to: "taskpane.css",
          from: "./src/taskpane/taskpane.css"
        }
      ])
```

To:

```
new CopyWebpackPlugin({
        patterns: [
          {
            from: "./src/taskpane/taskpane.css", 
            to: "taskpane.css" 
          },
        ],
      }),
```
