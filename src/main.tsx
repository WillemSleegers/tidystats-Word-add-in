import { createRoot } from "react-dom/client"
import { App } from "./App"
import "./index.css"

import { FluentProvider, webLightTheme } from "@fluentui/react-components"

window.Office.onReady((info) => {
  console.log(`Office.js is now ready in ${info.host} on ${info.platform}`)

  if (info.host) {
    // Fluent UI React v. 9 uses modern JavaScript syntax that is not supported in Trident (Internet Explorer) or EdgeHTML (Edge Legacy), so this add-in won't work in Office versions that use these webviews. The script below makes the following div display when an unsupported webview is in use.
    if (
      navigator.userAgent.indexOf("Trident") !== -1 ||
      navigator.userAgent.indexOf("Edge") !== -1
    ) {
      createRoot(document.getElementById("root")!).render(
        <div id="tridentmessage">
          This add-in will not run in your version of Office. Please upgrade
          either to perpetual Office 2021 (or later) or to a Microsoft 365
          account.
        </div>
      )
    } else {
      createRoot(document.getElementById("root")!).render(
        <FluentProvider theme={webLightTheme}>
          {/* <StrictMode> */}
          <App host={info.host} />
          {/* </StrictMode> */}
        </FluentProvider>
      )
    }
  } else {
    createRoot(document.getElementById("root")!).render(
      <div>Loaded outside of Microsoft Word.</div>
    )
  }
})
