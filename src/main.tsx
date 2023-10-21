import React from "react"
import ReactDOM from "react-dom/client"
import { App } from "./App"

import { FluentProvider, webLightTheme } from "@fluentui/react-components"

window.Office.onReady((info) => {
  console.log(`Office.js is now ready in ${info.host} on ${info.platform}`)

  ReactDOM.createRoot(document.getElementById("root")!).render(
    <React.StrictMode>
      <FluentProvider theme={webLightTheme}>
        {/* <StrictMode> */}
        <App host={info.host} />
        {/* </StrictMode> */}
      </FluentProvider>
    </React.StrictMode>
  )
})
