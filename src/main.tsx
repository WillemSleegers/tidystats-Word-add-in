import { createRoot } from "react-dom/client"
import { App } from "./App"

import { FluentProvider, webLightTheme } from "@fluentui/react-components"

window.Office.onReady((info) => {
  console.log(`Office.js is now ready in ${info.host} on ${info.platform}`)

  if (info.host) {
    createRoot(document.getElementById("root")!).render(
      <FluentProvider theme={webLightTheme}>
        {/* <StrictMode> */}
        <App host={info.host} />
        {/* </StrictMode> */}
      </FluentProvider>
    )
  } else {
    createRoot(document.getElementById("root")!).render(
      <div>Loaded outside of Microsoft Word.</div>
    )
  }
})
