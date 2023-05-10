import { StrictMode } from "react"
import { FluentProvider, webLightTheme } from "@fluentui/react-components"
import { createRoot } from "react-dom/client"
import { App } from "./App"

const element = document.getElementById("app") as HTMLElement
const root = createRoot(element)

window.Office.onReady((info) => {
  console.log(`Office.js is now ready in ${info.host} on ${info.platform}`)

  root.render(
    <FluentProvider theme={webLightTheme}>
      {/* <StrictMode> */}
      <App host={info.host} />
      {/* </StrictMode> */}
    </FluentProvider>
  )
})
