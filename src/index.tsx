import { StrictMode } from "react"
import { ThemeProvider } from "@fluentui/react"
import { initializeIcons } from "@fluentui/font-icons-mdl2"
import { createRoot } from "react-dom/client"
import { App } from "./components/App"
import "./index.css"

initializeIcons()

const element = document.getElementById("app") as HTMLElement
const root = createRoot(element)

window.Office.onReady((info) => {
  console.log(`Office.js is now ready in ${info.host} on ${info.platform}`)

  root.render(
    <ThemeProvider>
      <StrictMode>
        <App host={info.host} />
      </StrictMode>
    </ThemeProvider>
  )
})
