import { createRoot } from "react-dom/client"
import { ThemeProvider } from "@fluentui/react"
import { initializeIcons } from "@fluentui/font-icons-mdl2"
import { App } from "./components/App"
import "./index.css"

initializeIcons()

const container = document.getElementById("container")!
const root = createRoot(container)

window.Office.onReady((info) => {
  console.log(`Office.js is now ready in ${info.host} on ${info.platform}`)

  root.render(
    <ThemeProvider>
      <App host={info.host} />
    </ThemeProvider>
  )
})
