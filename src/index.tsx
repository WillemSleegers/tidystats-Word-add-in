import ReactDOM from "react-dom"
import "./index.css"
import { App } from "./components/App"

const Office = window.Office

let isOfficeInitialized = false
let host = ""
let savedFileName: string | null = null
let savedStatistics: string | null = null

Office.onReady((info) => {
  isOfficeInitialized = true

  if (info.host === Office.HostType.Word) {
    host = "Word"

    // Retrieve saved data
    savedFileName = Office.context.document.settings.get("fileName")
    savedStatistics = Office.context.document.settings.get("data")

    if (savedFileName === "") savedFileName = null
    if (savedStatistics === "") savedStatistics = null
  }

  ReactDOM.render(
    <App
      isOfficeInitialized={isOfficeInitialized}
      host={host}
      savedFileName={savedFileName}
      savedStatistics={savedStatistics}
    />,
    document.getElementById("root")
  )

  console.log(`Office.js is now ready in ${info.host} on ${info.platform}`)
})
