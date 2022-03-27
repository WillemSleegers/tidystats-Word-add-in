import ReactDOM from "react-dom"
import "./index.css"
import { App } from "./components/App"

const Office = window.Office

let isOfficeInitialized = false

Office.initialize = () => {
  isOfficeInitialized = true

  ReactDOM.render(
    <App isOfficeInitialized={isOfficeInitialized} />,
    document.getElementById("root")
  )
}
