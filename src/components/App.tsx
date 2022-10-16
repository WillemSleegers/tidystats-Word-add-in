import { useState, useEffect } from "react"
import { Pivot, PivotItem } from "@fluentui/react"

import { Tidystats } from "../classes/Tidystats"

import { AnalysesTable } from "./AnalysesTable"
import { Logo } from "./Logo"
import { Upload } from "./Upload"
import { Actions } from "./Actions"
import { Support } from "./Support"

import logoSrc from "../assets/tidystats-icon.svg"

type AppProps = {
  host: Office.HostType
}

export const App = (props: AppProps) => {
  const { host } = props

  const [tidystats, setTidystats] = useState<Tidystats>()

  useEffect(() => {
    if (host === Office.HostType.Word) {
      const savedStatistics = Office.context.document.settings.get("data")
      if (savedStatistics) {
        const savedTidystats = new Tidystats(JSON.parse(savedStatistics))
        setTidystats(savedTidystats)
      }
    }
  }, [host])

  return (
    <>
      <Logo title="tidystats" logo={logoSrc} />
      <div style={{ margin: "0 10px" }}>
        <Pivot
          aria-label="tidystats navigation"
          styles={{ root: { marginBottom: "1rem" } }}
        >
          <PivotItem headerText="Statistics">
            <Upload setTidystats={setTidystats} />
            {tidystats && <AnalysesTable tidystats={tidystats} />}
          </PivotItem>
          <PivotItem headerText="Actions">
            <Actions tidystats={tidystats} />
          </PivotItem>
          <PivotItem headerText="Support">
            <Support />
          </PivotItem>
        </Pivot>
      </div>
    </>
  )
}
