import { useState, useEffect } from "react"
import { Pivot, PivotItem } from "@fluentui/react"
import { Analyses } from "./Analyses"
import { Header } from "./Header"
import { Upload } from "./Upload"
import { Actions } from "./Actions"
import { Support } from "./Support"
import { Tidystats } from "../classes/Tidystats"

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
      <Header />
      <Pivot
        aria-label="tidystats navigation"
        styles={{ root: { marginBottom: "1rem" } }}
      >
        <PivotItem headerText="Statistics">
          <Upload setTidystats={setTidystats} />
          {tidystats && <Analyses tidystats={tidystats} />}
        </PivotItem>
        <PivotItem headerText="Actions">
          <Actions tidystats={tidystats} />
        </PivotItem>
        <PivotItem headerText="Support">
          <Support />
        </PivotItem>
      </Pivot>
    </>
  )
}
