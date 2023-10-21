import { useState, useEffect } from "react"
import {
  Tab,
  TabList,
  TabValue,
  SelectTabEvent,
  SelectTabData,
  makeStyles,
} from "@fluentui/react-components"
import { Analyses } from "./components/Analyses"
import { Header } from "./components/Header"
import { Upload } from "./components/Upload"
import { Actions } from "./components/Actions"
import { Support } from "./components/Support"
import { Tidystats } from "./classes/Tidystats"
import { getSettingsData } from "./functions/settings"

const useStyles = makeStyles({
  app: {},
  main: {},
  content: {
    marginLeft: "1rem",
    marginRight: "1rem",
    marginBottom: "1rem",
  },
})

type AppProps = {
  host: Office.HostType
}

export const App = (props: AppProps) => {
  const { host } = props
  const styles = useStyles()

  const [tidystats, setTidystats] = useState<Tidystats>()
  const [selectedTab, setSelectedTab] = useState<TabValue>("statistics")

  const onTabSelect = (_event: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value)
  }

  useEffect(() => {
    if (host === Office.HostType.Word) {
      const savedStatistics = getSettingsData("statistics")

      if (savedStatistics) {
        const savedTidystats = new Tidystats(JSON.parse(savedStatistics))
        setTidystats(savedTidystats)
      }
    }
  }, [])

  return (
    <div className={styles.app}>
      <Header />
      <div className={styles.main}>
        <TabList
          selectedValue={selectedTab}
          onTabSelect={onTabSelect}
          aria-label="Tabs"
        >
          <Tab value="statistics" aria-label="Statistics">
            Statistics
          </Tab>
          <Tab value="actions" aria-label="Actions">
            Actions
          </Tab>
          <Tab value="support" aria-label="Support">
            Support
          </Tab>
        </TabList>

        <div className={styles.content}>
          {selectedTab === "statistics" && (
            <>
              <Upload setTidystats={setTidystats} />
              {tidystats && <Analyses tidystats={tidystats} />}
            </>
          )}
          {selectedTab === "actions" && <Actions tidystats={tidystats} />}
          {selectedTab === "support" && <Support />}
        </div>
      </div>
    </div>
  )
}
