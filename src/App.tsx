import { useState } from "react"
import {
  Tab,
  TabList,
  TabValue,
  SelectTabEvent,
  SelectTabData,
  makeStyles,
} from "@fluentui/react-components"
import { StatisticsTab } from "./components/tabs/StatisticsTab"
import { Header } from "./components/Header"
import { Upload } from "./components/Upload"
import { ActionsTab } from "./components/tabs/ActionsTab"
import { SupportTab } from "./components/tabs/SupportTab"
import { Analysis } from "./types"
import { parseAnalyses } from "./utils/parseAnalyses"
import { getSettingsData } from "./word/settings"

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

  const [analyses, setAnalyses] = useState<Analysis[] | undefined>(() => {
    if (host === Office.HostType.Word) {
      const savedStatistics = getSettingsData("statistics")
      if (savedStatistics) {
        return parseAnalyses(JSON.parse(savedStatistics))
      }
    }
    return undefined
  })
  const [selectedTab, setSelectedTab] = useState<TabValue>("statistics")

  const onTabSelect = (_event: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value)
  }

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
              <Upload setAnalyses={setAnalyses} />
              {analyses && <StatisticsTab analyses={analyses} />}
            </>
          )}
          {selectedTab === "actions" && <ActionsTab analyses={analyses} />}
          {selectedTab === "support" && <SupportTab />}
        </div>
      </div>
    </div>
  )
}
