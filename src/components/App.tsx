import { useState, useEffect } from "react"
import {
  Tab,
  TabList,
  TabValue,
  SelectTabEvent,
  SelectTabData,
} from "@fluentui/react-components"
import { Analyses } from "./Analyses"
import { Header } from "./Header"
import { Upload } from "./Upload"
import { Actions } from "./Actions"
import { Support } from "./Support"
import { Tidystats } from "../classes/Tidystats"
import { getSettingsData } from "../functions/settings"

type AppProps = {
  host: Office.HostType
}

export const App = (props: AppProps) => {
  const { host } = props

  const [tidystats, setTidystats] = useState<Tidystats>()
  const [selectedTab, setSelectedTab] = useState<TabValue>("statistics")

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
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
    <>
      <Header />

      <TabList selectedValue={selectedTab} onTabSelect={onTabSelect}>
        <Tab value="statistics">Statistics</Tab>
        <Tab value="actions">Actions</Tab>
        <Tab value="support">Support</Tab>
      </TabList>

      {selectedTab === "statistics" && <Upload setTidystats={setTidystats} />}
      {selectedTab === "statistics" && tidystats && (
        <Analyses tidystats={tidystats} />
      )}
      {selectedTab === "actions" && <Actions tidystats={tidystats} />}
      {selectedTab === "support" && <Support />}
    </>
  )
}
