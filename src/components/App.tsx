import { useState } from "react"

import { Tidystats } from "../classes/Tidystats"

import { AnalysesTable } from "./AnalysesTable"
import { Logo } from "./Logo"
import { Upload } from "./Upload"
import { Progress } from "./Progress"
import { Actions } from "./Actions"
import { Support } from "./Support"

import styled from "styled-components"
import { initializeIcons } from "@fluentui/font-icons-mdl2"
import { MessageBar, MessageBarType, Pivot, PivotItem } from "@fluentui/react"

import logoSrc from "../assets/tidystats-icon.svg"

initializeIcons()

const Main = styled.div`
  margin-left: 0.5rem;
  margin-right: 0.5rem;
  margin-bottom: 0.5rem;
`

type AppProps = {
  isOfficeInitialized: boolean
  host: string
  savedFileName: string | null
  savedStatistics: string | null
}

const App = (props: AppProps) => {
  const { isOfficeInitialized, host, savedFileName, savedStatistics } = props

  const [fileName, setFileName] = useState(savedFileName)
  const [tidystats, setTidystats] = useState(
    savedStatistics === null ? null : new Tidystats(JSON.parse(savedStatistics))
  )

  const statisticsUpload = (
    <Upload
      host={host}
      fileName={fileName}
      setFileName={setFileName}
      setTidystats={setTidystats}
    />
  )

  let content
  if (isOfficeInitialized) {
    if (tidystats) {
      content = <AnalysesTable tidystats={tidystats} />
    }
  } else {
    content = <Progress message="Please sideload your addin to see app body." />
  }

  const actionContent = <Actions tidystats={tidystats} />
  const support = <Support />

  return (
    <>
      <Logo title="tidystats" logo={logoSrc} />

      <Main>
        {host !== "Word" && (
          <div style={{ marginTop: "0.5rem" }}>
            <MessageBar messageBarType={MessageBarType.warning}>
              Add-in loaded outside of Microsoft Word; functionality is limited.
            </MessageBar>
          </div>
        )}
        <Pivot aria-label="tidystats navigation">
          <PivotItem headerText="Statistics">
            {statisticsUpload}
            {isOfficeInitialized && content}
          </PivotItem>
          <PivotItem headerText="Actions">{actionContent}</PivotItem>
          <PivotItem headerText="Support">{support}</PivotItem>
        </Pivot>
      </Main>
    </>
  )
}

export { App }
