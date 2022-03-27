import styled from "styled-components"

import { Analysis } from "../classes/Analysis"

import { Row } from "./Row"
import { RowName } from "./RowName"
import { RowValue } from "./RowValue"

import { GroupRows } from "./GroupRows"
import { StatisticsRows } from "./StatisticsRows"
import { Collapsible } from "./Collapsible"

const AnalysisDiv = styled.div`
  margin-top: 4px;
  margin-bottom: 4px;
`

type AnalysisRowsProps = {
  analysis: Analysis
}

const AnalysisRows = (props: AnalysisRowsProps) => {
  const { analysis } = props

  // Create the method row
  const methodRow = (
    <Row primary={false}>
      <RowName header={false} bold={true} name="Method" />
      <RowValue value={analysis.method} />
    </Row>
  )

  // Create the statistics section
  // If the statistics contain statistics, create StatisticsRows
  // Else create a one or more GroupRows
  let statisticsRows

  if (analysis.statistics) {
    statisticsRows = <StatisticsRows statistics={analysis.statistics} />
  }

  if (analysis.groups) {
    statisticsRows = (
      <>
        {analysis.groups.map((x) => {
          let group

          if (x.groups) {
            group = <GroupRows key={x.name} name={x.name} groups={x.groups} />
          }

          if (x.statistics) {
            group = (
              <GroupRows key={x.name} name={x.name} statistics={x.statistics} />
            )
          }

          return group
        })}
      </>
    )
  }

  // Combine the method and statistics section into a single element
  const content = (
    <>
      {methodRow}
      {statisticsRows}
    </>
  )

  // Create a collapsible element containing the identifier row and the content
  const collapsible = (
    <AnalysisDiv>
      <Collapsible
        primary={true}
        bold={false}
        name={analysis.identifier}
        content={content}
        open={false}
      />
    </AnalysisDiv>
  )

  return collapsible
}

export { AnalysisRows }
