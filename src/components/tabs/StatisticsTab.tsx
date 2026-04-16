import { Analysis } from "../../types"
import { Collapsible } from "../rows/Collapsible"
import { GroupRows } from "../rows/GroupRows"
import { StatisticRows } from "../rows/StatisticRows"
import { Row, RowName, RowValue } from "../rows/Row"
import { Input, makeStyles } from "@fluentui/react-components"
import { useState } from "react"
import { Upload } from "../Upload"

const useStyles = makeStyles({
  search: {
    marginTop: "1rem",
    marginBottom: "1rem",
    width: "100%",
  },
})

type StatisticsTabProps = {
  analyses: Analysis[] | undefined
  setAnalyses: (analyses: Analysis[] | undefined) => void
}

export const StatisticsTab = (props: StatisticsTabProps) => {
  const { analyses, setAnalyses } = props

  const [search, setSearch] = useState("")

  const styles = useStyles()

  return (
    <>
      <Upload setAnalyses={setAnalyses} />
      {analyses && (
        <>
          <Input
            className={styles.search}
            placeholder="Search..."
            spellCheck={false}
            onChange={(e) => setSearch(e.target.value)}
          />
          {analyses
            .filter((x) => x.identifier.includes(search))
            .map((x) => {
              const statistics = x.statistics
              const groups = x.groups

              return (
                <Collapsible
                  key={x.identifier}
                  open={false}
                  header={x.identifier}
                  indentation={0}
                  isPrimary
                >
                  <Row indented>
                    <RowName isHeader={false} isBold>
                      Method:
                    </RowName>
                    <RowValue>{x.method}</RowValue>
                  </Row>
                  {statistics && <StatisticRows data={statistics} />}
                  {groups && <GroupRows data={groups} depth={0} />}
                </Collapsible>
              )
            })}
        </>
      )}
    </>
  )
}
