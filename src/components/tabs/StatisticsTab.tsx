import { Analysis } from "../../types"
import { Collapsible } from "../rows/Collapsible"
import { GroupRows } from "../rows/GroupRows"
import { StatisticRows } from "../rows/StatisticRows"
import { Row, RowName, RowValue } from "../rows/Row"
import { Input, makeStyles } from "@fluentui/react-components"
import { useState } from "react"

const useStyles = makeStyles({
  search: {
    marginBottom: "1rem",
    width: "100%",
  },
})

type StatisticsTabProps = {
  analyses: Analysis[]
}

export const StatisticsTab = (props: StatisticsTabProps) => {
  const { analyses } = props

  const [search, setSearch] = useState("")

  const styles = useStyles()

  return (
    <>
      <h2>Statistics</h2>
      <Input
        className={styles.search}
        placeholder="Search..."
        spellCheck={false}
        onChange={(e) => setSearch(e.target.value)}
      ></Input>
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
  )
}
