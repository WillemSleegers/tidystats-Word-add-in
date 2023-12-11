import { Tidystats } from "../classes/Tidystats"
import { Collapsible } from "./Collapsible"
import { Groups } from "./Groups"
import { Statistics } from "./Statistics"
import { Row, RowName, RowValue } from "./Row"
import { Input, makeStyles } from "@fluentui/react-components"
import { useState } from "react"

const useStyles = makeStyles({
  search: {
    marginBottom: "1rem",
    width: "100%",
  },
})

type AnalysesProps = {
  tidystats: Tidystats
}

export const Analyses = (props: AnalysesProps) => {
  const { tidystats } = props

  const [search, setSearch] = useState("")

  const styles = useStyles()

  return (
    <>
      <h2>Analyses</h2>
      <Input
        className={styles.search}
        placeholder="Search..."
        spellCheck={false}
        onChange={(e) => setSearch(e.target.value)}
      ></Input>
      {tidystats.analyses
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
              {statistics && <Statistics data={statistics} />}
              {groups && <Groups data={groups} depth={0} />}
            </Collapsible>
          )
        })}
    </>
  )
}
