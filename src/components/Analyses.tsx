import { Tidystats } from "../classes/Tidystats"
import { Collapsible } from "./Collapsible"
import { Groups } from "./Groups"
import { Statistics } from "./Statistics"
import { Row, RowName } from "./Row"

type AnalysesProps = {
  tidystats: Tidystats
}

export const Analyses = (props: AnalysesProps) => {
  const { tidystats } = props

  return (
    <>
      <h2>Analyses</h2>
      {tidystats.analyses.map((x) => {
        const statistics = x.statistics
        const groups = x.groups

        return (
          <Collapsible
            key={x.identifier}
            open={false}
            header={x.identifier}
            isPrimary
          >
            <Row indented hasBorder>
              <RowName isHeader={false} isBold>
                Method:
              </RowName>
              <div>{x.method}</div>
            </Row>
            {statistics && <Statistics data={statistics} />}
            {groups && <Groups data={groups} />}
          </Collapsible>
        )
      })}
    </>
  )
}
