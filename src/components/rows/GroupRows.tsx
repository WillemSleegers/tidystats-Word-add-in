import { Collapsible } from "./Collapsible"
import { Group } from "../../types"
import { StatisticRows } from "./StatisticRows"
import { formatName } from "../../utils/formatName"
import { insertTable } from "../../word/insertTable"

type GroupRowsProps = {
  data: Group[]
  depth: number
}

export const GroupRows = (props: GroupRowsProps) => {
  const { data, depth } = props

  return (
    <>
      {data.map((x: Group) => {
        const statistics = x.statistics
        const groups = x.groups

        let canInsertTable = false
        const open = data.length == 1

        if (groups) {
          canInsertTable =
            groups.filter((group) => group.statistics).length ==
              groups.length && groups.length > 1
        }

        return (
          <Collapsible
            key={x.identifier}
            open={open}
            header={formatName(x)}
            indentation={depth}
            onInsertClick={canInsertTable ? () => insertTable(x) : undefined}
          >
            {groups && <GroupRows data={groups} depth={depth + 1} />}
            {statistics && <StatisticRows data={x.statistics!} />}
          </Collapsible>
        )
      })}
    </>
  )
}
