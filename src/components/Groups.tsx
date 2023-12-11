import { Collapsible } from "./Collapsible"
import { Group } from "../classes/Group"
import { Statistics } from "./Statistics"
import { formatName } from "../functions/formatName"
import { insertTable } from "../functions/insertTable"

type GroupsProps = {
  data: Group[]
  depth: number
}

export const Groups = (props: GroupsProps) => {
  const { data, depth } = props

  return (
    <>
      {data.map((x: Group) => {
        const statistics = x.statistics
        const groups = x.groups

        let canInsertTable = false
        let open = data.length == 1

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
            {groups && <Groups data={groups} depth={depth + 1} />}
            {statistics && <Statistics data={x.statistics!} />}
          </Collapsible>
        )
      })}
    </>
  )
}
