import { Collapsible } from "./Collapsible"
import { Group } from "../classes/Group"
import { Statistics } from "./Statistics"
import { insertTable } from "../functions/insertTable"

type GroupsProps = {
  data: Group[]
}

export const Groups = (props: GroupsProps) => {
  const { data } = props

  return (
    <>
      {data.map((x: Group) => {
        const statistics = x.statistics
        const groups = x.groups

        return (
          <Collapsible
            key={x.identifier}
            open={false}
            header={x.name}
            onInsertClick={() => insertTable(x)}
          >
            {statistics && <Statistics data={x.statistics!} />}
            {groups && <Groups data={groups} />}
          </Collapsible>
        )
      })}
    </>
  )
}
