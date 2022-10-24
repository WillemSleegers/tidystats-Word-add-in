import { Collapsible } from "./Collapsible"
import { Group } from "../classes/Group"
import { Statistics } from "./Statistics"

type GroupsProps = {
  data: Group[]
}

export const Groups = (props: GroupsProps) => {
  const { data } = props

  const handleAddClick = () => {
    console.log("inserting table")
  }

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
            onInsertClick={handleAddClick}
          >
            {statistics && <Statistics data={x.statistics!} />}
            {groups && <Groups data={groups} />}
          </Collapsible>
        )
      })}
    </>
  )
}
