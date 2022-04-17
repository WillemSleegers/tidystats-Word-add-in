import { Group } from "../classes/Group"
import { Statistic } from "../classes/Statistic"

import { StatisticsRows } from "./StatisticsRows"
import { Collapsible } from "./Collapsible"

import { insertTable } from "../functions/insertTable"

type GroupRowsProps = {
  name: string
  statistics?: Statistic[]
  groups?: Group[]
}

const GroupRows = (props: GroupRowsProps) => {
  const { name, statistics, groups } = props

  let content

  if (statistics) {
    content = <StatisticsRows statistics={statistics} />
  }

  if (groups) {
    content = (
      <>
        {groups.map((x) => {
          let group

          group = (
            <GroupRows
              key={x.name}
              name={x.name}
              statistics={x.statistics}
              groups={x.groups}
            />
          )

          return group
        })}
      </>
    )
  }

  // Add addTable() function
  // TODO: Figure out when exactly to add this option
  let addTable = false
  if (Array.isArray(groups)) addTable = true

  const handleAddClick = () => {
    insertTable(name, groups)
  }

  return (
    <Collapsible
      primary={false}
      bold={true}
      name={name}
      content={content}
      handleAddClick={addTable ? handleAddClick : undefined}
      open={false}
    />
  )
}

export { GroupRows }
