import { useState, useEffect } from "react"
import styled from "styled-components"

import { Statistic, RangedStatistic } from "../classes/Statistic"
import { Row } from "./Row"
import { RowName } from "./RowName"
import { RowValue } from "./RowValue"

import { formatValue } from "../functions/formatValue"
import { insertStatistic } from "../functions/insertStatistic"
import { insertStatistics } from "../functions/insertStatistics"

import { Collapsible } from "./Collapsible"

import { Checkbox, IIconProps } from "@fluentui/react"
import { IconButton } from "@fluentui/react/lib/Button"

const addIcon: IIconProps = { iconName: "Add" }

// Create a new checkbox that aligns with the settings icon
const MyCheckbox = styled(Checkbox)`
  width: 28px;
  justify-content: center;
`

type StatisticsRowsProps = {
  statistics: Statistic[]
}

type itemsProps = {
  name: string
  identifier: string
  symbol: string
  subscript?: string
  value: string
  checked: boolean
}

const StatisticsRows = (props: StatisticsRowsProps) => {
  const { statistics } = props

  const [clickedSettings, setClickedSettings] = useState(false)
  const [items, setItems] = useState<Array<itemsProps>>([])

  useEffect(() => {
    const initialItems: itemsProps[] = []

    statistics.forEach((x) => {
      if ("level" in x) {
        const y = x as RangedStatistic

        const item = {
          identifier: y.identifier,
          name: y.name,
          symbol: y.symbol !== undefined ? y.symbol : y.name,
          subscript: y.subscript,
          value: formatValue(y, 2),
          checked: true,
        }
        const item_lower = {
          identifier: y.identifier + "$lower",
          name: "lower",
          symbol: y.level * 100 + "% " + y.interval,
          subscript: "lower",
          value: formatValue(y, 2, "lower"),
          checked: true,
        }
        const item_upper = {
          identifier: y.identifier + "$upper",
          name: "upper",
          symbol: y.level * 100 + "% " + y.interval,
          subscript: "upper",
          value: formatValue(y, 2, "upper"),
          checked: true,
        }
        initialItems.push(item)
        initialItems.push(item_lower)
        initialItems.push(item_upper)
      } else {
        const item = {
          identifier: x.identifier,
          name: x.name,
          symbol: x.symbol !== undefined ? x.symbol : x.name,
          subscript: x.subscript,
          value: formatValue(x, 2),
          checked: true,
        }
        initialItems.push(item)
      }
    })

    setItems(initialItems)
  }, [statistics])

  const handleCheck = (name: string) => {
    const newItems = items.map((item) =>
      item.name === name ? { ...item, checked: !item.checked } : item
    )
    setItems(newItems)
  }

  const handleSettingsClick = () => {
    setClickedSettings((prev) => !prev)
  }

  const handleAddClick = () => {
    insertStatistics(items)
  }

  const content = items.map((x) => {
    return (
      <Row primary={false} key={x.identifier}>
        <RowName
          key={x.name}
          header={false}
          bold={false}
          name={x.symbol}
          subscript={x.subscript}
        />
        <RowValue key={x.value} value={x.value} />
        <>
          {clickedSettings && (
            <MyCheckbox
              checked={x.checked}
              onChange={() => handleCheck(x.name)}
            />
          )}
        </>
        <IconButton
          iconProps={addIcon}
          onClick={() => insertStatistic(x.value, x.identifier)}
        />
      </Row>
    )
  })

  return (
    <Collapsible
      primary={false}
      bold={true}
      name="Statistics:"
      handleSettingsClick={handleSettingsClick}
      handleAddClick={handleAddClick}
      content={content}
      open={true}
      disabled={true}
    />
  )
}
export { StatisticsRows }
