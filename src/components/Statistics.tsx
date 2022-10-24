import { useEffect, useState } from "react"
import { Checkbox, IIconProps } from "@fluentui/react"
import { IconButton } from "@fluentui/react/lib/Button"
import { Row, RowName, RowValue } from "./Row"
import { Statistic, RangedStatistic } from "../classes/Statistic"
import { formatValue } from "../functions/formatValue"
import { insertStatistic } from "../functions/insertStatistic"
import { insertStatistics } from "../functions/insertStatistics"

const gearIcon: IIconProps = { iconName: "Settings" }
const addIcon: IIconProps = { iconName: "Add" }

type StatisticsProps = {
  data: Statistic[]
}

type itemProps = {
  name: string
  identifier: string
  symbol?: string
  subscript?: string
  value: string
  checked: boolean
}

export const Statistics = (props: StatisticsProps) => {
  const { data } = props

  const [items, setItems] = useState<Array<itemProps>>([])
  const [clickedSettings, setClickedSettings] = useState(false)

  const toggleSettings = () => {
    setClickedSettings((prev) => !prev)
  }

  useEffect(() => {
    const initialItems: itemProps[] = []

    data.forEach((x) => {
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
          //symbol: y.level * 100 + "% " + y.interval,
          symbol: y.symbol,
          subscript: "lower",
          value: formatValue(y, 2, "lower"),
          checked: true,
        }
        const item_upper = {
          identifier: y.identifier + "$upper",
          name: "upper",
          //symbol: y.level * 100 + "% " + y.interval,
          symbol: y.symbol,
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
  }, [data])

  const handleAddClick = () => {
    console.log("Inserting statistic")
  }

  const toggleCheck = (name: string) => {
    const newItems = items.map((item) =>
      item.name === name ? { ...item, checked: !item.checked } : item
    )
    setItems(newItems)
  }

  return (
    <>
      <Row indentationLevel={1} hasBorder={true}>
        <RowName isHeader={true} isBold={true}>
          Statistics:
        </RowName>

        <IconButton iconProps={gearIcon} onClick={toggleSettings} />
        <IconButton iconProps={addIcon} onClick={() => console.log("test")} />
      </Row>
      {items.map((x: itemProps, index: number) => {
        const lastRow = index === items.length - 1
        return (
          <Row
            key={x.identifier}
            // indent more if the statistic is an upper or lower bound
            indentationLevel={
              x.subscript === "upper" || x.subscript === "lower" ? 3 : 2
            }
            hasBorder={!lastRow}
          >
            <RowName isHeader={false}>
              {x.symbol}
              {x.subscript && <sub>{x.subscript}</sub>}
            </RowName>
            <RowValue>{x.value}</RowValue>
            {clickedSettings && (
              <Checkbox
                styles={{
                  root: {
                    marginRight: "2px",
                    "&:hover .ms-Checkbox-checkbox": {
                      borderColor: "rgb(16, 110, 190)",
                    },
                  },
                  checkbox: {
                    borderColor: "rgb(16, 110, 190)",
                    [":hover"]: {
                      borderColor: "red",
                    },
                  },
                }}
                checked={x.checked}
                onChange={() => toggleCheck(x.name)}
              />
            )}
            <IconButton
              iconProps={addIcon}
              onClick={() => insertStatistic(x.value, x.identifier)}
            />
          </Row>
        )
      })}
    </>
  )
}
