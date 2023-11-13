import { useEffect, useState } from "react"
import { makeStyles, Button, Checkbox } from "@fluentui/react-components"
import {
  Add24Filled,
  Add24Regular,
  Settings24Filled,
  Settings24Regular,
  bundleIcon,
} from "@fluentui/react-icons"
import { Row, RowName, RowValue } from "./Row"
import { RangedStatistic, Statistic } from "../classes/Statistic"
import { formatValue } from "../functions/formatValue"
import {
  insertStatistic,
  insertStatistics,
} from "../functions/insertStatistics"

const useStyles = makeStyles({
  statisticsWrapper: {
    marginLeft: "2rem",
  },
})

const AddIcon = bundleIcon(Add24Filled, Add24Regular)
const GearIcon = bundleIcon(Settings24Filled, Settings24Regular)

type StatisticsProps = {
  data: Statistic[] | RangedStatistic[]
}

type SelectedStatistic = {
  identifier: string
  name: string
  symbol?: string
  subscript?: string
  interval?: string
  level?: number
  value: string
  checked: boolean
}

export const Statistics = (props: StatisticsProps) => {
  const { data } = props

  const styles = useStyles()

  const [statistics, setStatistics] = useState<SelectedStatistic[]>()
  const [clickedSettings, setClickedSettings] = useState(false)

  useEffect(() => {
    const initialStatistics: SelectedStatistic[] = []

    data.forEach((x: Statistic | RangedStatistic) => {
      const selectedStatistic: SelectedStatistic = {
        identifier: x.identifier,
        name: x.name,
        symbol: x.symbol,
        subscript: x.subscript,
        value: formatValue(x, 2),
        checked: true,
      }

      initialStatistics.push(selectedStatistic)

      if ("level" in x) {
        const selectedStatisticLower = {
          identifier: x.identifier + "$lower",
          name: "LL",
          value: formatValue(x, 2, "lower"),
          level: x.level,
          interval: x.interval,
          checked: true,
        }
        const selectedStatisticUpper = {
          identifier: x.identifier + "$upper",
          name: "UL",
          value: formatValue(x, 2, "upper"),
          checked: true,
        }

        initialStatistics.push(selectedStatisticLower)
        initialStatistics.push(selectedStatisticUpper)
      }
    })

    setStatistics(initialStatistics)
  }, [])

  const toggleCheck = (name: string) => {
    setStatistics(
      statistics!.map((item) =>
        item.name === name ||
        (name === "LL" && item.name == "UL") ||
        (name === "UL" && item.name == "LL")
          ? { ...item, checked: !item.checked }
          : item
      )
    )
  }

  return (
    <>
      <Row indented hasBorder>
        <RowName isHeader isBold>
          Statistics:
        </RowName>
        {statistics && statistics.length > 1 && (
          <Button
            icon={<GearIcon />}
            appearance="transparent"
            onClick={() => setClickedSettings((prev) => !prev)}
          />
        )}
        <Button
          icon={<AddIcon />}
          appearance="transparent"
          onClick={() => insertStatistics(statistics!)}
        />
      </Row>
      <div className={styles.statisticsWrapper}>
        {statistics &&
          statistics.map((x) => {
            return (
              <Row
                key={x.identifier}
                hasBorder
                indented={x.name === "UL" || x.name === "LL"}
              >
                <RowName isHeader={false}>
                  {x.symbol ? x.symbol : x.name}
                  {x.subscript && <sub>{x.subscript}</sub>}
                </RowName>
                <RowValue>{x.value}</RowValue>
                {clickedSettings && (
                  <Checkbox
                    checked={x.checked}
                    onChange={() => toggleCheck(x.name)}
                  />
                )}
                <Button
                  icon={<AddIcon />}
                  appearance="transparent"
                  onClick={() => insertStatistic(x.value, x.identifier)}
                />
              </Row>
            )
          })}
      </div>
    </>
  )
}
