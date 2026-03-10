import { Fragment, useState } from "react"
import { makeStyles, Button, Checkbox } from "@fluentui/react-components"
import {
  Add24Filled,
  Add24Regular,
  Settings24Filled,
  Settings24Regular,
  bundleIcon,
} from "@fluentui/react-icons"
import { Row, RowName, RowValue } from "./Row"
import { RangedStatistic, Statistic } from "../../types"
import { formatValue } from "../../utils/formatValue"
import { filterStatistics } from "../../utils/filterStatistics"
import {
  insertStatistic,
  insertStatistics,
} from "../../word/insertStatistics"

const useStyles = makeStyles({
  statisticsWrapper: {
    marginLeft: "2rem",
  },
})

const AddIcon = bundleIcon(Add24Filled, Add24Regular)
const GearIcon = bundleIcon(Settings24Filled, Settings24Regular)

type StatisticRowsProps = {
  data: (Statistic | RangedStatistic)[]
}

export const StatisticRows = (props: StatisticRowsProps) => {
  const { data } = props

  const styles = useStyles()

  const [checkedIds, setCheckedIds] = useState<Set<string>>(
    () => new Set(data.map((x) => x.identifier))
  )
  const [clickedSettings, setClickedSettings] = useState(false)

  const toggleCheck = (id: string) => {
    setCheckedIds((prev) => {
      const next = new Set(prev)
      if (next.has(id)) {
        next.delete(id)
      } else {
        next.add(id)
      }
      return next
    })
  }

  return (
    <>
      <Row indented hasBorder>
        <RowName isHeader isBold>
          Statistics:
        </RowName>
        {data.length > 1 && (
          <Button
            icon={<GearIcon />}
            appearance="transparent"
            onClick={() => setClickedSettings((prev) => !prev)}
          />
        )}
        <Button
          icon={<AddIcon />}
          appearance="transparent"
          onClick={() => insertStatistics(filterStatistics(data, checkedIds))}
        />
      </Row>
      <div className={styles.statisticsWrapper}>
        {data.map((x) => (
          <Fragment key={x.identifier}>
            <Row hasBorder>
              <RowName isHeader={false}>
                {x.symbol ? x.symbol : x.name}
                {x.subscript && <sub>{x.subscript}</sub>}
              </RowName>
              <RowValue>{formatValue(x, 2)}</RowValue>
              {clickedSettings && (
                <Checkbox
                  checked={checkedIds.has(x.identifier)}
                  onChange={() => toggleCheck(x.identifier)}
                />
              )}
              <Button
                icon={<AddIcon />}
                appearance="transparent"
                onClick={() => insertStatistic(x)}
              />
            </Row>
            {"level" in x && (
              <>
                <Row hasBorder indented>
                  <RowName isHeader={false}>LL</RowName>
                  <RowValue>{formatValue(x, 2, "lower")}</RowValue>
                  <Button
                    icon={<AddIcon />}
                    appearance="transparent"
                    onClick={() => insertStatistic(x, "lower")}
                  />
                </Row>
                <Row hasBorder indented>
                  <RowName isHeader={false}>UL</RowName>
                  <RowValue>{formatValue(x, 2, "upper")}</RowValue>
                  <Button
                    icon={<AddIcon />}
                    appearance="transparent"
                    onClick={() => insertStatistic(x, "upper")}
                  />
                </Row>
              </>
            )}
          </Fragment>
        ))}
      </div>
    </>
  )
}
