import { Statistic, RangedStatistic } from "../types"

export function filterStatistics(
  statistics: (Statistic | RangedStatistic)[],
  checkedIds: Set<string>
): (Statistic | RangedStatistic)[] {
  return statistics.filter((x) => checkedIds.has(x.identifier))
}
