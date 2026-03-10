import { Analysis, Statistic, RangedStatistic } from "../types"

export function findStatistic(
  analyses: Analysis[],
  id: string
): Statistic | RangedStatistic | undefined {
  const components = id.split("$")

  if (components[components.length - 1].match(/lower|upper|level/)) {
    components.pop()
  }

  const identifier = components[0]
  const statisticName = components[components.length - 1]
  const groupNames = components.slice(1, components.length - 1)

  const analysis = analyses.find((x) => x.identifier === identifier)

  let statistics

  if (groupNames.length) {
    let groups, group

    groups = analysis?.groups

    for (let i = 0; i < groupNames.length; i++) {
      group = groups?.find((x) => x.name === groupNames[i])
      groups = group?.groups
    }

    statistics = group?.statistics
  } else {
    statistics = analysis?.statistics
  }

  return statistics?.find((x) => x.name === statisticName)
}
