import { Analysis, Group, Statistic, RangedStatistic } from "../types"

function createStatistic(
  identifier: string,
  data: Record<string, unknown>
): Statistic | RangedStatistic {
  const base: Statistic = {
    identifier: identifier + "$" + (data.name as string),
    name: data.name as string,
    symbol: data.symbol as string | undefined,
    subscript: data.subscript as string | undefined,
    value: data.value as number,
  }

  if ("level" in data) {
    return {
      ...base,
      interval: data.interval as string,
      level: data.level as number,
      lower: data.lower as number | string,
      upper: data.upper as number | string,
    }
  }

  return base
}

function createGroup(
  identifier: string,
  data: Record<string, unknown>
): Group {
  let groupIdentifier: string
  const group: Group = { identifier: "" }

  if ("name" in data) {
    group.name = data.name as string
    groupIdentifier = identifier + "$" + group.name
  } else {
    group.names = data.names as { name: string }[]
    groupIdentifier =
      identifier + "$" + group.names![0].name + "-" + group.names![1].name
  }

  group.identifier = groupIdentifier

  if (Array.isArray(data.statistics)) {
    group.statistics = data.statistics.map((datum: Record<string, unknown>) =>
      createStatistic(groupIdentifier, datum)
    )
  }

  if (Array.isArray(data.groups)) {
    group.groups = data.groups.map((datum: Record<string, unknown>) =>
      createGroup(groupIdentifier, datum)
    )
  }

  return group
}

function createAnalysis(
  identifier: string,
  data: Record<string, unknown>
): Analysis {
  const analysis: Analysis = {
    identifier,
    method: data.method as string,
  }

  if (Array.isArray(data.statistics)) {
    analysis.statistics = data.statistics.map((datum: Record<string, unknown>) =>
      createStatistic(identifier, datum)
    )
  }

  if (Array.isArray(data.groups)) {
    analysis.groups = data.groups.map((datum: Record<string, unknown>) =>
      createGroup(identifier, datum)
    )
  }

  return analysis
}

export function parseAnalyses(data: Record<string, unknown>): Analysis[] {
  return Object.keys(data).map((key) =>
    createAnalysis(key, data[key] as Record<string, unknown>)
  )
}

