export interface Statistic {
  identifier: string
  name: string
  symbol?: string
  subscript?: string
  value: number
}

export interface RangedStatistic extends Statistic {
  interval: string
  level: number
  lower: number | string
  upper: number | string
}

export interface Group {
  identifier: string
  name?: string
  names?: { name: string }[]
  statistics?: (Statistic | RangedStatistic)[]
  groups?: Group[]
}

export interface Analysis {
  identifier: string
  method: string
  statistics?: (Statistic | RangedStatistic)[]
  groups?: Group[]
}
