class Statistic {
  name: string
  identifier: string
  symbol?: string
  subscript?: string
  value: number

  constructor(identifier: string, data: Statistic) {
    this.name = data.name
    this.identifier = identifier + "$" + this.name
    this.symbol = data.symbol
    this.subscript = data.subscript
    this.value = data.value
  }
}

class RangedStatistic extends Statistic {
  interval: string
  level: number
  lower: number | string
  upper: number | string

  constructor(identifier: string, data: RangedStatistic) {
    super(identifier, data)

    this.interval = data.interval
    this.level = data.level
    this.lower = data.lower
    this.upper = data.upper
  }
}

export { Statistic, RangedStatistic }
