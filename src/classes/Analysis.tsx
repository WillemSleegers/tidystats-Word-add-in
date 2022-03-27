import { Statistic, RangedStatistic } from "./Statistic"
import { Group } from "./Group"

class Analysis {
  identifier: string
  method: string
  statistics?: Statistic[]
  groups?: Group[]

  constructor(identifier: string, data: Analysis) {
    this.identifier = identifier
    this.method = data.method

    if (data.statistics) {
      const statistics = []

      for (let datum of data.statistics) {
        let statistic

        if ("level" in datum) {
          statistic = new RangedStatistic(
            this.identifier,
            datum as RangedStatistic
          )
        } else {
          statistic = new Statistic(this.identifier, datum)
        }

        statistics.push(statistic)
      }

      this.statistics = statistics
    }

    if (data.groups) {
      const groups = []

      for (let datum of data.groups) {
        const group = new Group(this.identifier, datum)
        groups.push(group)
      }

      this.groups = groups
    }
  }
}

export { Analysis }
