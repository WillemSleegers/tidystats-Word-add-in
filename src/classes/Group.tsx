import { Statistic, RangedStatistic } from "./Statistic"

class Group {
  identifier: string
  name?: string
  names?: { name: string }[]
  statistics?: Statistic[]
  groups?: Group[]

  constructor(identifier: string, data: Group) {
    if (data.name) {
      this.name = data.name
      this.identifier = identifier + "$" + this.name
    } else {
      this.names = data.names
      this.identifier =
        identifier + "$" + this.names![0].name + "-" + this.names![1].name
    }

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

export { Group }
