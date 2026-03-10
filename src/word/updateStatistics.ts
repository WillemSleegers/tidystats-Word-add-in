import { Analysis } from "../types"
import { findStatistic } from "../utils/findStatistic"
import { formatValue } from "../utils/formatValue"

export const updateStatistics = async (analyses: Analysis[]) => {
  Word.run(async (context) => {
    const contentControls = context.document.contentControls
    context.load(contentControls, "items")

    return context.sync().then(function () {
      const items = contentControls.items

      for (const item of items) {
        const statistic = findStatistic(analyses, item.tag)

        if (statistic) {
          const components = item.tag.split("$")

          let value
          const suffix = components[components.length - 1]

          if (suffix === "level" && "level" in statistic) {
            value = (statistic.level * 100).toString()
          } else if (suffix.match(/lower|upper/)) {
            value = formatValue(statistic, 2, suffix as "lower" | "upper")
          } else {
            value = formatValue(statistic, 2)
          }

          item.insertText(value, Word.InsertLocation.replace)
        }
      }
    })
  }).catch(function (error) {
    console.log("Error: " + error)
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo))
    }
  })
}

