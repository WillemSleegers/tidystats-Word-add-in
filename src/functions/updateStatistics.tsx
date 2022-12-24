import { Tidystats } from "../classes/Tidystats"
import { formatValue } from "../functions/formatValue"

const updateStatistics = async (tidystats: Tidystats) => {
  Word.run(async (context) => {
    console.log("Updating statistic")

    const contentControls = context.document.contentControls
    context.load(contentControls, "items")

    return context.sync().then(function () {
      const items = contentControls.items

      for (const item of items) {
        const statistic = tidystats.findStatistic(item.tag)

        if (statistic) {
          const components = item.tag.split("$")

          let bound
          if (components[components.length - 1].match(/lower|upper/)) {
            bound = components.pop()
          }

          const value = formatValue(statistic, 2, bound as "lower" | "upper")

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

export { updateStatistics }
