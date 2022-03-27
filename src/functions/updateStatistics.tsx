import { Tidystats } from "../classes/Tidystats"

import { formatValue } from "../functions/formatValue"

const updateStatistics = async (tidystats: Tidystats) => {
  Word.run(async (context) => {
    console.log("Updating statistic")

    // Find all content control items in the document
    const contentControls = context.document.contentControls
    context.load(contentControls, "items")

    return context.sync().then(function () {
      const items = contentControls.items

      // Loop over the content controls items and update the statistics
      for (const item of items) {
        // Find the statistic
        const statistic = tidystats.findStatistic(item.tag)

        // Replace the statistic reported in the document with the new one, if there is one
        if (statistic) {
          // Check whether a lower or upper bound was reported
          const components = item.tag.split("$")

          let bound
          if (components[components.length - 1].match(/lower|upper/)) {
            bound = components.pop()
          }

          const value = formatValue(statistic, 2, bound as "lower" | "upper")

          // Insert text
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
