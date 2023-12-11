import { RangedStatistic, Statistic } from "../classes/Statistic"
import { Group } from "../classes/Group"
import { formatName } from "../functions/formatName"
import { formatValue } from "../functions/formatValue"

export const insertTable = async (data: Group) => {
  const groups = data.groups!

  const rows = groups.length + 1
  const columns =
    Math.max(...groups.map((group) => countStatistics(group.statistics!))) + 1

  // Get the index of the group with the most statistics and its statistics
  // names in the header row
  const index = groups.findIndex(
    (group) => countStatistics(group.statistics!) == columns - 1
  )

  Word.run(async (context) => {
    const range = context.document.getSelection()
    const table = range.insertTable(rows, columns, Word.InsertLocation.after)

    table.getBorder("All").type = "None"
    table.getBorder("Top").type = "Single"
    table.getBorder("Bottom").type = "Single"

    // Set the first cell's content to the group name
    const cellName = table.getCell(0, 0)
    cellName.getBorder("Bottom").type = "Single"
    cellName.body
      .getRange()
      .insertText(formatName(data)!, Word.InsertLocation.replace)

    // Set the content of the remaining cells in the first row to the names of
    // the statistics
    let intervalsCount = 0
    groups[index].statistics!.forEach(
      (statistic: Statistic | RangedStatistic, i) => {
        const cell = table.getCell(0, i + 1 + intervalsCount)

        cell.getBorder("Bottom").type = "Single"

        const range = cell.body.getRange("End")

        range.insertText(
          statistic.symbol ? statistic.symbol : statistic.name,
          Word.InsertLocation.replace
        )

        range.font.italic = true

        if (statistic.subscript) {
          const subscriptRange = cell.body.getRange("End")
          subscriptRange.insertText(
            statistic.subscript,
            Word.InsertLocation.end
          )
          subscriptRange.font.subscript = true
        }

        if ("level" in statistic) {
          intervalsCount++

          const cell = table.getCell(0, i + 1 + intervalsCount)

          cell.getBorder("Bottom").type = "Single"

          const range = cell.body.getRange("End")
          const contentControl = range.insertContentControl()

          contentControl.insertText(
            (statistic.level * 100).toString(),
            Word.InsertLocation.replace
          )
          contentControl.tag = statistic.identifier + "$level"

          range.insertText("% " + statistic.interval, Word.InsertLocation.end)
        }
      }
    )

    // Loop over each group and set the name and values

    groups.forEach((group, i) => {
      intervalsCount = 0
      table
        .getCell(i + 1, 0)
        .body.getRange()
        .insertText(formatName(group)!, Word.InsertLocation.replace)

      group.statistics?.forEach((statistic: Statistic | RangedStatistic, j) => {
        const value = table
          .getCell(i + 1, j + 1 + intervalsCount)
          .body.getRange()
          .insertContentControl()
        value.tag = statistic.identifier
        value.insertText(formatValue(statistic, 2), Word.InsertLocation.replace)

        if ("level" in statistic) {
          intervalsCount++
          const range = table
            .getCell(i + 1, j + 1 + intervalsCount)
            .body.getRange()

          range.insertText(" [", Word.InsertLocation.start)
          const lowerRange = range.getRange(Word.RangeLocation.end)
          const lowerCC = lowerRange.insertContentControl()
          lowerCC.insertText(
            formatValue(statistic, 2, "lower"),
            Word.InsertLocation.replace
          )
          lowerCC.tag = statistic.identifier + "$lower"

          const commaRange = range.getRange(Word.RangeLocation.end)
          commaRange.insertText(", ", Word.InsertLocation.end)

          const upperRange = range.getRange(Word.RangeLocation.end)
          const upperCC = upperRange.insertContentControl()
          upperCC.insertText(
            formatValue(statistic, 2, "upper"),
            Word.InsertLocation.replace
          )
          upperCC.tag = statistic.identifier + "$upper"

          const rightBracketRange = range.getRange(Word.RangeLocation.end)
          rightBracketRange.insertText("]", Word.InsertLocation.end)
        }
      })
    })

    return context.sync
  }).catch(function (error) {
    console.log("Error: " + error)
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo))
    }
  })
}

const countStatistics = (data: Statistic[] | RangedStatistic[]) => {
  return data.map((x) => ("level" in x ? 2 : 1)).reduce((a, b) => a + b, 0)
}

// const getStatisticsNames = (data: Statistic[] | RangedStatistic[]) => {
//   return data.map((x) => ("level" in x ? 3 : 1)).reduce((a, b) => a + b, 0)
// }
