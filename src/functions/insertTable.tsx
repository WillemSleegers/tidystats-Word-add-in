import { RangedStatistic, Statistic } from "../classes/Statistic"
import { Group } from "../classes/Group"
import { formatName } from "../functions/formatName"
import { formatValue } from "../functions/formatValue"

export const insertTable = async (data: Group) => {
  const groups = data.groups!

  const rows = groups.length
  const columns = Math.max(
    ...groups.map((group) => getStatisticsLength(group.statistics!))
  )
  console.log(columns)
  const index = groups.findIndex((group) => group.statistics!.length == columns)

  Word.run(async (context) => {
    const range = context.document.getSelection()
    const table = range.insertTable(
      rows + 1,
      columns + 1,
      Word.InsertLocation.after
    )

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
    groups[index].statistics!.forEach((statistic, i) => {
      const cell = table.getCell(0, i + 1)

      cell.getBorder("Bottom").type = "Single"

      const range = cell.body.getRange("End")

      range.insertText(
        statistic.symbol ? statistic.symbol : statistic.name,
        Word.InsertLocation.replace
      )

      range.font.italic = true

      if (statistic.subscript) {
        const subscriptRange = cell.body.getRange("End")
        subscriptRange.insertText(statistic.subscript, Word.InsertLocation.end)
        subscriptRange.font.subscript = true
      }
    })

    // Loop over each group and set the name and values
    groups.forEach((group, i) => {
      table
        .getCell(i + 1, 0)
        .body.getRange()
        .insertText(formatName(group)!, Word.InsertLocation.replace)

      group.statistics?.forEach((statistic, j) => {
        const value = table
          .getCell(i + 1, j + 1)
          .body.getRange()
          .insertContentControl()
        value.tag = statistic.identifier
        value.insertText(formatValue(statistic, 2), Word.InsertLocation.replace)
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

const getStatisticsLength = (data: Statistic[] | RangedStatistic[]) => {
  return data.map((x) => ("level" in x ? 3 : 1)).reduce((a, b) => a + b, 0)
}

const getStatisticsNames = (data: Statistic[] | RangedStatistic[]) => {
  return data.map((x) => ("level" in x ? 3 : 1)).reduce((a, b) => a + b, 0)
}
