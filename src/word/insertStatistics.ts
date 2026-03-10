import { Statistic, RangedStatistic } from "../types"
import { formatValue } from "../utils/formatValue"

export const insertStatistic = async (
  statistic: Statistic | RangedStatistic,
  bound?: "lower" | "upper"
) => {
  const value = formatValue(statistic, 2, bound)
  const id = bound ? statistic.identifier + "$" + bound : statistic.identifier

  Word.run(async (context) => {
    const contentControl = context.document
      .getSelection()
      .insertContentControl()
    contentControl.tag = id
    contentControl.insertText(value, "End")

    return context.sync()
  }).catch(function (error) {
    console.log("Error: " + error)
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo))
    }
  })
}

export const insertStatistics = async (
  statistics: (Statistic | RangedStatistic)[]
) => {
  Word.run(async (context) => {
    // Embed degrees of freedom with the test statistic rather than listing separately
    let display = [...statistics]
    if (display.some((x) => x.name == "statistic")) {
      if (
        display.some(
          (x) =>
            ["t", "z", "χ²"].includes(x.symbol!) &&
            display.some((x) => x.name == "df")
        )
      ) {
        display = display.filter((x) => x.name != "df")
      } else if (
        display.some((x) => x.symbol == "F") &&
        display.some((x) => x.name == "df numerator") &&
        display.some((x) => x.name == "df denominator")
      ) {
        display = display.filter(
          (x) => !["df numerator", "df denominator"].includes(x.name)
        )
      }
    }

    const range = context.document.getSelection()

    display.forEach((x, i) => {
      if (i !== 0) {
        range.getRange().insertText(", ", Word.InsertLocation.end)
      }

      if ("level" in x) {
        // Create the confidence interval section
        const levelRange = range.getRange()
        const levelCC = levelRange.insertContentControl()
        levelCC.insertText(
          (x.level * 100).toString(),
          Word.InsertLocation.replace
        )
        levelCC.tag = x.identifier + "$level"
        levelRange.insertText("% " + x.interval + " [", Word.InsertLocation.end)

        const lowerRange = range.getRange()
        const lowerCC = lowerRange.insertContentControl()
        lowerCC.insertText(
          formatValue(x, 2, "lower"),
          Word.InsertLocation.replace
        )
        lowerCC.tag = x.identifier + "$lower"

        range.getRange().insertText(", ", Word.InsertLocation.end)

        const upperRange = range.getRange()
        const upperCC = upperRange.insertContentControl()
        upperCC.insertText(
          formatValue(x, 2, "upper"),
          Word.InsertLocation.replace
        )
        upperCC.tag = x.identifier + "$upper"

        range.getRange("End").insertText("]", "End")
      } else {
        // Create the test statistic section
        if (
          (["t", "z", "χ²"].includes(x.symbol!) &&
            statistics.find((s) => s.name === "df")) ||
          (x.symbol === "F" &&
            statistics.find((s) => s.name === "df numerator") &&
            statistics.find((s) => s.name === "df denominator"))
        ) {
          if (["t", "z", "χ²"].includes(x.symbol!)) {
            const name = range.getRange("End")
            name.insertText(x.symbol!, "End")
            name.font.italic = true

            const parenthesisLeft = range.getRange("End")
            parenthesisLeft.insertText("(", "End")
            parenthesisLeft.font.italic = false

            const df = statistics.find((s) => s.name === "df")
            if (df) {
              const dfValue = range.getRange("End")
              const dfValueCC = dfValue.insertContentControl()
              dfValueCC.insertText(
                formatValue(df, 2),
                Word.InsertLocation.replace
              )
              dfValueCC.tag = df.identifier
            }

            const parenthesisRight = range.getRange("End")
            parenthesisRight.insertText(")", "End")
          } else if (x.symbol === "F") {
            const name = range.getRange("End")
            name.insertText(x.symbol, "End")
            name.font.italic = true

            const parenthesisLeft = range.getRange("End")
            parenthesisLeft.insertText("(", "End")
            parenthesisLeft.font.italic = false

            const dfNum = statistics.find((s) => s.name === "df numerator")
            if (dfNum) {
              const dfValue = range.getRange("End")
              const dfValueCC = dfValue.insertContentControl()
              dfValueCC.insertText(
                formatValue(dfNum, 2),
                Word.InsertLocation.replace
              )
              dfValueCC.tag = dfNum.identifier
            }

            range.getRange().insertText(", ", "End")

            const dfDen = statistics.find((s) => s.name === "df denominator")
            if (dfDen) {
              const dfValue = range.getRange()
              const dfValueCC = dfValue.insertContentControl()
              dfValueCC.insertText(
                formatValue(dfDen, 2),
                Word.InsertLocation.replace
              )
              dfValueCC.tag = dfDen.identifier
            }

            range.getRange().insertText(")", "End")
          }
        } else {
          if (x.symbol != "%") {
            const name = range.getRange()
            name.insertText(x.symbol ? x.symbol : x.name, "End")
            name.font.italic = true

            if (x.subscript) {
              const subscript = range.getRange()
              subscript.insertText(x.subscript, "End")
              subscript.font.subscript = true
            }
          }
        }

        const value = formatValue(x, 2)

        if (x.symbol != "%") {
          const equal = range.getRange()
          if (value != "< .001") {
            equal.insertText(" = ", "End")
          } else {
            equal.insertText(" ", "End")
          }
          equal.font.italic = false
          equal.font.subscript = false
        }

        const valueCC = range.getRange().insertContentControl()
        valueCC.insertText(value, Word.InsertLocation.end)
        valueCC.tag = x.identifier

        if (x.symbol == "%") {
          const percentage = range.getRange()
          percentage.insertText("%", "End")
          percentage.font.italic = false
          percentage.font.subscript = false
        }
      }
    })

    return context.sync()
  }).catch(function (error) {
    console.log("Error: " + error)
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo))
    }
  })
}
