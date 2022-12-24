export const insertStatistic = async (statistic: string, id: string) => {
  Word.run(async (context) => {
    const contentControl = context.document
      .getSelection()
      .insertContentControl()
    contentControl.tag = id
    contentControl.insertText(statistic, "End")

    return context.sync
  }).catch(function (error) {
    console.log("Error: " + error)
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo))
    }
  })
}

type StatisticProps = {
  identifier: string
  name: string
  symbol?: string
  subscript?: string
  interval?: string
  level?: number
  value: string
  checked: boolean
}

export const insertStatistics = async (statistics: StatisticProps[]) => {
  Word.run(async (context) => {
    let selectedStatistics

    // Filter out the unchecked statistics
    selectedStatistics = statistics.filter(
      (statistic: StatisticProps) => statistic.checked
    )

    // Filter out the degrees of freedom if there's a test statistic
    // (e.g., t, F) because we will report those together with the test
    // statistic itself
    if (
      selectedStatistics.some(
        (statistic: StatisticProps) => statistic.name === "statistic"
      )
    ) {
      selectedStatistics = selectedStatistics.filter(
        (statistic: StatisticProps) =>
          !["df", "df numerator", "df denominator"].includes(statistic.name)
      )
    }

    // If both the lower and upper bound of an interval are present, remove one,
    // because we'll report it together with the other one
    const lower = statistics.find((x: StatisticProps) => x.name === "LL")
    const upper = statistics.find((x: StatisticProps) => x.name === "UL")

    if (lower && upper) {
      selectedStatistics = selectedStatistics.filter(
        (statistic: StatisticProps) => statistic.name !== "UL"
      )
    }

    // Loop over the statistics and insert them
    const range = context.document.getSelection()

    selectedStatistics.forEach((statistic: StatisticProps, i: number) => {
      // Add a comma starting after the first element
      if (i !== 0) {
        const comma = range.getRange()
        comma.insertText(", ", "End")
      }

      // Create the confidence interval section
      if (statistic.name === "LL" && lower && upper) {
        const interval = range.getRange()

        const text = statistic.level! * 100 + "% " + statistic.interval!
        interval.insertText(text, Word.InsertLocation.end)
        interval.insertText(" [", "End")

        const lowerRange = range.getRange()
        const lowerCC = lowerRange.insertContentControl()
        lowerCC.insertText(lower.value, Word.InsertLocation.start)
        lowerCC.tag = lower.identifier

        const intervalComma = range.getRange()
        intervalComma.insertText(", ", "End")

        const upperRange = range.getRange()
        const upperCC = upperRange.insertContentControl()
        upperCC.insertText(upper.value, Word.InsertLocation.start)
        upperCC.tag = upper.identifier

        const rightBracket = range.getRange("End")
        rightBracket.insertText("]", "End")
      } else {
        // Create the test statistic section
        if (statistic.name === "statistic") {
          if (
            statistic.symbol === "t" &&
            statistics.find((x: StatisticProps) => x.name === "df")
          ) {
            const name = range.getRange("End")
            name.insertText(statistic.symbol, "End")
            name.font.italic = true

            const parenthesisLeft = range.getRange("End")
            parenthesisLeft.insertText("(", "End")
            parenthesisLeft.font.italic = false

            const df = statistics.find((x: StatisticProps) => x.name === "df")
            if (df) {
              const dfValue = range.getRange("End")
              const dfValueCC = dfValue.insertContentControl()
              dfValueCC.insertText(df.value, Word.InsertLocation.replace)
              dfValueCC.tag = df.identifier
            }

            const parenthesisRight = range.getRange("End")
            parenthesisRight.insertText(")", "End")
          } else if (
            statistic.symbol === "F" &&
            statistics.find((x: StatisticProps) => x.name === "df numerator") &&
            statistics.find((x: StatisticProps) => x.name === "df denominator")
          ) {
            const name = range.getRange("End")
            name.insertText(statistic.symbol, "End")
            name.font.italic = true

            const parenthesisLeft = range.getRange("End")
            parenthesisLeft.insertText("(", "End")
            parenthesisLeft.font.italic = false

            const dfNum = statistics.find(
              (x: StatisticProps) => x.name === "df numerator"
            )
            if (dfNum) {
              const dfValue = range.getRange("End")
              const dfValueCC = dfValue.insertContentControl()
              dfValueCC.insertText(dfNum.value, Word.InsertLocation.replace)
              dfValueCC.tag = dfNum.identifier
            }

            const dfComma = range.getRange()
            dfComma.insertText(", ", "End")

            const dfDen = statistics.find(
              (x: StatisticProps) => x.name === "df denominator"
            )
            if (dfDen) {
              const dfValue = range.getRange()
              const dfValueCC = dfValue.insertContentControl()
              dfValueCC.insertText(dfDen.value, Word.InsertLocation.replace)
              dfValueCC.tag = dfDen.identifier
            }

            const parenthesisRight = range.getRange()
            parenthesisRight.insertText(")", "End")
          }
        } else {
          const name = range.getRange()
          name.insertText(
            statistic.symbol ? statistic.symbol : statistic.name,
            "End"
          )
          name.font.italic = true

          if (statistic.subscript) {
            const subscript = range.getRange()
            subscript.insertText(statistic.subscript, "End")
            subscript.font.subscript = true
          }
        }

        // Insert an equal sign and set the style back to normal
        const equal = range.getRange()
        equal.insertText(" = ", "End")
        equal.font.italic = false
        equal.font.subscript = false

        // Insert the value as a content control
        const value = range.getRange()
        const valueCC = value.insertContentControl()
        valueCC.insertText(statistic.value, Word.InsertLocation.end)
        valueCC.tag = statistic.identifier
      }
    })

    return context.sync
  }).catch(function (error) {
    console.log("Error: " + error)
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo))
    }
  })
}
