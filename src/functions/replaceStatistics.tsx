export const replaceStatistics = async (x: string) => {
  Word.run(async (context) => {
    console.log("Replacing statistic")

    const contentControls = context.document.contentControls
    context.load(contentControls, "items")

    return context.sync().then(function () {
      const items = contentControls.items

      for (const item of items) {
        item.insertText(x, Word.InsertLocation.replace)
        item.font.highlightColor = "yellow"
      }
    })
  }).catch(function (error) {
    console.log("Error: " + error)
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo))
    }
  })
}
