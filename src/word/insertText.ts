export const insertText = async (x: string) => {
  Word.run(async (context) => {
    const range = context.document.getSelection()
    range.insertText(x, "End")

    return context.sync
  }).catch(function (error) {
    console.log("Error: " + error)
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo))
    }
  })
}
