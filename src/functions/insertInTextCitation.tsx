const insertInTextCitation = async () => {
  Word.run(async (context) => {
    const range = context.document.getSelection()
    range.insertText("Sleegers (2021)", "End")

    return context.sync
  }).catch(function (error) {
    console.log("Error: " + error)
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo))
    }
  })
}

export { insertInTextCitation }
