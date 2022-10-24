const insertStatistic = async (statistic: string, id: string) => {
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

export { insertStatistic }
