const clearSettingsData = () => {
  Office.context.document.settings.set("fileName", "")
  Office.context.document.settings.set("data", "")

  Office.context.document.settings.saveAsync(function (asyncResult) {
    console.log("Removed saved data with status: " + asyncResult.status)
  })
}

export { clearSettingsData }
