const clearSettingsData = () => {
  Office.context.document.settings.set("fileName", null)
  Office.context.document.settings.set("data", null)

  Office.context.document.settings.saveAsync(function (asyncResult) {
    console.log("Removed saved data with status: " + asyncResult.status)
  })
}

export { clearSettingsData }
