export const setSettingsData = (name: string, value: string | boolean) => {
  Office.context.document.settings.set(name, value)

  Office.context.document.settings.saveAsync(function (asyncResult) {
    console.log("Settings saved with status: " + asyncResult.status)
  })
}

export const getSettingsData = (name: string) => {
  const data = Office.context.document.settings.get(name)

  return data
}

export const removeSettingsData = (name: string) => {
  Office.context.document.settings.remove(name)

  Office.context.document.settings.saveAsync(function (asyncResult) {
    console.log("Removed saved data with status: " + asyncResult.status)
  })
}
