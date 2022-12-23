import { MouseEvent } from "react"
import { makeStyles, Button } from "@fluentui/react-components"
import { removeSettingsData } from "../functions/settings"

const useStyles = makeStyles({
  resetButton: { width: "180px" },
})

export const Settings = () => {
  const classes = useStyles()

  const handleClick = (e: MouseEvent<HTMLButtonElement>) => {
    removeSettingsData("dismissedUploadHelpMessage")
    removeSettingsData("dismissedAutomaticUpdatingMessage")

    const target = e.target as HTMLLabelElement
    target.innerHTML = "Reset!"
    setTimeout(() => {
      target.innerHTML = "Reset help messages"
    }, 2000)
  }

  return (
    <>
      <h3>Help</h3>
      <Button
        className={classes.resetButton}
        appearance="primary"
        onClick={handleClick}
      >
        Reset help messages
      </Button>
    </>
  )
}
