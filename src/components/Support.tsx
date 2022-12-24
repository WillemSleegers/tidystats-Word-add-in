import { MouseEvent } from "react"
import { makeStyles, Button, Link } from "@fluentui/react-components"
import { removeSettingsData } from "../functions/settings"

const useStyles = makeStyles({
  resetButton: { width: "180px" },
})

export const Support = () => {
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
      <p>
        For more information on how to use tidystats, including examples and
        FAQs, see the tidystats{" "}
        <Link href="https://www.tidystats.io" target="_blank">
          website
        </Link>
        .
      </p>

      <p>Click the button below to reset the help messages.</p>
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
