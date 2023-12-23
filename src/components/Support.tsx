import { MouseEvent } from "react"
import { makeStyles, Button, Link } from "@fluentui/react-components"
import { removeSettingsData } from "../functions/settings"

const useStyles = makeStyles({
  h3: {
    marginBottom: "4px",
  },
  p: {
    marginTop: "0",
    marginBottom: "8px",
  },
  resetButton: { width: "180px" },
})

export const Support = () => {
  const styles = useStyles()

  const handleClick = (e: MouseEvent<HTMLButtonElement>) => {
    removeSettingsData("dismissedUploadHelpMessage")
    removeSettingsData("dismissedUpdateHelpMessage")

    const target = e.target as HTMLLabelElement
    target.innerHTML = "Done"
    setTimeout(() => {
      target.innerHTML = "Reset help messages"
    }, 2000)
  }

  return (
    <>
      <h3 className={styles.h3}>Help</h3>
      <p className={styles.p}>
        For more information on how to use tidystats, including examples and
        FAQs, see the tidystats{" "}
        <Link href="https://www.tidystats.io" target="_blank">
          website
        </Link>
        .
      </p>

      <h3 className={styles.h3}>Help messages</h3>
      <p className={styles.p}>
        Click the button below to reset the help messages.
      </p>
      <Button
        className={styles.resetButton}
        appearance="primary"
        onClick={handleClick}
      >
        Reset help messages
      </Button>

      <h3 className={styles.h3}>Cite tidystats</h3>
      <p className={styles.p}>
        Please consider{" "}
        <Link target="_blank" href="https://www.tidystats.io/citation/">
          citing
        </Link>{" "}
        tidystats if you've found it useful.
      </p>
    </>
  )
}
