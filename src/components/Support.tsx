import { MouseEvent } from "react"
import { makeStyles, Button, Link } from "@fluentui/react-components"
import { removeSettingsData } from "../functions/settings"
import { insertText } from "../functions/insertText"
import { citation, bibTexCitation } from "../assets/citation"

const useStyles = makeStyles({
  h3: {
    marginBottom: "4px",
  },
  p: {
    marginTop: "0",
    marginBottom: "8px",
  },
  resetButton: { width: "180px" },
  citation: {
    paddingLeft: "0.5rem",
    borderLeftWidth: "0.2rem",
    borderLeftStyle: "solid",
    borderLeftColor: "gray",
  },
  citationButtonsWrapper: {
    display: "flex",
    columnGap: "0.5rem",
  },
  citationBibtexButton: {
    width: "11rem",
  },
})

export const Support = () => {
  const styles = useStyles()

  const handleClick = (e: MouseEvent<HTMLButtonElement>) => {
    removeSettingsData("dismissedUploadHelpMessage")
    removeSettingsData("dismissedAutomaticUpdatingMessage")

    const target = e.target as HTMLLabelElement
    target.innerHTML = "Reset!"
    setTimeout(() => {
      target.innerHTML = "Reset help messages"
    }, 2000)
  }

  const handleCopyBibTexClick = (e: MouseEvent<HTMLButtonElement>) => {
    navigator.clipboard.writeText(bibTexCitation)

    const target = e.target as HTMLLabelElement
    target.innerHTML = "Copied!"
    setTimeout(() => {
      target.innerHTML = "Copy BibTex citation"
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
        Please consider citing tidystats if you've found it useful. Thanks!
      </p>
      <p className={styles.citation}>{citation}</p>
      <div className={styles.citationButtonsWrapper}>
        <Button appearance="primary" onClick={() => insertText(citation)}>
          Insert citation
        </Button>
        <Button
          className={styles.citationBibtexButton}
          appearance="primary"
          onClick={handleCopyBibTexClick}
        >
          Copy BibTex citation
        </Button>
      </div>
    </>
  )
}
