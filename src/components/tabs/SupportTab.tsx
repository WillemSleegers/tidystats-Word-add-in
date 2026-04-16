import { MouseEvent } from "react"
import { makeStyles, Button, Link, tokens } from "@fluentui/react-components"
import { removeSettingsData } from "../../word/settings"

const useStyles = makeStyles({
  h3: {
    marginBottom: "4px",
  },
  p: {
    marginTop: "0",
    marginBottom: "8px",
  },
  blockquote: {
    display: "flex",
    gap: "0.75rem",
    margin: "0",
    marginBottom: "8px",
    "::before": {
      content: '""',
      display: "block",
      minWidth: "3px",
      backgroundColor: tokens.colorBrandBackground,
    },
  },
  resetButton: { width: "180px" },
})

export const SupportTab = () => {
  const styles = useStyles()

  const handleClick = (e: MouseEvent<HTMLButtonElement>) => {
    removeSettingsData("dismissedUploadHelpMessage")
    removeSettingsData("dismissedUpdateHelpMessage")

    const target = e.target as HTMLLabelElement
    target.innerHTML = "Done"
    setTimeout(() => {
      target.innerHTML = "Re-enable help tips"
    }, 2000)
  }

  return (
    <>
      <h3 className={styles.h3}>Website</h3>
      <p className={styles.p}>
        See the tidystats{" "}
        <Link
          href="https://willemsleegers.github.io/tidystats/articles/word-add-in.html"
          target="_blank"
        >
          website
        </Link>{" "}
        for more information on how to use tidystats.
      </p>

      <h3 className={styles.h3}>Help tips</h3>
      <p className={styles.p}>
        Click the button below to re-enable the help tips.
      </p>
      <Button
        className={styles.resetButton}
        appearance="primary"
        onClick={handleClick}
      >
        Re-enable help tips
      </Button>

      <h3 className={styles.h3}>Cite tidystats</h3>
      <p className={styles.p}>
        Please consider citing tidystats if you've found it useful.
      </p>
      <blockquote className={styles.blockquote}>
        <p style={{ margin: 0 }}>
          Sleegers, W. W. A. (2026).{" "}
          <em>tidystats: Save output of statistical tests</em>.{" "}
          <Link target="_blank" href="https://doi.org/10.5281/zenodo.4041858">
            https://doi.org/10.5281/zenodo.4041858
          </Link>
        </p>
      </blockquote>
    </>
  )
}
