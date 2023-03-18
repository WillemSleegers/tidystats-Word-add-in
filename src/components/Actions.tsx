import { useRef, MouseEvent } from "react"
import {
  Button,
  makeStyles,
  useId,
  Input,
  Label,
  Checkbox,
} from "@fluentui/react-components"
import { Tidystats } from "../classes/Tidystats"
import { updateStatistics } from "../functions/updateStatistics"
import { replaceStatistics } from "../functions/replaceStatistics"
import { insertText } from "../functions/insertText"
import { citation, bibTexCitation } from "../assets/citation"

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    rowGap: "5px",
    maxWidth: "220px",
  },
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

type ActionsProps = {
  tidystats?: Tidystats
}

export const Actions = (props: ActionsProps) => {
  const { tidystats } = props

  const styles = useStyles()

  const replacementInputId = useId("input")

  const replacementInput = useRef<HTMLInputElement>(null)
  const replacementCheck = useRef<HTMLInputElement>(null)

  const handleReplaceStatisticsClick = () => {
    const value = replacementInput.current!.value
    const highlight = replacementCheck.current!.checked

    replaceStatistics(value ? value : "NA", highlight)
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
      <h3>Update statistics</h3>
      <p>
        Automatically update all reported statistics after uploading a new file.
      </p>
      <Button
        id="updateStatsButton"
        appearance="primary"
        disabled={tidystats ? false : true}
        onClick={() => updateStatistics(tidystats!)}
      >
        Update statistics
      </Button>

      <h3>Replace statistics</h3>
      <div className={styles.root}>
        <Label htmlFor={replacementInputId}>
          Replace reported statistics with:
        </Label>
        <Input
          placeholder="NA"
          id={replacementInputId}
          ref={replacementInput}
        />
      </div>
      <div>
        <Checkbox
          ref={replacementCheck}
          label="Highlight replacements"
          defaultChecked={true}
        />
      </div>
      <p>
        <Button appearance="primary" onClick={handleReplaceStatisticsClick}>
          Replace statistics
        </Button>
      </p>

      <h3>Cite tidystats</h3>
      <p>Please consider citing tidystats if you've found it useful. Thanks!</p>
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
