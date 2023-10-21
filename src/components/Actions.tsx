import { useRef } from "react"
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

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    rowGap: "0.5rem",
    maxWidth: "220px",
  },
  h3: {
    marginBottom: "4px",
  },
  p: {
    marginTop: "0",
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

  return (
    <>
      <h3 className={styles.h3}>Update statistics</h3>
      <p className={styles.p}>
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

      <h3 className={styles.h3}>Replace statistics</h3>
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
    </>
  )
}
