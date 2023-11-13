import { useEffect, useRef, useState } from "react"
import {
  Button,
  makeStyles,
  useId,
  Input,
  Label,
  Checkbox,
  Popover,
  PopoverSurface,
  tokens,
  PopoverTrigger,
  shorthands,
} from "@fluentui/react-components"
import { Tidystats } from "../classes/Tidystats"
import { updateStatistics } from "../functions/updateStatistics"
import { replaceStatistics } from "../functions/replaceStatistics"
import { getSettingsData, setSettingsData } from "../functions/settings"

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    ...shorthands.gap("2px"),
    maxWidth: "220px",
  },
  h3: {
    marginBottom: "8px",
  },
  popover: {
    paddingTop: "0",
    maxWidth: "80%",
  },
  dismissMessageButton: {
    marginLeft: "1rem",
    color: tokens.colorNeutralBackground1,
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

  const [showHelpMessage, setShowHelpMessage] = useState(false)

  useEffect(() => {
    const messageDismissed = getSettingsData("dismissedUpdateHelpMessage")
    setShowHelpMessage(!messageDismissed)
  }, [])

  const handleReplaceStatisticsClick = () => {
    const value = replacementInput.current!.value
    const highlight = replacementCheck.current!.checked

    replaceStatistics(value ? value : "NA", highlight)
  }

  const handleMessageClick = () => {
    setShowHelpMessage(false)
    setSettingsData("dismissedUpdateHelpMessage", true)
  }

  return (
    <>
      <h3 className={styles.h3}>Update statistics</h3>
      <Popover
        withArrow
        open={showHelpMessage}
        trapFocus
        positioning={{
          position: "below",
          align: "start",
        }}
        appearance="brand"
      >
        <PopoverTrigger disableButtonEnhancement>
          <Button
            appearance="primary"
            disabled={tidystats ? false : true}
            onClick={() => updateStatistics(tidystats!)}
          >
            Update
          </Button>
        </PopoverTrigger>
        <PopoverSurface
          className={styles.popover}
          aria-label="Update statistics"
        >
          <p>
            Automatically update reported statistics after uploading a new file.
          </p>
          <Button
            as="a"
            href="https://www.tidystats.io/word-add-in/"
            target="_blank"
            aria-label="Learn more"
          >
            Learn more
          </Button>
          <Button
            className={styles.dismissMessageButton}
            onClick={handleMessageClick}
            appearance="outline"
            aria-label="Got it"
          >
            Got it
          </Button>
        </PopoverSurface>
      </Popover>

      <h3 className={styles.h3}>Replace statistics</h3>
      <div className={styles.root}>
        <Label htmlFor={replacementInputId}>Replacement:</Label>
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
          Replace
        </Button>
      </p>
    </>
  )
}
