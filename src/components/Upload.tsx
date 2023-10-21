import { useState, useRef, ChangeEvent, useEffect } from "react"
import {
  makeStyles,
  SplitButton,
  Menu,
  MenuTrigger,
  MenuButtonProps,
  MenuPopover,
  MenuList,
  MenuItem,
  Popover,
  PopoverSurface,
  Button,
  PositioningImperativeRef,
  Label,
  tokens,
} from "@fluentui/react-components"
import {
  bundleIcon,
  Dismiss16Regular,
  Dismiss16Filled,
} from "@fluentui/react-icons"
import { Tidystats } from "../classes/Tidystats"
import {
  setSettingsData,
  getSettingsData,
  removeSettingsData,
} from "../functions/settings"

const useStyles = makeStyles({
  uploadButton: { marginTop: "1rem" },
  popover: {
    paddingTop: "0",
    maxWidth: "80%",
  },
  dismissMessageButton: {
    marginLeft: "1rem",
    color: tokens.colorNeutralBackground1,
  },
  errorMessage: {
    display: "inline-block",
    marginTop: "0.5rem",
    color: tokens.colorPaletteRedBackground3,
  },
  hiddenFileInput: {
    display: "none",
  },
})

type UploadProps = {
  setTidystats: Function
}

const Upload = (props: UploadProps) => {
  const { setTidystats } = props

  const styles = useStyles()

  const fileInput = useRef(null)
  const hiddenFileInput = useRef<HTMLInputElement>(null)
  const positioningRef = useRef<PositioningImperativeRef>(null)

  const [fileName, setFileName] = useState<string | null>()
  const [showHelpMessage, setShowHelpMessage] = useState(false)
  const [showErrorMessage, setShowErrorMessage] = useState(false)

  useEffect(() => {
    const savedFileName = getSettingsData("fileName")
    setFileName(savedFileName)

    const messageDismissed = getSettingsData("dismissedUploadHelpMessage")
    setShowHelpMessage(!messageDismissed)
  }, [])

  useEffect(() => {
    positioningRef.current?.setTarget(fileInput.current!)
  }, [fileInput, positioningRef])

  const handleInputClick = () => {
    if (null !== hiddenFileInput.current) {
      hiddenFileInput.current.click()

      // Reset the value so a new file can be selected
      hiddenFileInput.current.value = ""
    }
  }

  const handleInputChange = (event: ChangeEvent<HTMLInputElement>) => {
    if (event.target.files) {
      const file = event.target.files[0]

      setFileName(file.name)
      setTidystats(null) // reset the statistics

      if (file.type === "application/json") {
        const reader = new FileReader()
        reader.onload = () => {
          const text = reader.result as string
          const tidystats = new Tidystats(JSON.parse(text))
          setTidystats(tidystats)

          setSettingsData("fileName", file.name)
          setSettingsData("statistics", text)
        }
        reader.readAsText(file)

        if (showErrorMessage) setShowErrorMessage(false)
      } else {
        setShowErrorMessage(true)
      }
    }
  }

  const handleMessageClick = () => {
    setShowHelpMessage(false)
    setSettingsData("dismissedUploadHelpMessage", true)
  }

  const handleRemoveFileClick = () => {
    setFileName(null)
    setTidystats(null)

    if (showErrorMessage) setShowErrorMessage(false)

    removeSettingsData("fileName")
    removeSettingsData("statistics")
  }

  const DismissIcon = bundleIcon(Dismiss16Filled, Dismiss16Regular)

  return (
    <>
      <Menu positioning="below-end">
        <MenuTrigger disableButtonEnhancement>
          {(triggerProps: MenuButtonProps) => (
            <SplitButton
              ref={fileInput}
              className={styles.uploadButton}
              appearance="primary"
              primaryActionButton={{ onClick: handleInputClick }}
              menuButton={triggerProps}
            >
              {fileName ? fileName : "Upload statistics"}
            </SplitButton>
          )}
        </MenuTrigger>

        <MenuPopover>
          <MenuList>
            <MenuItem
              icon={<DismissIcon />}
              disabled={!fileName}
              onClick={handleRemoveFileClick}
            >
              Remove file
            </MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <input
        ref={hiddenFileInput}
        className={styles.hiddenFileInput}
        type="file"
        accept="application/json"
        onChange={handleInputChange}
        onClick={handleInputClick}
      />

      <Popover
        withArrow
        open={showHelpMessage}
        trapFocus
        positioning={{
          positioningRef,
          position: "below",
          align: "start",
        }}
        appearance="brand"
      >
        <PopoverSurface
          className={styles.popover}
          aria-label="Upload statistics"
        >
          <p>
            Upload your statistics created with the tidystats R package here.
          </p>
          <Button
            as="a"
            href="https://www.tidystats.io/r-package/"
            target="_blank"
            aria-label="Learn more"
          >
            Learn more
          </Button>
          <Button
            className={styles.dismissMessageButton}
            onClick={handleMessageClick}
            appearance="outline"
            aria-label="Got it!"
          >
            Got it!
          </Button>
        </PopoverSurface>
      </Popover>

      {showErrorMessage && (
        <div>
          <Label className={styles.errorMessage} weight="semibold">
            File must be a tidystats JSON file.
          </Label>
        </div>
      )}
    </>
  )
}

export { Upload }
