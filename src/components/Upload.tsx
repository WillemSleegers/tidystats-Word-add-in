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
  PopoverTrigger,
} from "@fluentui/react-components"
import {
  bundleIcon,
  Dismiss16Regular,
  Dismiss16Filled,
} from "@fluentui/react-icons"
import { Analysis } from "../types"
import { parseAnalyses } from "../utils/parseAnalyses"
import {
  setSettingsData,
  getSettingsData,
  removeSettingsData,
} from "../word/settings"

const useStyles = makeStyles({
  uploadButton: { marginTop: "1rem" },
  popover: {
    paddingTop: "0",
    maxWidth: "80%",
  },
  dismissMessageButton: {
    marginLeft: "1rem",
    color: tokens.colorNeutralBackground1,
    ":hover": {
      color: tokens.colorNeutralBackground1,
      backgroundColor: tokens.colorBrandBackgroundHover,
    },
    ":hover:active": {
      color: tokens.colorNeutralBackground1,
      backgroundColor: tokens.colorBrandBackgroundHover,
    },
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
  setAnalyses: (analyses: Analysis[] | undefined) => void
}

const DismissIcon = bundleIcon(Dismiss16Filled, Dismiss16Regular)

export const Upload = (props: UploadProps) => {
  const { setAnalyses } = props

  const styles = useStyles()

  const fileInput = useRef<HTMLButtonElement>(null)
  const hiddenFileInput = useRef<HTMLInputElement>(null)
  const positioningRef = useRef<PositioningImperativeRef>(null)

  const [fileName, setFileName] = useState<string | null>(() =>
    getSettingsData("fileName")
  )
  const [showHelpMessage, setShowHelpMessage] = useState(
    () => !getSettingsData("dismissedUploadHelpMessage")
  )
  const [showErrorMessage, setShowErrorMessage] = useState(false)

  useEffect(() => {
    positioningRef.current?.setTarget(fileInput.current!)
  }, [])

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
      setAnalyses(undefined) // reset the statistics

      if (file.type === "application/json") {
        const reader = new FileReader()
        reader.onload = () => {
          const text = reader.result as string
          setAnalyses(parseAnalyses(JSON.parse(text)))

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
    setAnalyses(undefined)

    if (showErrorMessage) setShowErrorMessage(false)

    removeSettingsData("fileName")
    removeSettingsData("statistics")
  }

  return (
    <>
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
        <PopoverTrigger>
          <div>
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
          </div>
        </PopoverTrigger>
        <PopoverSurface
          className={styles.popover}
          aria-label="Upload statistics"
        >
          <p>Upload statistics created with the analyses R package here.</p>
          <Button
            as="a"
            href="https://willemsleegers.github.io/analyses/articles/word-add-in.html"
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

      {showErrorMessage && (
        <div>
          <Label className={styles.errorMessage} weight="semibold">
            File must be a analyses JSON file.
          </Label>
        </div>
      )}
    </>
  )
}

