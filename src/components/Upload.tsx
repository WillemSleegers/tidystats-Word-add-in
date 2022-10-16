import { useState, useRef, ChangeEvent } from "react"
import {
  IContextualMenuProps,
  TeachingBubble,
  DirectionalHint,
  MessageBar,
  MessageBarType,
} from "@fluentui/react"
import { PrimaryButton } from "@fluentui/react/lib/Button"

import { Tidystats } from "../classes/Tidystats"
import { clearSettingsData } from "../functions/clearSettingsData"

type UploadProps = {
  setTidystats: Function
}

const Upload = (props: UploadProps) => {
  const { setTidystats } = props

  // Load settings
  const settings = Office.context.document.settings
  const savedFileName = settings.get("fileName")
  const messageDismissed = settings.get("messageDismissed")

  const hiddenFileInput = useRef<HTMLInputElement>(null)
  const [fileName, setFileName] = useState(
    savedFileName === null ? "Upload statistics" : savedFileName
  )

  const [showErrorMessage, setShowErrorMessage] = useState(false)
  const [hideTeachingBubble, setHideTeachingBubble] = useState(
    messageDismissed === null ? true : messageDismissed
  )

  const fileMenuProps: IContextualMenuProps = {
    items: [
      {
        key: "removeFile",
        text: "Remove file",
        iconProps: { iconName: "Cancel" },
        onClick: () => {
          setFileName("Upload statistics")
          setTidystats(null)
          setShowErrorMessage(false)
          clearSettingsData()
        },
      },
    ],
  }

  const handleClick = () => {
    if (null !== hiddenFileInput.current) {
      hiddenFileInput.current.click()

      // Reset the value so a new file can be selected
      hiddenFileInput.current.value = ""
    }
  }

  const handleChange = (event: ChangeEvent<HTMLInputElement>) => {
    if (event.target.files) {
      const file = event.target.files[0]
      setFileName(file.name)

      if (file.type === "application/json") {
        const reader = new FileReader()
        reader.onload = () => {
          const text = reader.result
          const data = JSON.parse(text as string)
          const tidystats = new Tidystats(data)
          setTidystats(tidystats)

          Office.context.document.settings.set("fileName", file.name)
          Office.context.document.settings.set("data", text)
          Office.context.document.settings.saveAsync(function (asyncResult) {
            console.log("Settings saved with status: " + asyncResult.status)
          })
        }
        reader.readAsText(file)

        setShowErrorMessage(false)
      } else {
        setShowErrorMessage(true)
        setTidystats(null)
      }
    }
  }

  return (
    <>
      <PrimaryButton
        id="fileUpload"
        split
        splitButtonAriaLabel="See cancel file option"
        aria-roledescription="Upload/cancel file"
        onClick={handleClick}
        menuProps={fileMenuProps}
        styles={{ splitButtonMenuButton: { marginLeft: "-3px" } }}
      >
        {fileName === null ? "Upload statistics" : fileName}
      </PrimaryButton>
      <input
        type="file"
        accept="application/json"
        ref={hiddenFileInput}
        onChange={handleChange}
        onClick={handleClick}
        style={{ display: "none" }}
      />
      {!hideTeachingBubble && (
        <TeachingBubble
          target={"#fileUpload"}
          calloutProps={{
            directionalHint: DirectionalHint.bottomCenter,
            calloutWidth: window.innerWidth - 16,
          }}
          primaryButtonProps={{
            text: "Learn more",
            href: "https://www.tidystats.io/r-package/",
            target: "_blank",
            type: "button",
            styles: { rootHovered: { textDecoration: "none" } },
          }}
          secondaryButtonProps={{
            text: "Got it!",
            onClick: () => {
              setHideTeachingBubble(true)
              Office.context.document.settings.set("messageDismissed", true)
              Office.context.document.settings.saveAsync(function (
                asyncResult
              ) {
                console.log("Settings saved with status: " + asyncResult.status)
              })
            },
          }}
        >
          Upload your statistics created with the tidystats R package.
        </TeachingBubble>
      )}

      {showErrorMessage && (
        <MessageBar
          messageBarType={MessageBarType.error}
          styles={{ root: { marginTop: "1rem" } }}
        >
          File must be a tidystats JSON file.
        </MessageBar>
      )}
    </>
  )
}

export { Upload }
