import { useState, useRef, ChangeEvent } from "react"
import styled from "styled-components"

import { IContextualMenuProps } from "@fluentui/react"
import { PrimaryButton } from "@fluentui/react/lib/Button"

import { Tidystats } from "../classes/Tidystats"
import { clearSettingsData } from "../functions/clearSettingsData"

// A fix to deal with some whitespace in the split button
const PrimaryButtonFix = styled.div`
  button:nth-child(2) {
    margin-left: -3px;
  }
`
const ErrorMessage = styled.p`
  color: red;
`

type UploadProps = {
  host: string
  fileName: string | null
  setFileName: Function
  setTidystats: Function
}

const Upload = (props: UploadProps) => {
  const { host, fileName, setFileName, setTidystats } = props

  // Set the state for whether the Remove file option should be disabled or not
  // By default disable the option if the file name is null
  const [disableFile, setDisableFile] = useState(
    fileName === null ? true : false
  )

  // Set the state for whether the user uploaded a non-JSON file or not
  // By default there should be no error
  const [error, setError] = useState(false)

  const fileMenuProps: IContextualMenuProps = {
    items: [
      {
        key: "removeFile",
        text: "Remove file",
        iconProps: { iconName: "Cancel" },
        disabled: disableFile,
        onClick: () => {
          setFileName("Upload statistics")
          setDisableFile(true)
          setTidystats(null)
          setError(false)
          clearSettingsData()
        },
      },
    ],
  }

  const hiddenFileInput = useRef<HTMLInputElement>(null)

  const handleClick = () => {
    if (null !== hiddenFileInput.current) {
      hiddenFileInput.current.click()

      // Set to an empty string to reset the value so a new file can be selected
      hiddenFileInput.current.value = ""
    }
  }

  const handleChange = (event: ChangeEvent<HTMLInputElement>) => {
    if (event.target.files) {
      const file = event.target.files[0]
      setFileName(file.name)
      setDisableFile(false)

      if (file.type === "application/json") {
        const reader = new FileReader()
        reader.onload = () => {
          const text = reader.result
          const data = JSON.parse(text as string)
          const tidystats = new Tidystats(data)
          setTidystats(tidystats)

          if (host === "Word") {
            Office.context.document.settings.set("fileName", file.name)
            Office.context.document.settings.set("data", text)
            Office.context.document.settings.saveAsync(function (asyncResult) {
              console.log("Settings saved with status: " + asyncResult.status)
            })
          }
        }
        reader.readAsText(file)

        setError(false)
      } else {
        setError(true)
        setTidystats(null)
      }
    }
  }

  return (
    <>
      <p>
        Upload your statistics created with the tidystats{" "}
        <a href="https://www.tidystats.io/" target="_blank" rel="noreferrer">
          R package
        </a>{" "}
        to get started:
      </p>
      <PrimaryButtonFix>
        <PrimaryButton
          split
          splitButtonAriaLabel="See cancel file option"
          aria-roledescription="Upload/cancel file"
          onClick={handleClick}
          menuProps={fileMenuProps}
        >
          {fileName === null ? "Upload statistics" : fileName}
        </PrimaryButton>
      </PrimaryButtonFix>
      <input
        type="file"
        accept="application/json"
        ref={hiddenFileInput}
        onChange={handleChange}
        onClick={handleClick}
        style={{ display: "none" }}
      />
      {error && (
        <ErrorMessage>
          File must be a .JSON file (created with the tidystats R package).
        </ErrorMessage>
      )}
    </>
  )
}

export { Upload }
