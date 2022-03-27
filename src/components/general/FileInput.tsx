import { useRef, useState, ChangeEvent } from "react"

import { PrimaryButton } from "@fluentui/react/lib/Button"

type FileInputProps = {
  initialFileName: string
  handleFile: Function
}

const FileInput = (props: FileInputProps) => {
  const { initialFileName, handleFile } = props

  const [fileName, setFileName] = useState(initialFileName)

  const hiddenFileInput = useRef<HTMLInputElement>(null)

  const handleClick = () => {
    if (null !== hiddenFileInput.current) hiddenFileInput.current.click()
  }

  const handleChange = (event: ChangeEvent<HTMLInputElement>) => {
    if (event.target.files) {
      const file = event.target.files[0]
      setFileName(file.name)
      handleFile(file)
    }
  }

  return (
    <>
      <PrimaryButton onClick={handleClick}>{fileName}</PrimaryButton>
      <input
        type="file"
        accept="application/JSON"
        ref={hiddenFileInput}
        onChange={handleChange}
        onClick={handleClick}
        style={{ display: "none" }}
      />
    </>
  )
}

export { FileInput }
