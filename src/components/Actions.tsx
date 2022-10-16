import { MouseEvent } from "react"
import { PrimaryButton } from "@fluentui/react/lib/Button"

import { Tidystats } from "../classes/Tidystats"

import { updateStatistics } from "../functions/updateStatistics"
import { insertInTextCitation } from "../functions/insertInTextCitation"
import { insertFullCitation } from "../functions/insertFullCitation"

type ActionsProps = {
  tidystats?: Tidystats
}

export const Actions = (props: ActionsProps) => {
  const { tidystats } = props
  const buttonWidth = "180px"
  const buttonMargin = "1rem"

  const handleBibTexClick = (e: MouseEvent<HTMLButtonElement>) => {
    const citation = `
      @software{sleegers2021,
        title = {tidystats: {{Save}} Output of Statistical Tests},
        author = {Sleegers, Willem W. A.},
        date = {2021},
        url = {https://doi.org/10.5281/zenodo.4041859},
        version = {0.51}
      }
    `
    navigator.clipboard.writeText(citation)

    const target = e.target as HTMLLabelElement
    target.innerHTML = "Copied!"
    setTimeout(() => {
      target.innerHTML = "Copy BibTex citation"
    }, 2000)
  }

  return (
    <>
      <h3 style={{ marginBottom: "0" }}>Automatic updating</h3>
      <p style={{ marginTop: "0.5rem" }}>
        Automatically update all statistics in your document after uploading a
        new statistics file.
      </p>
      <PrimaryButton
        disabled={tidystats === undefined ? true : false}
        onClick={() => updateStatistics(tidystats!)}
        styles={{ root: { minWidth: buttonWidth } }}
      >
        Update statistics
      </PrimaryButton>

      <h3 style={{ marginBottom: "0" }}>Citation</h3>
      <p style={{ marginTop: "0.5rem" }}>
        Was tidystats useful to you? If so, please consider citing it. Thanks!
      </p>
      <PrimaryButton
        onClick={insertInTextCitation}
        styles={{ root: { minWidth: buttonWidth, marginBottom: buttonMargin } }}
      >
        Insert in-text citation
      </PrimaryButton>
      <PrimaryButton
        onClick={insertFullCitation}
        styles={{ root: { minWidth: buttonWidth, marginBottom: buttonMargin } }}
      >
        Insert full citation
      </PrimaryButton>
      <PrimaryButton
        onClick={handleBibTexClick}
        styles={{ root: { minWidth: buttonWidth } }}
      >
        Copy BibTex citation
      </PrimaryButton>
    </>
  )
}
