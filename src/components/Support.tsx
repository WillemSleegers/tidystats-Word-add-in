import { FontSizes, FontWeights } from "@fluentui/theme"
import styled from "styled-components"

const ActionInstructions = styled.p`
  font-size: ${FontSizes.size14};
  font-weight: ${FontWeights.regular};
`

const Support = () => {
  return (
    <>
      <ActionInstructions>
        If you have a question about tidystats or if you're having issues,
        please see the tidystats{" "}
        <a
          href="https://www.tidystats.io/support.html"
          target="_blank"
          rel="noreferrer"
        >
          support
        </a>{" "}
        page.
      </ActionInstructions>
    </>
  )
}

export { Support }
