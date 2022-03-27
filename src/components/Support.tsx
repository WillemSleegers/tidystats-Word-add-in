import { FontSizes, FontWeights } from "@fluentui/theme"
import styled from "styled-components"

const ActionHeader = styled.h3`
  margin-bottom: 0;
`

const ActionInstructions = styled.p`
  font-size: ${FontSizes.size14};
  font-weight: ${FontWeights.regular};
`

const Support = () => {
  return (
    <>
      <ActionHeader>Support</ActionHeader>
      <ActionInstructions>
        If you have a question about tidystats or if you're having issues,
        please see the tidystats{" "}
        <a href="https://www.tidystats.io/support.html">support</a> page.
      </ActionInstructions>
    </>
  )
}

export { Support }
