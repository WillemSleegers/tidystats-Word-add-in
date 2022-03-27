import styled from "styled-components"

const RowValueDiv = styled.div`
  flex-grow: 1;
`

type RowValueProps = {
  value: string
}

const RowValue = (props: RowValueProps) => {
  const { value } = props

  return <RowValueDiv>{value}</RowValueDiv>
}

export { RowValue }
