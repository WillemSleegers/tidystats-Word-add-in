import styled from "styled-components"

const RowValueDiv = styled.div`
  flex-grow: 1;
  margin: 5px;
`

type RowValueProps = {
  value: string
}

const RowValue = (props: RowValueProps) => {
  const { value } = props

  return <RowValueDiv>{value}</RowValueDiv>
}

export { RowValue }
