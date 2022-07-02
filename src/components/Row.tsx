import styled from "styled-components"

const RowDiv = styled.div<RowProps>`
  display: flex;
  flex-direction: row;
  align-items: center;
  min-height: 32px;
  margin-left: ${(p) => (p.indent ? "1rem" : "0")};
  font-size: 14px;
  background: ${(p) => (p.primary ? "#f3f2f1" : "white")};
  border-bottom: 1px solid #eee;
`

type RowProps = {
  primary: boolean
  indent: boolean
  children?: React.ReactChild | React.ReactChild[]
}

const Row = (props: RowProps) => {
  const { primary, indent, children } = props

  return (
    <RowDiv primary={primary} indent={indent}>
      {children}
    </RowDiv>
  )
}

export { Row }
