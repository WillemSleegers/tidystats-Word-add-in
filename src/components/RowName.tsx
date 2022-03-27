import styled from "styled-components"

type RowNameDivProps = {
  header: boolean
  bold: boolean
}

const RowNameDiv = styled.div<RowNameDivProps>`
  flex-grow: ${(p) => (p.header ? 1 : 0)};
  width: 90px;
  font-weight: ${(p) => (p.bold ? "bold" : "normal")};
`
type RowNameProps = {
  header: boolean
  bold: boolean
  name: string | React.ReactNode
  subscript?: string
}

const RowName = (props: RowNameProps) => {
  const { header, bold, name, subscript } = props

  return (
    <RowNameDiv header={header} bold={bold}>
      {name}
      {subscript && <sub>{subscript}</sub>}
    </RowNameDiv>
  )
}

export { RowName }
