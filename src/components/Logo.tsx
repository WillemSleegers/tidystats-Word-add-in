import styled from "styled-components"
import { FontSizes, FontWeights } from "@fluentui/theme"

const LogoDiv = styled.div`
  display: flex;
  flex-direction: row;
  align-items: center;
  justify-content: center;
  padding: 0.5rem 1rem;
  background-color: #f3f2f1;
`

const LogoImage = styled.img`
  width: 50px;
  height: 50px;
  padding-right: 0.5rem;
`

const LogoTitle = styled.h1`
  font-size: ${FontSizes.size24};
  font-weight: ${FontWeights.semibold};
`

type LogoProps = {
  title: string
  logo: string
}

const Logo = (props: LogoProps) => {
  const { title, logo } = props

  return (
    <LogoDiv>
      <LogoImage src={logo} alt={title} title={title} />
      <LogoTitle>{title}</LogoTitle>
    </LogoDiv>
  )
}

export { Logo }
