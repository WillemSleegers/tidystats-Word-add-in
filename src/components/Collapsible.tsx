import { ReactNode, useState } from "react"
import styled from "styled-components"

import { Row } from "../components/Row"
import { RowName } from "../components/RowName"

import { IIconProps } from "@fluentui/react"
import { IconButton } from "@fluentui/react/lib/Button"

export const Content = styled.div`
  padding-left: 16px;
`

const chevronDownIcon: IIconProps = { iconName: "ChevronDown" }
const chevronRightIcon: IIconProps = { iconName: "ChevronRight" }
const settingsIcon: IIconProps = { iconName: "Settings" }
const addIcon: IIconProps = { iconName: "Add" }

interface CollapsibleProps {
  primary: boolean
  bold: boolean
  name: string
  handleSettingsClick?: Function
  handleAddClick?: Function
  content: ReactNode
  open?: boolean
  disabled?: boolean
}

const Collapsible = (props: CollapsibleProps) => {
  const {
    primary,
    bold,
    name,
    handleSettingsClick,
    handleAddClick,
    content,
    open,
    disabled,
  } = props

  const [isOpen, setIsOpen] = useState(open)

  const toggleOpen = () => {
    setIsOpen((prev) => !prev)
  }

  return (
    <>
      <Row primary={primary} indent={false}>
        <>
          {!disabled && (
            <IconButton
              iconProps={!isOpen ? chevronRightIcon : chevronDownIcon}
              onClick={toggleOpen}
            />
          )}
        </>
        <RowName header={true} bold={bold} name={name} />
        <>
          {handleSettingsClick && (
            <IconButton
              iconProps={settingsIcon}
              onClick={() => handleSettingsClick()}
            />
          )}
        </>
        <>
          {handleAddClick && (
            <IconButton iconProps={addIcon} onClick={() => handleAddClick()} />
          )}
        </>
      </Row>
      <Content>{isOpen && content}</Content>
    </>
  )
}

export { Collapsible }
