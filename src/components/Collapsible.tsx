import { ReactNode, useState } from "react"

import { Row, RowName } from "./Row"

import { IconButton } from "@fluentui/react/lib/Button"

const chevronDownIcon = { iconName: "ChevronDown" }
const chevronRightIcon = { iconName: "ChevronRight" }
const TableIcon = { iconName: "Table" }

interface CollapsibleProps {
  header: string
  headerBackground?: "gray"
  onInsertClick?: Function
  open?: boolean
  children: ReactNode
}

const Collapsible = (props: CollapsibleProps) => {
  const { header, headerBackground, onInsertClick, open, children } = props

  const [isOpen, setIsOpen] = useState(open)

  const toggleOpen = () => {
    setIsOpen((prev) => !prev)
  }

  return (
    <>
      <div>
        <Row background={headerBackground}>
          <IconButton
            iconProps={!isOpen ? chevronRightIcon : chevronDownIcon}
            styles={{ rootHovered: { background: "rgba(0, 0,0, 0.05)" } }}
            onClick={toggleOpen}
          />

          <RowName isHeader={true}>{header}</RowName>

          {onInsertClick && (
            <IconButton iconProps={TableIcon} onClick={() => onInsertClick} />
          )}
        </Row>
        {isOpen && <div>{children}</div>}
      </div>
    </>
  )
}

export { Collapsible }
