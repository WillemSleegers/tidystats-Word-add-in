import { ReactNode } from "react"

type RowProps = {
  indentationLevel?: number
  background?: "gray"
  hasBorder?: boolean
  children: ReactNode
}

export const Row = (props: RowProps) => {
  const { indentationLevel, background, hasBorder, children } = props

  const styles = {
    minHeight: "2rem",
    marginLeft: `${indentationLevel}rem`,
    background: background ? "var(--gray)" : "",
    borderBottom: hasBorder ? "1px solid var(--gray)" : "",
    display: "flex",
    alignItems: "center",
  }

  return <div style={styles}>{children}</div>
}

type RowNameProps = {
  isHeader?: boolean
  isBold?: boolean
  children: ReactNode
}

export const RowName = (props: RowNameProps) => {
  const { isHeader, isBold, children } = props

  const styles = {
    width: isHeader ? "100%" : "5rem",
    fontWeight: isBold ? "bold" : "normal",
  }

  return <div style={styles}>{children}</div>
}

type RowValueProps = {
  children: ReactNode
}

export const RowValue = (props: RowValueProps) => {
  const { children } = props

  const styles = {
    flexGrow: "1",
  }

  return <div style={styles}>{children}</div>
}
