import { ReactNode } from "react"
import { makeStyles, mergeClasses, tokens } from "@fluentui/react-components"

const useStyles = makeStyles({
  row: {
    minHeight: "2rem",
    display: "flex",
    alignItems: "center",
  },
  indented: {
    marginLeft: "1rem",
  },
  border: {
    borderBottomWidth: "1px",
    borderBottomColor: tokens.colorNeutralBackground4,
    borderBottomStyle: "solid",
  },
  width: {
    width: "5rem",
  },
  header: {
    width: "100%",
  },
  bold: {
    fontWeight: "bold",
  },
  flexGrow: {
    flexGrow: 1,
  },
})

type RowProps = {
  indented?: boolean
  hasBorder?: boolean
  children: ReactNode
}

export const Row = (props: RowProps) => {
  const { indented, hasBorder, children } = props

  const styles = useStyles()

  return (
    <div
      className={mergeClasses(
        styles.row,
        indented && styles.indented,
        hasBorder && styles.border
      )}
    >
      {children}
    </div>
  )
}

type RowNameProps = {
  isHeader?: boolean
  isBold?: boolean
  children: ReactNode
}

export const RowName = (props: RowNameProps) => {
  const { isHeader, isBold, children } = props

  const styles = useStyles()

  return (
    <div
      className={mergeClasses(
        styles.width,
        isHeader && styles.header,
        isBold && styles.bold
      )}
    >
      {children}
    </div>
  )
}

type RowValueProps = {
  children: ReactNode
}

export const RowValue = (props: RowValueProps) => {
  const { children } = props

  const styles = useStyles()

  return <div className={styles.flexGrow}>{children}</div>
}
