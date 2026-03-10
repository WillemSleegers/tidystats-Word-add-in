import { makeStyles, tokens } from "@fluentui/react-components"
import src from "../assets/analyses-icon.svg"

const useStyles = makeStyles({
  header: {
    display: "flex",
    columnGap: "0.5rem",
    alignItems: "center",
    justifyContent: "center",
    paddingTop: "0.5rem",
    paddingBottom: "0.5rem",
    backgroundColor: tokens.colorNeutralBackground4,
  },
})

export const Header = () => {
  const styles = useStyles()

  return (
    <div className={styles.header}>
      <img width={48} height={48} src={src} alt="analyses" title="analyses" />
      <h1>analyses</h1>
    </div>
  )
}
