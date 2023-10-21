import { makeStyles, tokens } from "@fluentui/react-components"
import src from "../assets/tidystats-icon.svg"

const useStyles = makeStyles({
  header: {
    marginTop: "-0.5rem",
    marginRight: "-0.5rem",
    marginLeft: "-0.5rem",
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
      <img width={48} height={48} src={src} alt="tidystats" title="tidystats" />
      <h1>tidystats</h1>
    </div>
  )
}
