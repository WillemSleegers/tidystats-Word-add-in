import src from "../assets/tidystats-icon.svg"

export const Header = () => {
  const styles = {
    margin: "-0.5rem -0.5rem 0 -0.5rem",
    display: "flex",
    flexDirection: "row" as "row",
    alignItems: "center",
    justifyContent: "center",
    gap: "0.5rem",
    padding: "0.5rem 1rem",
    backgroundColor: "var(--gray)",
  }

  return (
    <div style={styles}>
      <img width={48} height={48} src={src} alt="tidystats" title="tidystats" />
      <h1>tidystats</h1>
    </div>
  )
}
