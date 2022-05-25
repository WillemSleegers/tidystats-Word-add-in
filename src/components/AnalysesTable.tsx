import { Tidystats } from "../classes/Tidystats"
import { AnalysisRows } from "./AnalysisRows"

type AnalysesTableProps = {
  tidystats: Tidystats
}

const AnalysesTable = (props: AnalysesTableProps) => {
  const { tidystats } = props

  return (
    <>
      <h3>Analyses</h3>
      {tidystats.analyses.map((x) => {
        return <AnalysisRows key={x.identifier} analysis={x} />
      })}
    </>
  )
}

export { AnalysesTable }
