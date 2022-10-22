import { Tidystats } from "../classes/Tidystats"
import { AnalysisRows } from "./AnalysisRows"

type AnalysesProps = {
  tidystats: Tidystats
}

export const Analyses = (props: AnalysesProps) => {
  const { tidystats } = props

  return (
    <>
      <h2>Analyses</h2>
      {tidystats.analyses.map((x) => {
        return <AnalysisRows key={x.identifier} analysis={x} />
      })}
    </>
  )
}
