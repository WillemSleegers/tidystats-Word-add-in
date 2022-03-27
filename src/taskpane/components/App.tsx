import * as React from "react";
import { useState } from "react";

import { Tidystats } from "../classes/Tidystats";

import { AnalysesTable } from "./AnalysesTable";
import { Logo } from "./Logo";
import { Upload } from "./Upload";
import { Progress } from "./Progress";
import { Actions } from "./Actions";
import { Support } from "./Support";

import styled from "styled-components";
import { initializeIcons } from "@fluentui/font-icons-mdl2";

// import data from "../assets/results.json" // For debugging

initializeIcons();

const Main = styled.div`
  margin: 0.5rem;
`;

type AppProps = {
  isOfficeInitialized: boolean;
};

const App = (props: AppProps) => {
  const { isOfficeInitialized } = props;

  //const [tidystats, setTidystats] = useState<Tidystats | undefined>(
  //  new Tidystats(data)
  //) // For debugging
  const [tidystats, setTidystats] = useState<Tidystats | undefined>();

  const parseStatistics = (text: string) => {
    const data = JSON.parse(text as string);
    const tidystats = new Tidystats(data);

    //setTidystats(undefined)
    setTidystats(tidystats);
  };

  const statisticsUpload = <Upload parseStatistics={parseStatistics} />;

  let content;
  if (isOfficeInitialized) {
    if (tidystats) {
      content = <AnalysesTable tidystats={tidystats} />;
    }
  } else {
    content = <Progress message="Please sideload your addin to see app body." />;
  }

  const support = <Support />;

  return (
    <>
      <Logo title="tidystats" logo={require("./../../../assets/tidystats-icon-64.png")} />
      <Main>
        {statisticsUpload}
        {tidystats && content}
        {tidystats && <Actions tidystats={tidystats} />}
        {support}
      </Main>
    </>
  );
};

export { App };
