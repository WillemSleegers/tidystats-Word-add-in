import * as React from "react";
import { useState } from "react";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { FontSizes, FontWeights } from "@fluentui/theme";
import styled from "styled-components";

import { Tidystats } from "../classes/Tidystats";

import { updateStatistics } from "../functions/updateStatistics";
import { insertInTextCitation } from "../functions/insertInTextCitation";
import { insertFullCitation } from "../functions/insertFullCitation";

const ActionHeader = styled.h3`
  margin-bottom: 0;
`;

const ActionInstructions = styled.p`
  font-size: ${FontSizes.size14};
  font-weight: ${FontWeights.regular};
`;
const ActionButton = styled(PrimaryButton)`
  display: block;
  margin-bottom: 0.5rem;
  min-width: 180px;
`;

type ActionsProps = {
  tidystats: Tidystats;
};

const Actions = (props: ActionsProps) => {
  const { tidystats } = props;

  const [bibTexButtonLabel, setBibTexButtonLabel] = useState("Copy BibTex citation");

  const handleBibTexClick = () => {
    navigator.clipboard.writeText(`
      @software{sleegers2021,
        title = {tidystats: {{Save}} Output of Statistical Tests},
        author = {Sleegers, Willem W. A.},
        date = {2021},
        url = {https://doi.org/10.5281/zenodo.4041859},
        version = {0.51}
      }
    `);
    setBibTexButtonLabel("Copied!");
    setTimeout(() => {
      setBibTexButtonLabel("Copy BibTex citation");
    }, 2000);
  };

  return (
    <>
      <ActionHeader>Actions:</ActionHeader>
      <ActionInstructions>
        Automatically update all statistics in your document after uploading a new statistics file.
      </ActionInstructions>
      <ActionButton onClick={() => updateStatistics(tidystats)}>Update statistics</ActionButton>
      <ActionInstructions>Was tidystats useful to you? If so, please consider citing it. Thanks!</ActionInstructions>

      <ActionButton onClick={insertInTextCitation}>Insert in-text citation</ActionButton>
      <ActionButton onClick={insertFullCitation}>Insert full citation</ActionButton>
      <ActionButton onClick={handleBibTexClick}>{bibTexButtonLabel}</ActionButton>
    </>
  );
};

export { Actions };
