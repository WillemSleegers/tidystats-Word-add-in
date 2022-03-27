import * as React from "react";
import { FileInput } from "./general/FileInput";

import styled from "styled-components";
import { FontSizes, FontWeights } from "@fluentui/theme";

const UploadInstructions = styled.p`
  font-size: ${FontSizes.size14};
  font-weight: ${FontWeights.regular};
`;

type UploadProps = {
  parseStatistics: Function;
};

const Upload = (props: UploadProps) => {
  const { parseStatistics } = props;

  const handleFile = (file: File) => {
    const reader = new FileReader();

    reader.onload = () => {
      parseStatistics(reader.result);
    };

    reader.readAsText(file);
  };

  return (
    <>
      <UploadInstructions>
        Upload your statistics created with the tidystats <a href="https://www.tidystats.io/">R</a> package to get
        started:
      </UploadInstructions>
      <FileInput initialFileName="Upload statistics" handleFile={handleFile} />
    </>
  );
};

export { Upload };
