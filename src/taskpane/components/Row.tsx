import * as React from "react";
import styled from "styled-components";

const RowDiv = styled.div<RowProps>`
  display: flex;
  flex-direction: row;
  align-items: center;
  min-height: 32px;
  font-size: 14px;
  background: ${(p) => (p.primary ? "#f3f2f1" : "white")};
  border-bottom: 1px solid #eee;
`;

type RowProps = {
  primary: boolean;
  children?: React.ReactChild | React.ReactChild[];
};

const Row = (props: RowProps) => {
  const { primary, children } = props;

  return <RowDiv primary={primary}>{children}</RowDiv>;
};

export { Row };
