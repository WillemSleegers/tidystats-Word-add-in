import * as React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react";

type ProgressProps = {
  message: string;
};

const Progress = (props: ProgressProps) => {
  const { message } = props;

  return (
    <div className="ms-u-fadeIn500">
      <Spinner size={SpinnerSize.large} label={message} />
    </div>
  );
};

export { Progress };
