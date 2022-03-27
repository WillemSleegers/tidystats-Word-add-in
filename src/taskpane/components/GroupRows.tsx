import * as React from "react";
import { Group } from "../classes/Group";
import { Statistic } from "../classes/Statistic";

import { StatisticsRows } from "./StatisticsRows";
import { Collapsible } from "./Collapsible";

import { insertTable } from "../functions/insertTable";

type GroupRowsProps = {
  name: string;
  statistics?: Statistic[];
  groups?: Group[];
};

const GroupRows = (props: GroupRowsProps) => {
  const { name, statistics, groups } = props;

  let content;

  if (statistics) {
    content = <StatisticsRows statistics={statistics} />;
  }

  if (groups) {
    content = (
      <>
        {groups.map((x) => {
          let group;

          if (x.groups) {
            group = <GroupRows key={x.name} name={x.name} groups={x.groups} />;
          }

          if (x.statistics) {
            group = <GroupRows key={x.name} name={x.name} statistics={x.statistics} />;
          }

          return group;
        })}
      </>
    );
  }

  // Add addTable() function, but only in some cases
  let addTable = false;
  if (name === "Coefficients") addTable = true;

  const handleAddClick = () => {
    insertTable(name, groups);
  };

  return (
    <Collapsible
      primary={false}
      bold={true}
      name={name}
      content={content}
      handleAddClick={addTable ? handleAddClick : undefined}
      open={false}
    />
  );
};

export { GroupRows };
