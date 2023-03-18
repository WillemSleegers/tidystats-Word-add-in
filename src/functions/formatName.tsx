import { Group } from "../classes/Group"

export const formatName = (x: Group) => {
  if (x.name) {
    return x.name
  } else {
    return x.names![0].name + "-" + x.names![1].name
  }
}
