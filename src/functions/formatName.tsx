import { Group } from "../classes/Group"

export const formatName = (x: Group) => {
  if ("name" in x) {
    return x.name!
  } else {
    return x.names![0].name + "-" + x.names![1].name
  }
}
