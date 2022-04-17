import { Statistic, RangedStatistic } from "../classes/Statistic"

const SMOL = ["p", "r", "R²"]
const INTEGERS = [
  "df",
  "df numerator",
  "df denominator",
  "n",
  "N",
  "k",
  "n parameters",
]

const formatValue = (
  x: Statistic | RangedStatistic,
  decimals: number,
  bound?: "lower" | "upper"
) => {
  let name
  let value

  if (x.symbol) {
    name = x.symbol
  } else {
    name = x.name
  }

  if ("lower" in x) {
    switch (bound) {
      case "lower":
        value = x.lower
        break
      case "upper":
        value = x.upper
        break
      default:
        value = x.value
    }
  } else {
    value = x.value
  }

  if (typeof value == "number") {
    if (INTEGERS.includes(x.name)) {
      value = Intl.NumberFormat(navigator.language, {
        minimumFractionDigits: 0,
      }).format(value)
    } else if (Math.abs(x.value) > 0.1) {
      value = Intl.NumberFormat(navigator.language, {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
      }).format(value)
    } else {
      value = Intl.NumberFormat(navigator.language, {
        minimumSignificantDigits: decimals,
        maximumSignificantDigits: decimals,
      }).format(value)
    }

    if (SMOL.includes(name)) {
      if (value.charAt(0) === "-") {
        value = "-" + value.substring(2)
      } else {
        value = value.substring(1)

        if (name === "p") {
          if (x.value < 0.001) {
            value = "< .001"
          }
        }
      }
    }
  }

  return value
}

export { formatValue }
