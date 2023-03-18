import { Statistic, RangedStatistic } from "../classes/Statistic"

const SMOL = ["p", "p̂", "r", "R²", "ges"]
const INTEGERS = [
  "df",
  "df numerator",
  "df denominator",
  "residual df",
  "n",
  "N",
  "k",
  "n parameters",
  "number of trials",
  "truncation lag",
]
const STATISTIC_INTEGERS = ["S", "T", "n", "k"]

export const formatValue = (
  x: Statistic | RangedStatistic,
  decimals: number,
  bound?: "lower" | "upper"
) => {
  let value

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
      if (value % 1 == 0) {
        value = Intl.NumberFormat(navigator.language, {
          minimumFractionDigits: 0,
        }).format(value)
      } else {
        value = Intl.NumberFormat(navigator.language, {
          minimumFractionDigits: decimals,
          maximumFractionDigits: decimals,
        }).format(value)
      }
    } else {
      if (x.name == "statistic" && STATISTIC_INTEGERS.includes(x.symbol!)) {
        value = Intl.NumberFormat(navigator.language, {
          minimumFractionDigits: 0,
        }).format(value)
      } else {
        // Add an extra decimal for each starting 0 in the decimals
        const extra_decimals =
          Math.abs(value) < 1
            ? -Math.floor(Math.log10(Math.abs(value) % 1)) - 1
            : 0

        value = Intl.NumberFormat(navigator.language, {
          minimumFractionDigits: decimals,
          maximumFractionDigits: Math.min(20, decimals + extra_decimals),
        }).format(value)
      }
    }

    const name = x.symbol ? x.symbol : x.name

    if ((x.value > 1000000 || 1 / x.value > 1000000) && x.value != 0) {
      console.log(x.value)
      value = x.value.toExponential(decimals)
    }

    if (SMOL.includes(name) && x.value < 1) {
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
  } else {
    if (value == "Inf") {
      value = "∞"
    }
  }

  return value
}
