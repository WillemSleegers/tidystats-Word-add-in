import { MouseEvent } from "react"
import { PrimaryButton } from "@fluentui/react/lib/Button"

export const Support = () => {
  const handleResetHelpClick = (e: MouseEvent<HTMLButtonElement>) => {
    Office.context.document.settings.set("messageDismissed", false)
    Office.context.document.settings.saveAsync(function (asyncResult) {
      console.log("Settings saved with status: " + asyncResult.status)
    })
  }

  return (
    <>
      <p>
        If you have a question about tidystats or if you're having issues,
        please see the tidystats{" "}
        <a
          href="https://www.tidystats.io/support.html"
          target="_blank"
          rel="noreferrer"
        >
          support
        </a>{" "}
        page.
      </p>
      <p>Click the button below to reset the help messages.</p>
      <PrimaryButton onClick={handleResetHelpClick}>
        Reset help messages
      </PrimaryButton>
    </>
  )
}
