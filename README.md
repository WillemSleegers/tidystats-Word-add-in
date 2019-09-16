# tidystats Word add-in

The tidystats Word add-in is an add-in for Microsoft Word to insert statistics from a file created with the [tidystats](https://github.com/WillemSleegers/tidystats) R package. 

**Please note that the add-in is currently in development**. This means the package may contain bugs and is subject to significant changes. It also means I can use your help in finding these bugs! Please see the 'Testing' section below to see how you could test out the add-in and how you can report bugs or feature requests.

## Requirements

You need the following to use add-in:
- An internet connection (the add-in is a web app that runs inside of Word)
- A compatible version of Microsoft Word, this is either:
	- (Mac) Microsoft Word version 15.18 or higher
	- (iPadOS) Microsoft Word for iPad
	- (Windows) Microsoft Word 2013 (or higher)

For more information on the requirements, see [here](https://docs.microsoft.com/en-us/office/dev/add-ins/concepts/requirements-for-running-office-add-ins).

## Installation

At this moment, it is only possible to install the add-in via a so-called sideloading method (see the 'Testing' section for details). Once the add-in is out of the testing phase, the add-in will be made available in the Office Add-ins Store

## Usage

The tidystats Word add-in is designed to accept a .JSON file created with the tidystats R package. Please see the following links for more information on how to use the R package to create this file:

- [tidystats' Github page](https://github.com/WillemSleegers/tidystats)
- [www.willemsleegers.com/tidystats](https://www.willemsleegers.com/tidystats.html)

Once you have created a JSON file containing the output of statistical tests, you can use the add-in to insert the statistics in a Word document. The add-in can be found in the Insert pane in Word. Click on the 'Insert Statistics' button, which will open a taskpane on the right side of the document. Next, click on the 'Choose File' button and select the JSON file. You will now see a list of statistical analyses, represented by their identifiers. Click on an identifier to look at the statistics of the analysis and click on either a specific statistic or on the bold text to insert an entire APA line of statistics. 

If you later change the output contained within the JSON file, you can update the numbers in the document by reading in the new file and clicking on the 'Update' button in the Actions section below the list of analyses.

## Testing

As mentioned before, the add-in is still under development. This means I can use your help with identifying bugs and feedbacking the add-in before I publish the add-in for all to use. 

You can help by performing the following three steps:
1. Install the add-in
2. Play around with the add-in (try to break it)
3. Post bugs/feature requests here on Github

### Step 1

First, download the manifest.xml file here from Github. This is the file that needs to be installed on your computer somewhere in order to run the add-in.

#### Mac

To install the add-in in Mac, see the 'Sideload an add-in in Office on Mac' section [here](Sideload an add-in in Office on Mac). Briefly put, you need to put the manifest file in the following folder:

> /Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef

If the 'wef' folder does not exist yet, simply create it and put the manifest file inside. After doing that, you can access the add-in via Insert > Add-ins > My Add-ins (click on the dropdown arrow). You should see the tidystats add-in in the list. Simply click on it and the add-in will open.

That's it for Mac <3. 

#### Windows

To install the add-in in Windows, see [this](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) webpage. When you copy the path to the shared folder, make sure it is a valid path (When I copied the path in step 5 of 'Share a folder' it also added some extra text to the front and end of the path).

### Step 2

Click on the 'Insert Statistics' button located in the Insert pane of the Word document. After opening the add-in, click on 'Choose File' and select the JSON file you created with the `tidystats` R package. If that succesfully loads, you should see a list of analyses and you can start playing with inserting statistics. If you hover over certain elements in the tables containing statistics, you will see that some elements are clickable (an underline appears). Click on these elements to insert statistics.

Once you have inserted some statistics, also try out the 'Update' button. Change the analyses you ran previously, create a new JSON file, and update the statistics after loading in the new file.

### Step 3

Found a bug or got an idea about how to improve the add-in? Create an issue here on Github! If you need some help figuring out how this works, see this support [video](https://help.github.com/en/articles/creating-an-issue) by Github or simply contact me on Twitter via @willemsleegers.