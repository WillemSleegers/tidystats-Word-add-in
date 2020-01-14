# tidystats Word add-in

The tidystats Word add-in is an add-in for Microsoft Word to insert statistics from a file created with the [tidystats](https://github.com/WillemSleegers/tidystats) R package. 

**Please note that the add-in is currently in development**. This means that the information about how to install the add-in is incorrect. If you want to install the add-in, please see the Testing section below.

## Requirements

You need the following to use add-in:
- An internet connection (the add-in is a web app that runs inside of Word)
- A compatible version of Microsoft Word, this is either:
	- (Mac) Microsoft Word version 15.18 or higher
	- (iPadOS) Microsoft Word for iPad
	- (Windows) Microsoft Word 2013 (or higher)

For more information on the requirements, see [here](https://docs.microsoft.com/en-us/office/dev/add-ins/concepts/requirements-for-running-office-add-ins).

## Installation

Go to the Insert pane in your Word document and click on 'Store'. Search for tidystats and click on the 'Add' button to install the add-in. 

## Usage

Before you can use the tidystats Word add-in, you must create a file containing your statistics using the `tidystats` R package. For details on how to do this, please see the [website](https://willemsleegers.github.io/tidystats/) of `tidystats`.

If you simply want to try out the add-in, you can also use [this](assets/tidystats/results.json) example file.

Once you have a file created with `tidystats` and you want to insert the statistics into your document, go to the Insert pane of your Word document and click on 'Insert Statistics'. 

![](assets/screens/start.png)

Next, click on 'Choose a file' and select the file you created with `tidystats`. 

![](assets/screens/choose_file.png)

This should reveal a list of your analyses, each with a name that identifies your analyses. You can use the Search textbox to search for your analyses. Click on one of the analyses to reveal its statistics. You can then click on the name of a statistic or set of statistics to insert them into your document.

![](assets/screens/insert.png)

If you want to update the statistics, simply choose a new file and click on the 'Update statistics' button. Note that this does require that your analyses have the same identifier as in the previous file.

Finally, if you found the add-in useful, please cite the work. You can use the citation buttons to quickly insert a citation; thanks!

## Supported statistical tests

The following statistical tests, in R, are supported:

Package: **stats**
- `t.test()`
- `cor.test()`
- `chisq.test()`
- `wilcox.test()`
- `fisher.test()`
- `oneway.test()`
- `aov()`
- `lm()`

## More resources

Please see the following links for more information on how to use `tidystats`.

- [tidystats website](https://willemsleegers.github.io/tidystats/)
- [www.willemsleegers.com/tidystats](https://www.willemsleegers.com/tidystats.html)

## Testing

As mentioned before, the add-in is still under development. This means I can use your help with identifying bugs and feedbacking the add-in before I publish the add-in for all to use. 

You can help by installing the add-in (see below) and posting bugs/feature requests here on Github.

### Installing the development version of tidystats for Word

Download the manifest.xml file here from Github. This is the file that needs to be installed on your computer somewhere in order to run the add-in.

#### Mac

To install the add-in in Mac, see the 'Sideload an add-in in Office on Mac' section [here](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac#sideload-an-add-in-in-office-on-mac). Briefly put, you need to put the manifest file in the following folder:

> /Users/\<username\>/Library/Containers/com.microsoft.Word/Data/Documents/wef

If the 'wef' folder does not exist yet, simply create it and put the manifest file inside. After doing that, you can access the add-in via Insert ribbon > Add-ins > My Add-ins (click on the dropdown arrow; only then will you see the Developer add-ins). You should see the tidystats add-in in the list. Simply click on it and the add-in will open.

That's it for Mac <3. 

#### Windows

To install the add-in in Windows, see [this](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) webpage. When you copy the path to the shared folder, make sure it is a valid path (When I copied the path in step 5 of 'Share a folder' it also added some extra text to the front and end of the path).

### Bugs and feature requests

Found a bug or got an idea about how to improve the add-in? Create an issue here on Github! If you need some help figuring out how this works, see this support [page](https://help.github.com/en/articles/creating-an-issue) by Github or simply contact me on Twitter via @willemsleegers.