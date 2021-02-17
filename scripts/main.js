// Global variables ------------------------------------------------------------

var analyses;
var inWord;

// Setup -----------------------------------------------------------------------

Office.onReady(function (info) {
  console.log("Office ready");

  // Make the file input section visible
  document.getElementById("app-input").style.display = "block";

  // Set functions
  document.getElementById("file-input").onchange = readFile;
  document.getElementById("file-input").onclick = resetFile;
  document.getElementById("search").onkeyup = search;
  document.getElementById("cite3").onclick = copyBib;

  // Check if a Word application is running
  inWord = info.host === Office.HostType.Word;

  // If so, hide the warning message, make the insert buttons visible
  if (inWord) {
    // Hide warning
    document.getElementById("word-warning").style.display = "none";

    // Set insert button functions
    document.getElementById("cite1").onclick = insertInTextCitation;
    document.getElementById("cite2").onclick = insertFullCitation;

    // Set update button function
    document.getElementById("update-button").onclick = updateStatistics;

    // test();
    // Check settings to see whether a file was already opened
    //console.log("Loaded add-in");
    //var file = Office.context.document.settings.get('file_name');
    //var saved_data = Office.context.document.settings.get('data');

    //if (instantly_load_data) {
    //  instantlyLoadData();
    //}
  } else {
    document.getElementById("word-warning").style.display = "flex";
    document.getElementById("loading-message").style.display = "none";
    document.getElementById("update").style.display = "none";
    document.getElementById("cite1").style.display = "none";
    document.getElementById("cite2").style.display = "none";
  }

  // Remove the loading message
  document.getElementById("loading-message").style.display = "none";
});

// Read .json file function ----------------------------------------------------

function readFile() {
  var fileInput, file, label, reader;

  fileInput = document.getElementById("file-input");
  file = fileInput.files[0];

  reader = new FileReader();
  reader.onload = function (e) {
    var text = reader.result;
    analyses = JSON.parse(text);

    // saveData(file.name, analyses);
    createAnalyses(analyses);
  };

  reader.readAsText(file);

  // Update the input label to reflect the file name
  label = document.getElementById("file-label");
  label.innerHTML = "File: " + file.name;
}

function resetFile() {
  this.value = null;
}

function saveData(file_name, statistics) {
  Office.context.document.settings.set("file_name", file_name);
  Office.context.document.settings.set("analyses", statistics);

  Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      console.log("Settings save failed. Error: " + asyncResult.error.message);
    } else {
      console.log("Settings saved.");
    }
  });
}

// Search related functions ----------------------------------------------------

function search() {
  var input, analyses, i, analysis, identifier;

  // Get the input
  input = document.getElementById("search");
  input = input.value.toUpperCase();

  // Get the analyses
  analyses = document.getElementsByClassName("identifier");

  // Loop over the analyses and hide them if the identifier does not match the
  // input
  for (i = 0; i < analyses.length; i++) {
    analysis = analyses[i].parentElement;

    identifier = analysis.getElementsByClassName("identifier-label");
    identifier = identifier[0].innerText;

    if (identifier.toUpperCase().indexOf(input) > -1) {
      analysis.style.display = "";
    } else {
      analysis.style.display = "none";
    }
  }
}

// Word functions --------------------------------------------------------------

function insert(attrs) {
  Word.run(function (context) {
    // Determine output
    var output = createStatisticsOutput(analyses, attrs);

    // Get the selection and load the font
    var doc = context.document;
    var selection = doc.getSelection();
    selection.load("font");

    return context
      .sync()
      .then(function () {
        // Create a content control
        var contentControl = selection.insertContentControl();

        // Set font and font size
        contentControl.font.name = selection.font.name;
        contentControl.font.size = selection.font.size;

        // Set analysis information
        contentControl.tag = stringifyAttributes(attrs);
        contentControl.insertHtml(output, Word.InsertLocation.replace);

        // Set cursor to the end of the selection
        selection.select(Word.InsertLocation.end);
      })
      .then(context.sync);
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function stringifyAttributes(attrs) {
  var string = JSON.stringify(attrs);

  return string;
}

function updateStatistics() {
  Word.run(function (context) {
    // Get all the content controls
    var contentControls = context.document.contentControls;

    // Sync context and loop over all content controls
    context.load(contentControls, "items");
    return context.sync().then(function () {
      var items = contentControls.items;

      for (var item in items) {
        // Extract the tag from the content control
        var contentControl = contentControls.items[item];
        var tag = contentControl.tag;

        // Use the tag to identify the reported analysis
        var attrs = JSON.parse(tag);

        // Try to get the statistics, it may be that statistics were previously reported that are not in the new file
        try {
          var output = createStatisticsOutput(analyses, attrs);

          // Set the new output
          contentControl.insertHtml(output, "Replace");
        } catch (err) {
          console.log(err.message);
        }
      }
    });
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function insertInTextCitation() {
  Word.run(function (context) {
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("(Sleegers, 2020)", Word.InsertLocation.end);

    // Set cursor to the end of the selection
    originalRange.select(Word.InsertLocation.end);

    return context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function insertFullCitation() {
  Word.run(function (context) {
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText(
      "Sleegers, W. W. A. (2020). tidystats: Save output of statistical tests (Version 0.5) [Computer software]. https://doi.org/10.5281/zenodo.4041859",
      Word.InsertLocation.end
    );

    // Set cursor to the end of the selection
    originalRange.select(Word.InsertLocation.end);

    return context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function copyBib() {
  /* Get the text field */
  var text =
    "@software{sleegers2020, title = {tidystats: Save output of statistical tests}, author = {Sleegers, Willem W. A.}, date = {2020}, url = {https://doi.org/10.5281/zenodo.4041859}, version = {0.5}}";
  var textArea = document.createElement("textarea");

  // Place in top-left corner of screen regardless of scroll position.
  textArea.style.position = "fixed";
  textArea.style.top = 0;
  textArea.style.left = 0;

  // Ensure it has a small width and height. Setting to 1px / 1em
  // doesn't work as this gives a negative w/h on some browsers.
  textArea.style.width = "2em";
  textArea.style.height = "2em";

  // We don't need padding, reducing the size if it does flash render.
  textArea.style.padding = 0;

  // Clean up any borders.
  textArea.style.border = "none";
  textArea.style.outline = "none";
  textArea.style.boxShadow = "none";

  // Avoid flash of white box if rendered for any reason.
  textArea.style.background = "transparent";
  textArea.value = text;

  document.body.appendChild(textArea);
  textArea.focus();
  textArea.select();

  try {
    var successful = document.execCommand("copy");
    var msg = successful ? "successful" : "unsuccessful";
    console.log("Copying text command was " + msg);
    document.getElementById("cite3").children[0].innerHTML = "Copied!";
    setTimeout(function () {
      document.getElementById("cite3").children[0].innerHTML = "Copy BibTeX";
    }, 3000);
  } catch (err) {
    console.log("Oops, unable to copy");
  }

  document.body.removeChild(textArea);
}

// Statistics retrieval/formatting functions -----------------------------------

function createStatisticsOutput(data, attrs) {
  var identifier, analysis, method, single, statistics;

  // Get the identifier from the attributes
  identifier = attrs["identifier"];

  // Use the identifier to extract the specific analysis and what kind of
  // method it is
  analysis = data[identifier];
  method = analysis.method;

  // Determine the statistics that should be inserted
  if ("model" in attrs) {
    var model = analysis.models.find(function (x) {
      return x.name == attrs.model;
    });
    statistics = model.statistics;
  } else if ("effect" in attrs) {
    var effect = analysis.effects[attrs.effect];

    if ("group" in attrs) {
      var group = effect.groups.find(function (x) {
        return x.name == attrs.group;
      });

      if ("term" in attrs) {
        var term = group.terms.find(function (x) {
          return x.name == attrs.term;
        });
        statistics = term.statistics;
      } else if ("pair1" in attrs) {
        var pair = group.pairs.find(function (x) {
          return (x.names[0] == attrs.pair1) & (x.names[1] == attrs.pair2);
        });
        statistics = pair.statistics;
      } else {
        statistics = group.statistics;
      }
    } else if ("term" in attrs) {
      var term = effect.terms.find(function (x) {
        return x.name == attrs.term;
      });
      statistics = term.statistics;
    } else if ("pair1" in attrs) {
      var pair = effect.pairs.find(function (x) {
        return (x.names[0] == attrs.pair1) & (x.names[1] == attrs.pair2);
      });
      statistics = pair.statistics;
    } else if ("statistic" in attrs) {
      statistics = effect.statistics;
    }
  } else if ("group" in attrs) {
    var group = analysis.groups.find(function (x) {
      return x.name == attrs.group;
    });

    if ("term" in attrs) {
      var term = group.terms.find(function (x) {
        return x.name == attrs.term;
      });
      statistics = term.statistics;
    } else {
      statistics = group.statistics;
    }
  } else if ("term" in attrs) {
    var term = analysis.terms.find(function (x) {
      return x.name == attrs.term;
    });
    statistics = term.statistics;
  } else {
    statistics = analysis.statistics;
  }

  // Determine whether a single statistic or multiple statistics should be inserted
  if ("statistic" in attrs) {
    single = true;
  } else {
    single = false;
  }

  // Create a variable to store the output in
  var output;

  // If single, report only a single statistic, otherwise a line of statistics
  if (single) {
    output = retrieveStatistic(statistics, attrs.statistic);
  } else {
    var selectedStatistics = attrs.statistics;
    output = createStatisticsLine(statistics, selectedStatistics);
  }
  return output;
}

function retrieveStatistic(statistics, statistic) {
  console.log("Retrieving statistics");

  var output;

  if (statistic == "estimate") {
    output = formatNumber(statistics.estimate.value, statistics.estimate.name);
  } else if (statistic == "statistic") {
    output = formatNumber(statistics.statistic.value);
  } else if (statistic == "df_numerator") {
    output = formatNumber(statistics.dfs.df_numerator, statistic);
  } else if (statistic == "df_denominator") {
    output = formatNumber(statistics.dfs.df_denominator, statistic);
  } else if (statistic == "df_null") {
    output = formatNumber(statistics.dfs.df_null, statistic);
  } else if (statistic == "df_residual") {
    output = formatNumber(statistics.dfs.df_residual, statistic);
  } else if (statistic == "CI_lower") {
    output = formatNumber(statistics.CI.CI_lower, statistic);
  } else if (statistic == "CI_upper") {
    output = formatNumber(statistics.CI.CI_upper, statistic);
  } else {
    output = formatNumber(statistics[statistic], statistic);
  }

  console.log(output);

  return output;
}

function createStatisticsLine(statistics, selectedStatistics) {
  console.log("Creating statistics line");

  var name, value, text;
  var output = [];

  for (i in selectedStatistics) {
    name = selectedStatistics[i];

    switch (name) {
      case "estimate":
        name = statistics[selectedStatistics[i]].name;
        value = statistics[selectedStatistics[i]].value;
        text = "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
        output.push(text);
        break;
      case "statistic":
        name = statistics[selectedStatistics[i]].name;
        value = statistics[selectedStatistics[i]].value;

        if (selectedStatistics.indexOf("df") > -1) {
          var df = statistics.df;
          text =
            "<i>" +
            formatName(name) +
            "</i>(" +
            formatNumber(df, "df") +
            ") = " +
            formatNumber(value, name);
        } else if (selectedStatistics.indexOf("df_numerator") > -1) {
          var df_numerator = statistics.dfs.df_numerator;
          var df_denominator = statistics.dfs.df_denominator;
          text =
            "<i>" +
            formatName(name) +
            "</i>(" +
            formatNumber(df_numerator, "df") +
            ", " +
            formatNumber(df_denominator, "df") +
            ") = " +
            formatNumber(value, name);
          // Check if the statistic is a Cohen's d or Hedges' g
        } else if (name == "d") {
          text = "Cohen's <i>d</i> = " + formatNumber(value, name);
        } else if (name == "g") {
          text = "Hedges' <i>g</i> = " + formatNumber(value, name);
        } else {
          text =
            "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
        }
        output.push(text);
        break;
      case "df":
        if (selectedStatistics.indexOf("statistic") > -1) {
          break;
        } else {
          if ("dfs" in statistics) {
            value = statistics.dfs[name];
          } else {
            value = statistics[name];
          }
          text =
            "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
          output.push(text);
          break;
        }
      case "df_numerator":
        if (selectedStatistics.indexOf("statistic") > -1) {
          break;
        } else {
          value = statistics[name];
          text =
            "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
          output.push(text);
          break;
        }
      case "df_denominator":
        if (selectedStatistics.indexOf("statistic") > -1) {
          break;
        } else {
          value = statistics[name];
          text =
            "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
          output.push(text);
          break;
        }
      case "df_residual":
        value = statistics.dfs[name];
        text = "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
        output.push(text);
        break;
      case "p":
        value = statistics[name];
        if (value < 0.001) {
          text = "<i>" + formatName(name) + "</i> " + formatNumber(value, name);
        } else {
          text =
            "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
        }
        output.push(text);
        break;
      case "CI_lower":
        text = createCILine(
          statistics.CI.CI_lower,
          statistics.CI.CI_upper,
          statistics.CI.CI_level
        );
        output.push(text);
        break;
      case "CI_upper":
        break;
      case "BF_01":
        value = statistics[name];

        if (selectedStatistics.indexOf("error") > -1) {
          var error = statistics["error"];
          text =
            formatName(name) +
            " = " +
            formatNumber(value, name) +
            " ±" +
            formatNumber(error) +
            "%";
        } else {
          value = statistics[name];
          text = formatName(name) + " = " + formatNumber(value, name);
        }
        output.push(text);
        break;
      case "BF_10":
        value = statistics[name];

        if (selectedStatistics.indexOf("error") > -1) {
          var error = statistics["error"];
          text =
            formatName(name) +
            " = " +
            formatNumber(value, name) +
            " ±" +
            formatNumber(error) +
            "%";
        } else {
          value = statistics[name];
          text = formatName(name) + " = " + formatNumber(value, name);
        }
        output.push(text);
        break;
      case "error":
        if (
          selectedStatistics.indexOf("BF_01") > -1 ||
          selectedStatistics.indexOf("BF_10") > -1
        ) {
          break;
        } else {
          value = statistics[name];
          text =
            "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
          output.push(text);
        }
        break;
      case "n":
        if (selectedStatistics.indexOf("pct") > -1) {
          var pct = statistics.pct;
          value = statistics[name];
          text =
            "<i>" +
            formatName(name) +
            "</i> = " +
            formatNumber(value, name) +
            " (" +
            formatNumber(pct, "pct") +
            "%)";
        } else {
          value = statistics[name];
          text =
            "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
        }
        output.push(text);
        break;
      case "pct":
        if (selectedStatistics.indexOf("n") > -1) {
          break;
        } else {
          value = statistics[name];
          text = formatNumber(value, name) + "%";
          output.push(text);
        }
        break;
      default:
        value = statistics[name];
        text = "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
        output.push(text);
    }
  }

  output = output.join(", ");

  return output;
}

function createCILine(lower, upper, level) {
  return (
    level * 100 +
    "% CI [" +
    formatNumber(lower) +
    ", " +
    formatNumber(upper) +
    "]"
  );
}
