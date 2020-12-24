// Global variables ------------------------------------------------------------

var analyses;

// Setup -----------------------------------------------------------------------

Office.onReady(function (info) {
  console.log("Office ready");
  // Check if a Word application is running; if not, show that tidystats failed
  // to load
  if (info.host === Office.HostType.Word) {
    document.getElementById("loading-message").style.display = "none";

    document.getElementById("file-input").onchange = readFile;
    document.getElementById("file-input").onclick = resetFile;

    document.getElementById("search").onkeyup = search;

    document.getElementById("update-button").onclick = updateStatistics;
    document.getElementById("cite1").onclick = insertInTextCitation;
    document.getElementById("cite2").onclick = insertFullCitation;
    document.getElementById("cite3").onclick = copyBib;

    // Make the file input section visible
    document.getElementById("app-input").style.display = "block";

    // test();
    // Check settings to see whether a file was already opened
    //console.log("Loaded add-in");
    //var file = Office.context.document.settings.get('file_name');
    //var saved_data = Office.context.document.settings.get('data');

    //if (instantly_load_data) {
    //  instantlyLoadData();
    //}
  } else {
    // If the add-in fails to load, remove the loading message and show the failed to load message
    document.getElementById("loading-message").style.display = "none";
    document.getElementById("fail-message").style.display = "block";
  }
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
    console.log(attrs);
    var output = createStatisticsOutput(analyses, attrs);

    // Create a context control
    var doc = context.document;
    var selection = doc.getSelection();
    var selection_font = selection.font;
    selection_font.load("name");
    selection_font.load("size");

    var content_control = selection.insertContentControl();

    // Set analysis specific information
    content_control.tag = stringifyAttributes(attrs);
    content_control.insertHtml(output, Word.InsertLocation.end);

    return context.sync().then(function () {
      // Match the font and font size to the selection
      content_control.font.name = selection_font.name;
      content_control.font.size = selection_font.size;
    });
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
  return Word.run(function (context) {
    // Get all the content controls
    var content_controls = context.document.contentControls;

    // Sync context and loop over all content controls
    context.load(content_controls, "items");
    return context.sync().then(function () {
      var items = content_controls.items;

      for (var item in items) {
        // Extract the tag from the content control
        var content_control = content_controls.items[item];
        var tag = content_control.tag;

        // Use the tag to identify the reported analysis
        var attrs = JSON.parse(tag);
        var output = createStatisticsOutput(analyses, attrs);

        // Set the new output
        content_control.insertHtml(output, "Replace");
      }
    });
  });
}

function insertInTextCitation() {
  Word.run(function (context) {
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("(Sleegers, 2020)", "End");

    return context.sync();

    // Create a context control
    //var doc = context.document;
    //var selection = doc.getSelection();
    //var selection_font = selection.font;
    //selection_font.load("name");
    //selection_font.load("size");

    //return context.sync()
    //  .then(function () {
    //    // Match the font and font size to the selection
    //   content_control.font.name = selection_font.name;
    //    content_control.font.size = selection_font.size;
    //  });
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
      "Sleegers, W.W.A. (2020) tidystats: Combine output of statistical tests. R package version 0.4. https://CRAN.R-project.org/package=tidystats",
      "End"
    );

    return context.sync();

    // Create a context control
    //var doc = context.document;
    //var selection = doc.getSelection();
    //var selection_font = selection.font;
    //selection_font.load("name");
    //selection_font.load("size");

    //return context.sync()
    //  .then(function () {
    //    // Match the font and font size to the selection
    //   content_control.font.name = selection_font.name;
    //    content_control.font.size = selection_font.size;
    //  });
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function copyBib() {
  /* Get the text field */
  var text = `@software{sleegers2020,
      title = {tidystats: Reproducibly report statistics in {{Microsoft Word}}},
      author = {Sleegers, Willem W. A.},
      date = {2020},
      url = {https://doi.org/10.5281/zenodo.4065574},
      version = {1.0}
    }`;
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
    var model = analysis.models.find((x) => x.name == attrs.model);
    statistics = model.statistics;
  } else if ("effect" in attrs) {
    var effect = analysis.effects[attrs.effect];

    if ("group" in attrs) {
      var group = effect.groups.find((x) => x.name == attrs.group);

      if ("term" in attrs) {
        var term = group.terms.find((x) => x.name == attrs.term);
        statistics = term.statistics;
      } else if ("pair1" in attrs) {
        var pair = group.pairs.find(
          (x) => (x.names[0] == attrs.pair1) & (x.names[1] == attrs.pair2)
        );
        statistics = pair.statistics;
      } else {
        statistics = group.statistics;
      }
    } else if ("term" in attrs) {
      var term = effect.terms.find((x) => x.name == attrs.term);
      statistics = term.statistics;
    } else if ("pair1" in attrs) {
      var pair = effect.pairs.find(
        (x) => (x.names[0] == attrs.pair1) & (x.names[1] == attrs.pair2)
      );
      statistics = pair.statistics;
    } else if ("statistic" in attrs) {
      statistics = effect.statistics;
    }
  } else if ("group" in attrs) {
    var group = analysis.groups.find((x) => x.name == attrs.group);

    if ("term" in attrs) {
      var term = group.terms.find((x) => x.name == attrs.term);
      statistics = term.statistics;
    } else {
      statistics = group.statistics;
    }
  } else if ("term" in attrs) {
    var term = analysis.terms.find((x) => x.name == attrs.term);
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
  console.log(statistics);
  console.log(statistic);

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

  return output;
}

function createStatisticsLine(statistics, selectedStatistics) {
  var name, value, text;
  var output = [];

  for (i in selectedStatistics) {
    name = selectedStatistics[i];
    console.log(name);

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

        if (selectedStatistics.includes("df")) {
          var df = statistics.df;
          text =
            "<i>" +
            formatName(name) +
            "</i>(" +
            formatNumber(df, "df") +
            ") = " +
            formatNumber(value, name);
        } else if (selectedStatistics.includes("df_numerator")) {
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
        } else {
          text =
            "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
        }

        output.push(text);
        break;
      case "df":
        if (selectedStatistics.includes("statistic")) {
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
        if (selectedStatistics.includes("statistic")) {
          break;
        } else {
          value = statistics[name];
          text =
            "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
          output.push(text);
          break;
        }
      case "df_denominator":
        if (selectedStatistics.includes("statistic")) {
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

        if (selectedStatistics.includes("error")) {
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

        if (selectedStatistics.includes("error")) {
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
          selectedStatistics.includes("BF_01") ||
          selectedStatistics.includes("BF_10")
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
        if (selectedStatistics.includes("pct")) {
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
        if (selectedStatistics.includes("n")) {
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
