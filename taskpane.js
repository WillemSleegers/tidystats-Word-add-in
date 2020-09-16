
// Global variables ------------------------------------------------------------

var analyses;

// Setup -----------------------------------------------------------------------

Office.onReady(function(info) {
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

  fileInput = document.getElementById('file-input');
  file = fileInput.files[0];
  
  reader = new FileReader();
  reader.onload = function(e) {
    var text = reader.result;
    analyses = JSON.parse(text);
        
    // saveData(file.name, analyses);
    createAnalyses(analyses);
  }

  reader.readAsText(file);

  // Update the input label to reflect the file name
  label = document.getElementById("file-label");
  label.innerHTML = 'File: ' + file.name;
}

function resetFile() {
  this.value = null;
}

function saveData(file_name, statistics) {
  Office.context.document.settings.set('file_name', file_name);
  Office.context.document.settings.set('analyses', statistics);

  Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        console.log('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        console.log('Settings saved.');
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

function insertStatistics() {
  console.log("Inserting statistics");

  // Obtain the attributes of the button that was clicked on
  attrs = getAttributes(this);
  console.log(attrs);
  
  Word.run(function (context) {
    // Determine output
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
    content_control.insertHtml(output, "Replace");
    
    return context.sync()
      .then(function () {
        // Match the font and font size to the selection
        content_control.font.name = selection_font.name;
        content_control.font.size = selection_font.size;
      });
  })
  .catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function getAttributes(element) {
  var attributes = {};

  attributes["identifier"] = element.getAttribute("identifier");
  attributes["single"] = element.getAttribute("single");

  if (element.hasAttribute("statistic")) {
    attributes["statistic"] = element.getAttribute("statistic");
  }
  if (element.hasAttribute("term")) {
    attributes["term"] = element.getAttribute("term");
  }
  if (element.hasAttribute("pair1")) {
    attributes["pair1"] = element.getAttribute("pair1");
    attributes["pair2"] = element.getAttribute("pair2");
  }
  if (element.hasAttribute("group")) {
    attributes["group"] = element.getAttribute("group");
  }
  if (element.hasAttribute("effect")) {
    attributes["effect"] = element.getAttribute("effect");
  }
  if (element.hasAttribute("model")) {
    attributes["model"] = element.getAttribute("model");
  }

  return(attributes)
}

function stringifyAttributes(attrs) {
  var string = JSON.stringify(attrs)

  return(string)
}

function updateStatistics() {

  return Word.run( function(context) {
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
  })
  .catch(function (error) {
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
    originalRange.insertText("Sleegers, W.W.A. (2020) tidystats: Combine output of statistical tests. R package version 0.4. https://CRAN.R-project.org/package=tidystats", "End");

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
  })
  .catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

// Statistics retrieval/formatting functions -----------------------------------

function createStatisticsOutput(data, attrs) {
    var identifier, analysis, method, single, statistics, statistic;

    // Get the identifier from the attributes
    identifier = attrs["identifier"];

    // Use the identifier to extract the specific analysis and what kind of 
    // method it is
    analysis = data[identifier];
    method = analysis.method; 

    // Get the single attribute to determine whether a single statistic should
    // be inserted or a whole line of statistics
    single = attrs["single"];

    // Create a variable to store the output in
    var output;

   if (/t-test/.test(method)) {
      statistics = analysis.statistics;

      if (single == "true") {
        output = retrieveStatistic(statistics, attrs["statistic"]);
      } else {
        output = createTTestLine(statistics);
      }
    } else if (/Pearson's product-moment correlation/.test(method)) {
      statistics = analysis.statistics;

      if (single == "true") {
        output = retrieveStatistic(statistics, attrs["statistic"]);
      } else {
        output = createPearsonCorrelationLine(statistics);
      }
    } else if (/Kendall's rank correlation tau/.test(method)) {
      statistics = analysis.statistics;

      if (single == "true") {
        output = retrieveStatistic(statistics, attrs["statistic"]);
      } else {
        output = createKendallCorrelationLine(statistics);
      }
    } else if (/Spearman's rank correlation rho/.test(method)) {
      statistics = analysis.statistics;

      if (single == "true") {
        output = retrieveStatistic(statistics, attrs["statistic"]);
      } else {
        output = createSpearmanCorrelationLine(statistics);
      }
    } else if (/Chi-squared test/.test(method)) {
      statistics = analysis.statistics;

      if (single == "true") {
        output = retrieveStatistic(statistics, attrs["statistic"]);
      } else {
        output = createChiSquaredLine(statistics);
      }
    } else if (/Wilcoxon/.test(method)) {
      statistics = analysis.statistics;

      if (single == "true") {
        output = retrieveStatistic(statistics, attrs["statistic"]);
      } else {
        output = createWilcoxonTestLine(statistics);
      }
    } else if (/Fisher's Exact Test/.test(method)) {
      statistics = analysis.statistics;

      if (single == "true") {
        output = retrieveStatistic(statistics, attrs["statistic"]);
      } else {
        output = createFisherExactTestLine(statistics);
      }
    } else if (/One-way analysis of means/.test(method)) {
      statistics = analysis.statistics;

      if (single == "true") {
        output = retrieveStatistic(statistics, attrs["statistic"]);
      } else {
        output = createOneWayAnalysisOfMeansLine(statistics);
      }
    } else if (/Linear regression/.test(method)) {
      if ("term" in attrs) {
        var term = analysis.terms.find(x => x.name == attrs["term"]);
        statistics = term.statistics;        

        if (single == "true") {
          output = retrieveStatistic(statistics, attrs["statistic"]);
        } else {
          output = createLinearModelTermLine(statistics);
        } 
      } else {
        statistics = analysis.statistics;

        if (single == "true") {
          output = retrieveStatistic(statistics, attrs["statistic"]);
        } else {
          output = createLinearModelModelLine(statistics);
        }
      }
    } else if (/ANOVA/.test(method)) {
      if ("group" in attrs) {
        var group = analysis.groups.find(x => x.name == attrs["group"]);
        var terms = group.terms;
      } else {
        var terms = analysis.terms;
      }

      var term = terms.find(x => x.name == attrs["term"]);  
      statistics = term.statistics;

      if (single == "true") {
        output = retrieveStatistic(statistics, attrs["statistic"]);
      } else {
        var term_residuals = terms.find(x => x.name == "Residuals");
        statistics_residuals = term_residuals.statistics;

        output = createANOVALine(statistics, statistics_residuals);
      } 
    } else if (/Descriptives/.test(method)) {
      if ("group" in attrs) {
        var group = analysis.groups.find(x => x.name == attrs["group"]);
        var statistics = group.statistics;
      } else {
        var statistics = analysis.statistics;
      }

      if (single == "true") {
          output = retrieveStatistic(statistics, attrs["statistic"]);
      } else {
          output = createDescriptivesLine(statistics);
      }
    } else if (/Counts/.test(method)) {
      var group = analysis.groups.find(x => x.name == attrs["group"]);
      var statistics = group.statistics;

      if (single == "true") {
          output = retrieveStatistic(statistics, attrs["statistic"]);
      } else {
          output = createCountsLine(statistics);
      }
    } else if (/Linear mixed model/.test(method)) {
      if (attrs["effect"] == "random_effect") {
        if ("group" in attrs) {
          var group = analysis.effects.random_effects.groups.find(x => 
            x.name == attrs["group"]);
          if ("term" in attrs) {
            var term = group.terms.find(x => x.name == attrs["term"]);
            var statistics = term.statistics;
          } else if ("pair1" in attrs) {
            var pair = group.pairs.find(x => x.names[0] == attrs["pair1"]);
            var statistics = pair.statistics;
          } else {
            var statistics = group.statistics;
          }  
        } else {
          var statistics = analysis.effects.random_effects.statistics;
        }
      } else {
        if ("term" in attrs) {
          var term = analysis.effects.fixed_effects.terms.find(x => 
            x.name == attrs["term"]);
          var statistics = term.statistics;

          if (single == "false") {
            output = createLinearMixedModelFixedEffectLine(statistics);
          }
        } else {
          var pair = analysis.effects.fixed_effects.pairs.find(x => 
            x.names[0] == attrs["pair1"]);
          var statistics = pair.statistics;
        }
      }

      if (single == "true") {
        output = retrieveStatistic(statistics, attrs["statistic"]);
      }
    } else if (/Generic/.test(method)) { 
      var statistics = analysis.statistics;

      if (single == "true") {
        output = retrieveStatistic(statistics, attrs["statistic"]);
      }
    } else {
      output = "Sorry, not supported";
    }

    return(output)
}

function retrieveStatistic(statistics, statistic) {

  var output;

  if (statistic == "estimate") {
    output = formatNumber(statistics.estimate.value, statistics.estimate.name)
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

  return(output)
}

// APA lines of statistics functions -------------------------------------------

function createTTestLine(statistics) {
  var t, df, p, output;
 
  t = formatNumber(statistics.statistic.value);
  df = formatNumber(statistics.df, "df");
  p = formatNumber(statistics.p, "p");

  if (p == "< .001") {
    output = "<i>t</i>(" + df + ") = " + t + ", <i>p</i> " + p;
  } else {
    output = "<i>t</i>(" + df + ") = " + t + ", <i>p</i> = " + p;
  }

  if ("CI" in statistics) {
    output = output + ", " + statistics.CI.CI_level * 100 + "% CI [" + 
      formatNumber(statistics.CI.CI_lower) + ", " + 
      formatNumber(statistics.CI.CI_upper) + "]"
  }

  return(output)
}

function createPearsonCorrelationLine(statistics) {
  var r, df, p, output;

  r = formatNumber(statistics.estimate.value);
  df = formatNumber(statistics.df, "df");
  p = formatNumber(statistics.p, "p");

  if (p == "< .001") {
    output = "<i>r</i>(" + df + ") = " + r + ", <i>p</i> " + p;
  } else {
    output = "<i>r</i>(" + df + ") = " + r + ", <i>p</i> = " + p;
  }

  if ("CI" in statistics) {
    output = output + ", " + statistics.CI.CI_level * 100 + "% CI [" + 
      formatNumber(statistics.CI.CI_lower) + ", " + 
      formatNumber(statistics.CI.CI_upper) + "]"
  }

  return(output) 
}

function createKendallCorrelationLine(statistics) {
  var tau, p, output;

  tau = formatNumber(statistics.estimate.value);
  p = formatNumber(statistics.p, "p");

  if (p == "< .001") {
    output = "<i>r<sub>&tau;</sub></i> = " + tau + ", <i>p</i> " + p;
  } else {
    output = "<i>r<sub>&tau;</sub></i> = " + tau + ", <i>p</i> = " + p;
  }

  return(output) 
}

function createSpearmanCorrelationLine(statistics) {
  var rho, p, output;

  rho = formatNumber(statistics.estimate.value);
  p = formatNumber(statistics.p, "p");

  if (p == "< .001") {
    output = "<i>r<sub>S</sub></i> = " + rho + ", <i>p</i> " + p;
  } else {
    output = "<i>r<sub>S</sub></i> = " + rho + ", <i>p</i> = " + p;
  }

  return(output) 
}

function createChiSquaredLine(statistics) {
  var chi_squared, df, p, output;

  chi_squared = formatNumber(statistics.statistic.value);
  df = formatNumber(statistics.df, "df");
  p = formatNumber(statistics.p, "p");

  if (p == "< .001") {
    output = "<i>&chi;²</i>(" + df + ") = " + chi_squared + ", <i>p</i> " + p;
  } else {
    output = "<i>&chi;²</i>(" + df + ") = " + chi_squared + ", <i>p</i> = " + p;
  }

  return(output) 
}

function createWilcoxonTestLine(statistics) {
  var statistic_name, statistic_value, p, output;

  statistic_name = statistics.statistic.name;
  statistic_value = formatNumber(statistics.statistic.value);
  p = formatNumber(statistics.p, "p");

  if (p == "< .001") {
    output = "<i>" + statistic_name + "</i> = " + statistic_value + 
      ", <i>p</i> " + p;
  } else {
    output = "<i>" + statistic_name + "</i> = " + statistic_value + 
      ", <i>p</i> = " + p;
  }

  return(output) 
}

function createFisherExactTestLine(statistics) {

  var output;
  var p = formatNumber(statistics.p, "p");
  
  if ("estimate" in statistics) {
    var OR = formatNumber(statistics.estimate.value);

    output = "<i>OR</i> = " + OR

    if (p == "< .001") {
      output = output + ", <i>p</i> " + p;
    } else {
      output = output + ", <i>p</i> = " + p;
    }
  } else {
    if (p == "< .001") {
      output = "<i>p</i> " + p;
    } else {
      output = "<i>p</i> = " + p;
    }
  }

  if ("CI" in statistics) {
    var lower = formatNumber(statistics.CI.lower);
    var upper = formatNumber(statistics.CI.upper);

    output = output + ", " + statistics.CI.level * 100 + "% CI [" + lower + 
      ", " + upper + "]"
  }

  return(output) 
}

function createOneWayAnalysisOfMeansLine(statistics) {
  var statistic, df1, df2, p, output;

  statistic = formatNumber(statistics.statistic.value);
  df1 = formatNumber(statistics.dfs.df_numerator, "df");
  df2 = formatNumber(statistics.dfs.df_denominator, "df");
  p = formatNumber(statistics.p, "p");

  if (p == "< .001") {
    output = "<i>F</i>(" + df1 + ", " + df2 + ") = " + statistic + 
      ", <i>p</i> " + p;
  } else {
    output = "<i>F</i>(" + df1 + ", " + df2 + ") = " + statistic + 
      ", <i>p</i> = " + p;
  }

  return(output) 
}

function createLinearModelTermLine(statistics) {
  var estimate, SE, statistic, df, p, output;

  estimate = formatNumber(statistics.estimate.value);
  SE = formatNumber(statistics.SE);
  statistic_name = statistics.statistic.name;
  statistic_value = formatNumber(statistics.statistic.value);
  df = formatNumber(statistics.df, "df");
  p = formatNumber(statistics.p, "p");

  if (p == "< .001") {
    output = "<i>b</i> = " + estimate + ", <i>SE</i> = " + SE + ", <i>" + 
      statistic_name + "</i>(" + df + ") = " + statistic_value + 
      ", <i>p</i> " + p;
  } else {
    output = "<i>b</i> = " + estimate + ", <i>SE</i> = " + SE + ", <i>" + 
      statistic_name + "</i>(" + df + ") = " + statistic_value + 
      ", <i>p</i> = " + p;
  }

  return(output) 
}

function createLinearModelModelLine(statistics) {
  var r_squared, statistic, df1, df2, p, output;

  r_squared = formatNumber(statistics.r_squared);
  statistic = formatNumber(statistics.statistic.value);
  df1 = formatNumber(statistics.dfs.df_numerator, "df");
  df2 = formatNumber(statistics.dfs.df_denominator, "df");
  p = formatNumber(statistics.p, "p");

  if (p == "< .001") {
    output = "<i>R²</i> = " + r_squared + ", <i>F</i>(" + df1 + ", " + df2 + 
      ") = " + statistic + ", <i>p</i> " + p;
  } else {
    output = "<i>R²</i> = " + r_squared + ", <i>F</i>(" + df1 + ", " + df2 + 
      ") = " + statistic + ", <i>p</i> = " + p;
  }

  return(output) 
}

function createANOVALine(statistics, statistics_residuals) {
  var statistic, df1, df2, p, output;

  statistic = formatNumber(statistics.statistic.value);

  if ("dfs" in statistics) {
    df1 = formatNumber(statistics.dfs.df_numerator, "df");
    df2 = formatNumber(statistics.dfs.df_denominator, "df");  
  } else {
    df1 = formatNumber(statistics.df, "df");
    df2 = formatNumber(statistics_residuals.df, "df");
  }

  p = formatNumber(statistics.p, "p");

  if (p == "< .001") {
    output = "<i>F</i>(" + df1 + ", " + df2 + ") = " + statistic + 
      ", <i>p</i> " + p;
  } else {
    output = "<i>F</i>(" + df1 + ", " + df2 + ") = " + statistic + 
      ", <i>p</i> = " + p;
  }

  if ("ges" in statistics) {
    var ges = formatNumber(statistics.ges);
    output = output + ", η²<sub>G</sub> = " + ges;
  }

  return(output) 
}

function createLinearMixedModelFixedEffectLine(statistics) {
  var b, SE, t, df, p, output;

  estimate = formatNumber(statistics.estimate.value);
  SE = formatNumber(statistics.SE);
  statistic = formatNumber(statistics.statistic.value);
  
  if ("df" in statistics) {
    df = formatNumber(statistics.df, "df");
    p = formatNumber(statistics.p, "p");

    if (p == "< .001") {
      output = "<i>b</i> = " + estimate + ", <i>SE</i> = " + SE + 
        ", <i>t</i>(" + df + ") = " + statistic + ", <i>p</i> " + p;
    } else {
      output = "<i>b</i> = " + estimate + ", <i>SE</i> = " + SE + 
        ", <i>t</i>(" + df + ") = " + statistic + ", <i>p</i> = " + p;
    }
  } else {
    if (p == "< .001") {
      output = "<i>b</i> = " + estimate + ", <i>SE</i> = " + SE + 
        ", <i>t</i> = " + statistic;
    } else {
      output = "<i>b</i> = " + estimate + ", <i>SE</i> = " + SE + 
        ", <i>t</i> = " + statistic;
    }
  }
  
  return(output) 
}

function createDescriptivesLine(statistics) {
  var M, SD, output;
 
  M = formatNumber(statistics.M);
  SD = formatNumber(statistics.SD);
 
  output = "(<i>M</i> = " + M + ", <i>SD</i> = " + SD + ")";

  return(output)
}

function createCountsLine(statistics) {
  var n, pct;
 
  n = formatNumber(statistics.n, type = "N");
  pct = formatNumber(statistics.pct);
 
  output = n + " (" + pct + "%)";

  return(output)
}

function createBayesFactorLine(statistics) {
  var BF, error, output;
 
  BF = formatNumber(statistics.BF_01);
  error = formatNumber(statistics.error);

  output = "<i>BF<sub>01</sub></i> = " + BF + " ± " + error + "%";
  
  return(output)
}

// Table creation functions ----------------------------------------------------


function createTTestWordTable (body, identifier, analysis) {
  var statistics = analysis.statistics;
  var data = [
    ["", "t", "df", "p", "95% CI"],
    ["name", "t-value", "df", "p-value", "CIs"]
  ];

  var table = body.insertTable(2, 5, Word.InsertLocation.end, data);

  // Format cell spacing
  table.verticalAlignment = "Center";
  table.setCellPadding("Top", 4);
  table.setCellPadding("Bottom", 4);
  
  // Format borders
  table.getBorder("Top").width = 1;
  table.getBorder("Bottom").width = 2;
  table.getBorder("InsideVertical").type = "None";
  table.getBorder("Left").type = "None";
  table.getBorder("Right").type = "None";
  
  // Format cell content
  table.getCell(0, 1).body.font.italic = true;
  table.getCell(0, 2).body.font.italic = true;
  table.getCell(0, 3).body.font.italic = true;

  var cell_name = table.getCell(1, 0);
  cell_name.value = analysis.data_name;

  var cell_t = table.getCell(1, 1);
  var cell_df = table.getCell(1, 2);
  var cell_p = table.getCell(1, 3);
  var cell_CI = table.getCell(1, 4);
  
  var content_control_t = cell_t.body.insertContentControl();
  var content_control_df = cell_df.body.insertContentControl();
  var content_control_p = cell_p.body.insertContentControl();
  var content_control_CI = cell_CI.body.insertContentControl();
  
  content_control_t.tag = identifier + "$t";
  content_control_t.title = identifier + "$t";
  content_control_t.insertHtml(formatNumber(statistics.statistic.value), "Replace");

  content_control_df.tag = identifier + "$df";
  content_control_df.title = identifier + "$df";
  content_control_df.insertHtml(formatNumber(statistics.df, "df"), "Replace");

  content_control_p.tag = identifier + "$p";
  content_control_p.title = identifier + "$p";
  content_control_p.insertHtml(formatNumber(statistics.p, "p"), "Replace");

  content_control_CI.tag = identifier + "$CI";
  content_control_CI.title = identifier + "$CI";
  content_control_CI.insertHtml(formatCIs(statistics.CI), "Replace");

  return(table)
}

function createRegressionWordTable(body, identifier, analysis) {
  var statistics = analysis.statistics;
  var data = [["", "b", "SE", "t", "df", "p"]];

  var coefficients = analysis.coefficients;

  for (i in coefficients) {
    data.push(["coefficient", "b-value", "SE-value", "t-value", "df-value", "p-value"]);
  }

  var table = body.insertTable(data.length, data[0].length, Word.InsertLocation.end, data);

  // Format cell spacing
  table.verticalAlignment = "Center";
  table.setCellPadding("Top", 4);
  table.setCellPadding("Bottom", 4);
  
  // Format borders
  table.getBorder("Top").width = 1;
  table.getBorder("Bottom").width = 2;
  table.getBorder("InsideVertical").type = "None";
  table.getBorder("InsideHorizontal").type = "None";
  table.getBorder("Left").type = "None";
  table.getBorder("Right").type = "None";

  var top_row_bottom_border = table.rows.getFirst().getBorder("Bottom");
  top_row_bottom_border.type = "Single";
  top_row_bottom_border.width = 1;

  // Format top row cells
  table.getCell(0, 1).body.font.italic = true;
  table.getCell(0, 2).body.font.italic = true;
  table.getCell(0, 3).body.font.italic = true;
  table.getCell(0, 4).body.font.italic = true;
  table.getCell(0, 5).body.font.italic = true;

  // Create a variable j to select the row of cells (using i throws an error)
  var j = 0;

  // Add content controls
  for (i in coefficients) {
    j++; 

    var coefficient = coefficients[i];
    var statistics = coefficient.statistics;

    var cell_name = table.getCell(j, 0);
    var cell_b = table.getCell(j, 1);
    var cell_SE = table.getCell(j, 2);
    var cell_t = table.getCell(j, 3);
    var cell_df = table.getCell(j, 4);
    var cell_p = table.getCell(j, 5);
    
    var name = coefficient["name"];
    cell_name.value = name;

    var content_control_b = cell_b.body.insertContentControl();
    var content_control_SE = cell_SE.body.insertContentControl();
    var content_control_t = cell_t.body.insertContentControl();
    var content_control_df = cell_df.body.insertContentControl();
    var content_control_p = cell_p.body.insertContentControl();
    
    content_control_b.tag = identifier + "$" + name + "$b";
    content_control_b.title = identifier + "$" + name + "$b";
    content_control_b.insertHtml(formatNumber(statistics.estimate), "Replace");

    content_control_SE.tag = identifier + "$" + name + "$SE";
    content_control_SE.title = identifier + "$" + name + "$SE";
    content_control_SE.insertHtml(formatNumber(statistics.SE), "Replace");

    content_control_t.tag = identifier + "$" + name + "$t";
    content_control_t.title = identifier + "$" + name + "$t";
    content_control_t.insertHtml(formatNumber(statistics.statistic.value), "Replace");

    content_control_df.tag = identifier + "$" + name + "$df";
    content_control_df.title = identifier + "$" + name + "$df";
    content_control_df.insertHtml(formatNumber(statistics.df, "df"), "Replace");

    content_control_p.tag = identifier + "$" + name + "$p";
    content_control_p.title = identifier + "$" + name + "$p";
    content_control_p.insertHtml(formatNumber(statistics.p, "p"), "Replace");
  }

  return(table)
}

// Debugging
var testAnalysis = '{"t_test_one_sample":{"method":"One Sample t-test", "description":"A simple t-test that I ran to test tidystats out.", "name":"cox$call_parent","statistics":{"estimate":25.775,"SE":1.064,"statistic":{"name":"t","value":24.2248},"df":199,"p":1.4581E-61,"CI":{"level":0.95,"lower":24.0167,"upper":"Inf"}},"alternative":{"direction":"greater","mean":0},"package":{"name":"stats","version":"3.6.1"},"notes":"A one-sample t-test on call_parent"}}'

function test() {
  var analysisDiv = document.getElementById("analyses-container");

  testAnalysis = JSON.parse(testAnalysis);

  var test2 = createAnalysis(testAnalysis, "t_test_one_sample");

  analysisDiv.appendChild(test2);
}


var instantly_load_data = false;

function instantlyLoadData() {
  console.log("Instantly loading data");
  
  var text = '{"m":{"method":"Linear mixed model","REML_criterion_at_convergence":1743.6283,"convergence_code":0,"random_effects":{"N":180,"groups":[{"name":"Subject","N":18,"variances":[{"name":"(Intercept)","statistics":{"var":611.8976,"SD":24.7366}},{"name":"Days","statistics":{"var":35.0811,"SD":5.9229}}],"correlations":[{"names":["(Intercept)","Days"],"statistics":{"r":0.0656}}]},{"name":"Residual","variances":[{"statistics":{"var":654.9408,"SD":25.5918}}]}]},"fixed_effects":{"coefficients":[{"name":"(Intercept)","statistics":{"estimate":251.4051,"SE":6.8238,"df":17.0052,"statistic":{"name":"t","value":36.8425},"p":1.1582E-17}},{"name":"Days","statistics":{"estimate":10.4673,"SE":1.546,"df":16.9953,"statistic":{"name":"t","value":6.7707},"p":3.273E-6}}],"correlations":[{"names":["(Intercept)","Days"],"statistics":{"r":-0.1375}}]},"package":{"name":"lme4","version":"1.1-21"}},"fm":{"method":"Linear mixed model","REML_criterion_at_convergence":2705.5037,"convergence_code":0,"random_effects":{"N":648,"groups":[{"name":"Consumer:Product","N":324,"variances":[{"name":"(Intercept)","statistics":{"var":3.1622,"SD":1.7783}}]},{"name":"Consumer","N":81,"variances":[{"name":"(Intercept)","statistics":{"var":0.3756,"SD":0.6129}}]},{"name":"Residual","variances":[{"statistics":{"var":1.6675,"SD":1.2913}}]}]},"fixed_effects":{"coefficients":[{"name":"(Intercept)","statistics":{"estimate":5.849,"SE":0.2843,"df":322.3361,"statistic":{"name":"t","value":20.5742},"p":1.1733E-60}},{"name":"Gender2","statistics":{"estimate":-0.2443,"SE":0.2606,"df":79,"statistic":{"name":"t","value":-0.9375},"p":0.3514}},{"name":"Information2","statistics":{"estimate":0.1605,"SE":0.2029,"df":320.0004,"statistic":{"name":"t","value":0.791},"p":0.4296}},{"name":"Product2","statistics":{"estimate":-0.8272,"SE":0.3453,"df":339.5108,"statistic":{"name":"t","value":-2.3953},"p":0.0171}},{"name":"Product3","statistics":{"estimate":0.1481,"SE":0.3453,"df":339.5108,"statistic":{"name":"t","value":0.429},"p":0.6682}},{"name":"Product4","statistics":{"estimate":0.2963,"SE":0.3453,"df":339.5108,"statistic":{"name":"t","value":0.858},"p":0.3915}},{"name":"Information2:Product2","statistics":{"estimate":0.2469,"SE":0.287,"df":320.0004,"statistic":{"name":"t","value":0.8605},"p":0.3902}},{"name":"Information2:Product3","statistics":{"estimate":0.2716,"SE":0.287,"df":320.0004,"statistic":{"name":"t","value":0.9465},"p":0.3446}},{"name":"Information2:Product4","statistics":{"estimate":-0.358,"SE":0.287,"df":320.0004,"statistic":{"name":"t","value":-1.2477},"p":0.2131}}],"correlations":[{"names":["(Intercept)","Gender2"],"statistics":{"r":-0.4526}},{"names":["(Intercept)","Information2"],"statistics":{"r":-0.3569}},{"names":["(Intercept)","Product2"],"statistics":{"r":-0.6074}},{"names":["(Intercept)","Product3"],"statistics":{"r":-0.6074}},{"names":["(Intercept)","Product4"],"statistics":{"r":-0.6074}},{"names":["(Intercept)","Information2:Product2"],"statistics":{"r":0.2523}},{"names":["(Intercept)","Information2:Product3"],"statistics":{"r":0.2523}},{"names":["(Intercept)","Information2:Product4"],"statistics":{"r":0.2523}},{"names":["Gender2","Information2"],"statistics":{"r":-1.1265E-16}},{"names":["Gender2","Product2"],"statistics":{"r":1.3608E-14}},{"names":["Gender2","Product3"],"statistics":{"r":1.3608E-14}},{"names":["Gender2","Product4"],"statistics":{"r":1.3608E-14}},{"names":["Gender2","Information2:Product2"],"statistics":{"r":7.9653E-17}},{"names":["Gender2","Information2:Product3"],"statistics":{"r":7.9653E-17}},{"names":["Gender2","Information2:Product4"],"statistics":{"r":7.9653E-17}},{"names":["Information2","Product2"],"statistics":{"r":0.2938}},{"names":["Information2","Product3"],"statistics":{"r":0.2938}},{"names":["Information2","Product4"],"statistics":{"r":0.2938}},{"names":["Information2","Information2:Product2"],"statistics":{"r":-0.7071}},{"names":["Information2","Information2:Product3"],"statistics":{"r":-0.7071}},{"names":["Information2","Information2:Product4"],"statistics":{"r":-0.7071}},{"names":["Product2","Product3"],"statistics":{"r":0.5}},{"names":["Product2","Product4"],"statistics":{"r":0.5}},{"names":["Product2","Information2:Product2"],"statistics":{"r":-0.4155}},{"names":["Product2","Information2:Product3"],"statistics":{"r":-0.2077}},{"names":["Product2","Information2:Product4"],"statistics":{"r":-0.2077}},{"names":["Product3","Product4"],"statistics":{"r":0.5}},{"names":["Product3","Information2:Product2"],"statistics":{"r":-0.2077}},{"names":["Product3","Information2:Product3"],"statistics":{"r":-0.4155}},{"names":["Product3","Information2:Product4"],"statistics":{"r":-0.2077}},{"names":["Product4","Information2:Product2"],"statistics":{"r":-0.2077}},{"names":["Product4","Information2:Product3"],"statistics":{"r":-0.2077}},{"names":["Product4","Information2:Product4"],"statistics":{"r":-0.4155}},{"names":["Information2:Product2","Information2:Product3"],"statistics":{"r":0.5}},{"names":["Information2:Product2","Information2:Product4"],"statistics":{"r":0.5}},{"names":["Information2:Product3","Information2:Product4"],"statistics":{"r":0.5}}]},"package":{"name":"lme4","version":"1.1-21"}}}';

  data = JSON.parse(text);

  createAnalysesList(data);  
}

function insertStatisticsTable() {


  var id = this.id;
  
  // Split the id by $, which should give us the necessary information to figure
  // out which statistic or statistics to retrieve
  id_components = id.split("$");

  // Get the identifier, which should be the second component
  identifier = id_components[1];

  // Retrieve the analysis and its method
  var analysis = data[identifier]
  var method = analysis.method;

  return Word.run(function (context) {

    var body = context.document.body;
    var table;

    // Determine table statistics
    if (/t-test/.test(method)) {
      table = createTTestWordTable(body, identifier, analysis);
    } else if (/Linear regression/.test(method)) {
      table = createRegressionWordTable(body, identifier, analysis);
    }

    return context.sync();
  }).catch(function (e) {
      console.log(e.message);
  })
}