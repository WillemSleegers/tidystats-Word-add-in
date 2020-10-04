
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

function insert(attrs) {
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
    textArea.style.position = 'fixed';
    textArea.style.top = 0;
    textArea.style.left = 0;
  
    // Ensure it has a small width and height. Setting to 1px / 1em
    // doesn't work as this gives a negative w/h on some browsers.
    textArea.style.width = '2em';
    textArea.style.height = '2em';
  
    // We don't need padding, reducing the size if it does flash render.
    textArea.style.padding = 0;
  
    // Clean up any borders.
    textArea.style.border = 'none';
    textArea.style.outline = 'none';
    textArea.style.boxShadow = 'none';
  
    // Avoid flash of white box if rendered for any reason.
    textArea.style.background = 'transparent';
    textArea.value = text;
  
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
  
    try {
      var successful = document.execCommand('copy');
      var msg = successful ? 'successful' : 'unsuccessful';
      console.log('Copying text command was ' + msg);
      document.getElementById("cite3").firstChild.innerHTML = "BibTeX copied!";
      setTimeout(function() {
        document.getElementById("cite3").firstChild.innerHTML = "Copy BibTeX";
      }, 3000);
    } catch (err) {
      console.log('Oops, unable to copy');
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
      var model = analysis.models.find(x => x.name == attrs.model);
      statistics = model.statistics;
    } else if ("effect" in attrs) {
      var effect = analysis.effects[attrs.effect];
      
      if ("group" in attrs) {
        var group = effect.groups.find(x => x.name == attrs.group);
        
        if ("term" in attrs) {
          var term = group.terms.find(x => x.name == attrs.term);
          statistics = term.statistics;  
        } else if ("pair1" in attrs) {
          var pair = group.pairs.find(x => x.names[0] == attrs.pair1 & x.names[1] == attrs.pair2);
          statistics = pair.statistics;
        } else {
          statistics = group.statistics; 
        }
      } else if ("term" in attrs) {
        var term = effect.terms.find(x => x.name == attrs.term);
        statistics = term.statistics;
      } else if ("pair1" in attrs) {
        var pair = effect.pairs.find(x => x.names[0] == attrs.pair1 & x.names[1] == attrs.pair2);
        statistics = pair.statistics;
      } else if ("statistic" in attrs) {
        statistics = effect.statistics;
      }
    } else if ("group" in attrs) {
        var group = analysis.groups.find(x => x.name == attrs.group);
        
        if ("term" in attrs) {
          var term = group.terms.find(x => x.name == attrs.term);
          statistics = term.statistics;  
        } else {
          statistics = group.statistics;  
        }
    } else if ("term" in attrs) {
      var term = analysis.terms.find(x => x.name == attrs.term);
      statistics = term.statistics;  
    } else {
      statistics = analysis.statistics;
    }

    // Determine whether a single statistic or multiple statistics should be inserted
    if ("statistic" in attrs) {
      single = true
    } else {
      single = false
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
            text = "<i>" + formatName(name) + "</i>(" + formatNumber(df, "df") + ") = " + 
              formatNumber(value, name);
          } else if (selectedStatistics.includes("df_numerator")) {
            var df_numerator = statistics.dfs.df_numerator;
            var df_denominator = statistics.dfs.df_denominator;
            text = "<i>" + formatName(name) + "</i>(" + 
              formatNumber(df_numerator, "df") + ", " + 
              formatNumber(df_denominator, "df") + ") = " + 
              formatNumber(value, name);
          } else {
            text = "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
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
            text = "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
            output.push(text);
            break;
          }
        case "df_numerator":
          if (selectedStatistics.includes("statistic")) {
            break;
          } else {
            value = statistics[name];
            text = "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
            output.push(text);
            break;
          }
        case "df_denominator":
          if (selectedStatistics.includes("statistic")) {
            break;
          } else {
            value = statistics[name];
            text = "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
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
          	text = "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
          }
          output.push(text);
          break;
        case "CI_lower":
          text = createCILine(statistics.CI.CI_lower, statistics.CI.CI_upper, 
            statistics.CI.CI_level);
          output.push(text);
          break;
        case "CI_upper":
          break;
        case "BF_01":
          value = statistics[name];
        
          if (selectedStatistics.includes("error")) {
            var error = statistics["error"];
            text = formatName(name) + " = " + formatNumber(value, name) + " ±" + 
              formatNumber(error) + "%";
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
            text = formatName(name) + " = " + formatNumber(value, name) + " ±" + 
              formatNumber(error) + "%";
          } else {
            value = statistics[name];
            text = formatName(name) + " = " + formatNumber(value, name);
          }
          output.push(text);
          break;
        case "error":
          if (selectedStatistics.includes("BF_01") || selectedStatistics.includes("BF_10")) {
            break;
          } else {
            value = statistics[name];
            text = "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
            output.push(text);
          }
          break;
        case "n":
          if (selectedStatistics.includes("pct")) {
            var pct = statistics.pct;
            value = statistics[name];
            text = "<i>" + formatName(name) + "</i> = " + formatNumber(value, name) + 
              " (" + formatNumber(pct, "pct") + "%)";
          } else {
            value = statistics[name];
            text = "<i>" + formatName(name) + "</i> = " + formatNumber(value, name);
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
  
  return(output)
}

function createCILine(lower, upper, level) {
  return(
    level * 100 + "% CI [" + 
      formatNumber(lower) + ", " + 
      formatNumber(upper) + "]"
  )
}

// Work in progress ----------------------------------------------------

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