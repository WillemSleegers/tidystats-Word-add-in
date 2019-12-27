
// Debugging
var instantly_load_data = false;

function instantlyLoadData() {
  console.log("Instantly loading data");
  
  var text = '{"m":{"method":"Linear mixed model","REML_criterion_at_convergence":1743.6283,"convergence_code":0,"random_effects":{"N":180,"groups":[{"name":"Subject","N":18,"variances":[{"name":"(Intercept)","statistics":{"var":611.8976,"SD":24.7366}},{"name":"Days","statistics":{"var":35.0811,"SD":5.9229}}],"correlations":[{"names":["(Intercept)","Days"],"statistics":{"r":0.0656}}]},{"name":"Residual","variances":[{"statistics":{"var":654.9408,"SD":25.5918}}]}]},"fixed_effects":{"coefficients":[{"name":"(Intercept)","statistics":{"estimate":251.4051,"SE":6.8238,"df":17.0052,"statistic":{"name":"t","value":36.8425},"p":1.1582E-17}},{"name":"Days","statistics":{"estimate":10.4673,"SE":1.546,"df":16.9953,"statistic":{"name":"t","value":6.7707},"p":3.273E-6}}],"correlations":[{"names":["(Intercept)","Days"],"statistics":{"r":-0.1375}}]},"package":{"name":"lme4","version":"1.1-21"}},"fm":{"method":"Linear mixed model","REML_criterion_at_convergence":2705.5037,"convergence_code":0,"random_effects":{"N":648,"groups":[{"name":"Consumer:Product","N":324,"variances":[{"name":"(Intercept)","statistics":{"var":3.1622,"SD":1.7783}}]},{"name":"Consumer","N":81,"variances":[{"name":"(Intercept)","statistics":{"var":0.3756,"SD":0.6129}}]},{"name":"Residual","variances":[{"statistics":{"var":1.6675,"SD":1.2913}}]}]},"fixed_effects":{"coefficients":[{"name":"(Intercept)","statistics":{"estimate":5.849,"SE":0.2843,"df":322.3361,"statistic":{"name":"t","value":20.5742},"p":1.1733E-60}},{"name":"Gender2","statistics":{"estimate":-0.2443,"SE":0.2606,"df":79,"statistic":{"name":"t","value":-0.9375},"p":0.3514}},{"name":"Information2","statistics":{"estimate":0.1605,"SE":0.2029,"df":320.0004,"statistic":{"name":"t","value":0.791},"p":0.4296}},{"name":"Product2","statistics":{"estimate":-0.8272,"SE":0.3453,"df":339.5108,"statistic":{"name":"t","value":-2.3953},"p":0.0171}},{"name":"Product3","statistics":{"estimate":0.1481,"SE":0.3453,"df":339.5108,"statistic":{"name":"t","value":0.429},"p":0.6682}},{"name":"Product4","statistics":{"estimate":0.2963,"SE":0.3453,"df":339.5108,"statistic":{"name":"t","value":0.858},"p":0.3915}},{"name":"Information2:Product2","statistics":{"estimate":0.2469,"SE":0.287,"df":320.0004,"statistic":{"name":"t","value":0.8605},"p":0.3902}},{"name":"Information2:Product3","statistics":{"estimate":0.2716,"SE":0.287,"df":320.0004,"statistic":{"name":"t","value":0.9465},"p":0.3446}},{"name":"Information2:Product4","statistics":{"estimate":-0.358,"SE":0.287,"df":320.0004,"statistic":{"name":"t","value":-1.2477},"p":0.2131}}],"correlations":[{"names":["(Intercept)","Gender2"],"statistics":{"r":-0.4526}},{"names":["(Intercept)","Information2"],"statistics":{"r":-0.3569}},{"names":["(Intercept)","Product2"],"statistics":{"r":-0.6074}},{"names":["(Intercept)","Product3"],"statistics":{"r":-0.6074}},{"names":["(Intercept)","Product4"],"statistics":{"r":-0.6074}},{"names":["(Intercept)","Information2:Product2"],"statistics":{"r":0.2523}},{"names":["(Intercept)","Information2:Product3"],"statistics":{"r":0.2523}},{"names":["(Intercept)","Information2:Product4"],"statistics":{"r":0.2523}},{"names":["Gender2","Information2"],"statistics":{"r":-1.1265E-16}},{"names":["Gender2","Product2"],"statistics":{"r":1.3608E-14}},{"names":["Gender2","Product3"],"statistics":{"r":1.3608E-14}},{"names":["Gender2","Product4"],"statistics":{"r":1.3608E-14}},{"names":["Gender2","Information2:Product2"],"statistics":{"r":7.9653E-17}},{"names":["Gender2","Information2:Product3"],"statistics":{"r":7.9653E-17}},{"names":["Gender2","Information2:Product4"],"statistics":{"r":7.9653E-17}},{"names":["Information2","Product2"],"statistics":{"r":0.2938}},{"names":["Information2","Product3"],"statistics":{"r":0.2938}},{"names":["Information2","Product4"],"statistics":{"r":0.2938}},{"names":["Information2","Information2:Product2"],"statistics":{"r":-0.7071}},{"names":["Information2","Information2:Product3"],"statistics":{"r":-0.7071}},{"names":["Information2","Information2:Product4"],"statistics":{"r":-0.7071}},{"names":["Product2","Product3"],"statistics":{"r":0.5}},{"names":["Product2","Product4"],"statistics":{"r":0.5}},{"names":["Product2","Information2:Product2"],"statistics":{"r":-0.4155}},{"names":["Product2","Information2:Product3"],"statistics":{"r":-0.2077}},{"names":["Product2","Information2:Product4"],"statistics":{"r":-0.2077}},{"names":["Product3","Product4"],"statistics":{"r":0.5}},{"names":["Product3","Information2:Product2"],"statistics":{"r":-0.2077}},{"names":["Product3","Information2:Product3"],"statistics":{"r":-0.4155}},{"names":["Product3","Information2:Product4"],"statistics":{"r":-0.2077}},{"names":["Product4","Information2:Product2"],"statistics":{"r":-0.2077}},{"names":["Product4","Information2:Product3"],"statistics":{"r":-0.2077}},{"names":["Product4","Information2:Product4"],"statistics":{"r":-0.4155}},{"names":["Information2:Product2","Information2:Product3"],"statistics":{"r":0.5}},{"names":["Information2:Product2","Information2:Product4"],"statistics":{"r":0.5}},{"names":["Information2:Product3","Information2:Product4"],"statistics":{"r":0.5}}]},"package":{"name":"lme4","version":"1.1-21"}}}';

  data = JSON.parse(text);

  createAnalysesList(data);  
}

// Global variables ------------------------------------------------------------

var data;

// Word functions --------------------------------------------------------------

Office.onReady(function(info) {
  // Check if a Word application is running; if not, show that tidystats failed 
  // to load
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    document.getElementById("file").onchange = readFile;
    document.getElementById("file").onclick = resetFile;
    
    document.getElementById("update_button").onclick = updateStatistics;
    document.getElementById("search").onkeyup = collapse;

    document.getElementById("cite_1").onclick = insertInTextCitation;
    document.getElementById("cite_2").onclick = insertFullCitation;

    // Check settings to see whether a file was already opened
    console.log("Loaded add-in");
    var file = Office.context.document.settings.get('file_name');
    var saved_data = Office.context.document.settings.get('data');

    if (instantly_load_data) {
      instantlyLoadData();
    }

  } else {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("fail-msg").style.display = "block";
  }
});

function insertStatistics() {
  // Save the identifier of the element that this function was attached to;
  // this defines the specific output
  var id = this.id;
  
  Word.run(function (context) {
    // Determine output
    var output = retrieveStatistics(data, id);
    
    // Create a context control
    var doc = context.document;
    var selection = doc.getSelection();
    var selection_font = selection.font;
    selection_font.load("name");
    selection_font.load("size");

    var content_control = selection.insertContentControl();

    // Set analysis specific information
    content_control.tag = id;
    content_control.title = id;
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
          var id = tag;
          var output = retrieveStatistics(data, id);

          // Set the new output
          content_control.insertHtml(output, "Replace");
        }
      });
  });
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

// Read .json file function ----------------------------------------------------

function readFile() {
  // Read .json file
  var fileInput = document.getElementById('file');
  var files = fileInput.files;
  var file = fileInput.files[0];
  
  // Update label
  var label = document.getElementById("file_label");
  label.innerHTML = 'File: ' + file.name;

  var reader = new FileReader();
  
  reader.onload = function(e) {
    var text = reader.result;
    data = JSON.parse(text);
        
    saveData(file.name, data);
    createAnalysesList(data);
  }

  reader.readAsText(file); 
}

function resetFile() {
  this.value = null;
}

function saveData(file_name, statistics) {
  Office.context.document.settings.set('file_name', file_name);
  Office.context.document.settings.set('data', statistics);

  Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        console.log('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        console.log('Settings saved.');
    }
  });
}

// Search related functions ----------------------------------------------------

function collapse() {
	var input, filter, analyses, div, a, button, i, txtValue;
	input = document.getElementById("search");
	filter = input.value.toUpperCase();
	analyses = document.getElementById("analyses");
	div = analyses.getElementsByClassName("analysis");
	for (i = 0; i < div.length; i++) {
	    var analysis = div[i];
	    
	    button = analysis.getElementsByClassName("collapsible")[0];
	    
	    txtValue = button.innerText;
	    
	    if (txtValue.toUpperCase().indexOf(filter) > -1) {
	        div[i].style.display = "";
	    } else {
	        div[i].style.display = "none";
	    }
	}
}

// Statistics retrieval/formatting functions -----------------------------------

function formatNumber(x, type) {

  // Set type to 'standard' if no type is provided
  type = type || 'standard';

  var number;

  if (x == "Inf") {
    number = "&infin;";
  } else if (x == "-Inf") {
    number = "-&infin;"; // Would be nice to solve this issue with .includes and .replace at some point
  } else {
    number = parseFloat(x);

    if (type == "df" | type == "N" | type == "missing") {
      if (number % 1 == 0) {
          number = number.toFixed(0);
      } else {
        number = number.toFixed(2);
      }
    } else if (type == "p") {
      if (number < 0.001) {
          number = "< .001"
      } else if (number < .01) {
          number = number.toFixed(4);
          number = number.toString();
          number = number.substr(1);
      } else if (number < .1) {
          number = number.toFixed(3);
          number = number.toString();
          number = number.substr(1);
      } else {
          number = number.toFixed(2);
          number = number.toString();
          number = number.substr(1);
      }
    } else if (type == "r") {
      // if correlation, omit the leading 0
      number = number.toFixed(2);
      if (number < 0) {
        number = number.toString();
        number = number.slice(0, 1) + number.slice(2);  
      } else {
        number = number.toString();
        number = number.substr(1);
      }
    } else {
      number = number.toFixed(2);
    }
  }                

  return(number)
}

function formatCIs(CIs) {
  var text = "[" + formatNumber(CIs.lower) + ", " + formatNumber(CIs.upper) + "]";
  return(text)
}

function retrieveStatistic(statistics, statistic) {

	var output;

	if (statistic == "b") {
		output = formatNumber(statistics.estimate);
	} else if (statistic == "r") {

    if ("estimate" in statistics) {
      output = formatNumber(statistics.estimate, statistic);  
    } else {
      output = formatNumber(statistics.r, statistic);  
    }

	} else if (statistic == "tau") {
		output = formatNumber(statistics.estimate);
	} else if (statistic == "rho") {
		output = formatNumber(statistics.estimate);
	} else if (statistic == "OR") {
		output = formatNumber(statistics.estimate);
	} else if (statistic == "t") {
		output = formatNumber(statistics.statistic.value);
	} else if (statistic == "F") {
		output = formatNumber(statistics.statistic.value);
	} else if (statistic == "z") {
		output = formatNumber(statistics.statistic.value);
	} else if (statistic == "S") {
		output = formatNumber(statistics.statistic.value);
	} else if (statistic == "X-squared") {
		output = formatNumber(statistics.statistic.value);
	} else if (statistic == "V") {
		output = formatNumber(statistics.statistic.value);
	} else if (statistic == "W") {
		output = formatNumber(statistics.statistic.value);
	} else if (statistic == "df") {
		output = formatNumber(statistics.df, "df");
	} else if (statistic == "df1") {
		output = formatNumber(statistics.dfs.numerator_df, "df");
	} else if (statistic == "df2") {
		output = formatNumber(statistics.dfs.denominator_df, "df");
	} else if (statistic == "p") {
		output = formatNumber(statistics.p, "p");
  } else if (statistic == "CI") {
    output = formatCIs(statistics.CI);
	} else if (statistic == "CI_lower") {
		output = formatNumber(statistics.CI.lower);
	} else if (statistic == "CI_upper") {
		output = formatNumber(statistics.CI.upper);
	} else if (statistic == "R-squared") {
		output = formatNumber(statistics.r_squared);
	} else if (statistic == "adj-R-squared") {
		output = formatNumber(statistics.adjusted_r_squared);
	} else {
		output = formatNumber(statistics[statistic], statistic);
	}

	return(output)
}

function retrieveStatistics(data, id) {
    var id_components, identifier, analysis, method, statistic;

    // Split the id by $, which should give us the necessary information to figure
    // out which statistic or statistics to retrieve
    id_components = id.split("$");

    // Get the identifier, which should be the first component
    identifier = id_components[0];

    // Extract the specific analysis and what kind of method it is
    analysis = data[identifier];
    method = analysis.method;

    // Create a variable to store the output in
    var output;

    if (/t-test/.test(method)) {
      var statistics = analysis["statistics"];

      // Determine whether a single statistic or line of statistics should be
      // created
      if (id_components.length == 2) {
        // Extract the statistic, which should be the last component
        statistic = id_components[id_components.length - 1];
        output = retrieveStatistic(statistics, statistic);
      } else {
        output = createTTestLine(statistics);
      }
      
    } else if (/Pearson's product-moment correlation/.test(method)) {
      var statistics = analysis["statistics"];

      // Determine whether a single statistic or line of statistics should be
      // created
      if (id_components.length == 2) {
        // Extract the statistic, which should be the last component
        statistic = id_components[id_components.length - 1];
        output = retrieveStatistic(statistics, statistic);
      } else {
        output = createPearsonCorrelationLine(statistics);
      }

    } else if (/Kendall's rank correlation tau/.test(method)) {
      var statistics = analysis["statistics"];

      // Determine whether a single statistic or line of statistics should be
      // created
      if (id_components.length == 2) {
        // Extract the statistic, which should be the last component
        statistic = id_components[id_components.length - 1];
        output = retrieveStatistic(statistics, statistic);
      } else {
        output = createKendallCorrelationLine(statistics);
      }

    } else if (/Spearman's rank correlation rho/.test(method)) {
      var statistics = analysis["statistics"];

      // Determine whether a single statistic or line of statistics should be
      // created
      if (id_components.length == 2) {
        // Extract the statistic, which should be the last component
        statistic = id_components[id_components.length - 1];
        output = retrieveStatistic(statistics, statistic);
      } else {
        output = createSpearmanCorrelationLine(statistics);
      }
      
    } else if (/Chi-squared test/.test(method)) {
      var statistics = analysis["statistics"];

      // Determine whether a single statistic or line of statistics should be
      // created
      if (id_components.length == 2) {
        // Extract the statistic, which should be the last component
        statistic = id_components[id_components.length - 1];
        output = retrieveStatistic(statistics, statistic);
      } else {
        output = createChiSquaredLine(statistics);
      }

    } else if (/Wilcoxon/.test(method)) {
      var statistics = analysis["statistics"];

      // Determine whether a single statistic or line of statistics should be
      // created
      if (id_components.length == 2) {
        // Extract the statistic, which should be the last component
        statistic = id_components[id_components.length - 1];
        output = retrieveStatistic(statistics, statistic);
      } else {
        output = createWilcoxonTestLine(statistics);
      }

    } else if (/Fisher's Exact Test/.test(method)) {
      var statistics = analysis["statistics"];

      // Determine whether a single statistic or line of statistics should be
      // created
      if (id_components.length == 2) {
        // Extract the statistic, which should be the last component
        statistic = id_components[id_components.length - 1];
        output = retrieveStatistic(statistics, statistic);
      } else {
        output = createFisherExactTestLine(statistics);
      }
      
    } else if (/One-way analysis of means/.test(method)) {
      var statistics = analysis["statistics"];

      // Determine whether a single statistic or line of statistics should be
      // created
      if (id_components.length == 2) {
        // Extract the statistic, which should be the last component
        statistic = id_components[id_components.length - 1];
        output = retrieveStatistic(statistics, statistic);
      } else {
        output = createOneWayAnalysisOfMeansLine(statistics);
      }

    } else if (/ANOVA/.test(method)) {

      // Check if it is an ANOVA with a within-subjects factor
      if ("groups" in analysis) {
        
      	// Retrieve the group
      	var group_name = id_components[1];
		var i = 0;
		while (group_name != analysis.groups[i].name) {
			i++;
		}
		var group = analysis.groups[i];

        // Retrieve the statistics
      	var coefficient_name = id_components[2];
		var i = 0;
		while (coefficient_name != group.coefficients[i].name) {
			i++;
		}
		var coefficient = group.coefficients[i];
		var statistics = coefficient.statistics;

        // Determine whether a single statistic or line of statistics should be
      	// created
      	if (id_components.length == 4) {
        	// Extract the statistic, which should be the last component
        	statistic = id_components[id_components.length - 1];
        	output = retrieveStatistic(statistics, statistic);
      	} else {
      		var statistics_residuals = group.coefficients[group.coefficients.length - 1].statistics;
        	output = createANOVALine(statistics, statistics_residuals);
      	}

      } else {
        var coefficient_name = id_components[1];

        // Retrieve the statistics
		var i = 0;
		while (coefficient_name != analysis.coefficients[i].name) {
			i++;
		}
		var statistics = analysis.coefficients[i].statistics;
		
		// Determine whether a single statistic or line of statistics should be
      	// created
      	if (id_components.length == 3) {
        	// Extract the statistic, which should be the last component
        	statistic = id_components[id_components.length - 1];
        	output = retrieveStatistic(statistics, statistic);
      	} else {
      		var statistics_residuals = 
          analysis.coefficients[analysis.coefficients.length-1].statistics;
        	output = createANOVALine(statistics, statistics_residuals);
      	}
      }
    } else if (/Linear regression/.test(method)) {
      // Check if a coefficient or model line should be produced
      if (id_components[1] == "model") {
        var statistics = analysis.model.statistics;

        // Determine whether a single statistic or line of statistics should be
      	// created
      	if (id_components.length == 3) {
        	// Extract the statistic, which should be the last component
        	statistic = id_components[id_components.length - 1];
        	output = retrieveStatistic(statistics, statistic);
      	} else {
        	output = createLinearModelModelFitLine(statistics);
      	}
        
      } else {
        var coefficient_name = id_components[1];

        // Retrieve the statistics
    		var i = 0;
    		while (coefficient_name != analysis.coefficients[i].name) {
    			i++;
    		}
    		var statistics = analysis.coefficients[i].statistics;

        // Determine whether a single statistic or line of statistics should be
      	// created
      	if (id_components.length == 3) {
        	// Extract the statistic, which should be the last component
        	statistic = id_components[id_components.length - 1];
        	output = retrieveStatistic(statistics, statistic);
      	} else {
        	output = createLinearModelCoefficientLine(statistics);
      	}
      }
    } else if (/Linear mixed model/.test(method)) {

      console.log(id_components);

      // Check whether the user click on a fixed effect name
      if (id_components[1] == "FE" & id_components.length == 4) {
        
        // Produce a line of output
        var statistics = analysis.fixed_effects.coefficients[id_components[2]].statistics;
        
        output = createLinearMixedModelFixedEffectLine(statistics);
      } else {
        // Produce a single statistic output
        statistic = id_components[id_components.length - 1];

        if (id_components[1] == "RE") {
          var group = analysis.random_effects.groups[id_components[2]];

          if (statistic == "r") {
            var statistics = group.correlations[id_components[4]].statistics;
          } else {
            var statistics = group.variances[id_components[4]].statistics;
          }
        } else {
          if (statistic == "r") {
            var statistics = analysis.fixed_effects.correlations[id_components[2]].statistics;
          } else {
            var statistics = analysis.fixed_effects.coefficients[id_components[2]].statistics;
          }
        }

        output = retrieveStatistic(statistics, statistic);
      }
      
    } else if (/Descriptives/.test(method)) {
      
      // Check whether the descriptives are from a group or not
      if ("groups" in analysis) {
      	// Extract the group
      	var group_name = id_components[1];
      	var i = 0;
		while (group_name != analysis.groups[i].name) {
			i++;
		}
		var statistics = analysis.groups[i].statistics;

      	// Check whether the user clicked on the name of the group or a specific statistic
      	if (id_components.length == 3) {
      		// Extract the statistic, which should be the last component
	        statistic = id_components[id_components.length - 1];
	        output = retrieveStatistic(statistics, statistic);
      	} else {
      		output = createDescriptivesLine(statistics);
      	}
      } else {
      	var statistics = analysis["statistics"];	
      	  // Determine whether a single statistic or line of statistics should be
	      // created
	      if (id_components.length == 2) {
	        // Extract the statistic, which should be the last component
	        statistic = id_components[id_components.length - 1];
	        output = retrieveStatistic(statistics, statistic);
	      } else {
	        output = createDescriptivesLine(statistics);
	      }
      }
    } else {
      output = "Sorry, not supported";
    }

    return(output)
}

// Table creation functions ----------------------------------------------------

function createAnalysesList(data) {
  
  var div_analyses = document.getElementById("analyses");

  // Reset analyses in case a file was already loaded in
  while (div_analyses.firstChild) {
      div_analyses.removeChild(div_analyses.firstChild);
  }

  for (var identifier in data) {
      var div_analysis = document.createElement("div");
      div_analysis.className = "analysis";

      var button = document.createElement("button");
      button.className = "collapsible";
      button.innerText = identifier;

      var content = document.createElement("div");
      content.className = "content";

      var analysis = data[identifier]
      
      // Add notes
      if ("notes" in analysis) {
          var p = document.createElement("p");
          p.className = "description";
          p.innerHTML = "<strong>Description: </strong>" + analysis["notes"];
          content.appendChild(p);
      }

      // Add method
      var method = analysis["method"];
      var p = document.createElement("p");
      p.className = "method";
      p.innerHTML = "<strong>Method: </strong>" + method;
      content.appendChild(p);

      // Add APA table(s) of statistics
      var table;

      var analysis = data[identifier];
      var method = analysis["method"];

      // Determine what kind of table to make
      if (/t-test/.test(method)) {

          // Create a t-test table
          table = createTTestTable(identifier, analysis);
          content.appendChild(table);    
          
      } else if (/Pearson's product-moment correlation/.test(method)) {
          // Create a Pearson correlation table
          table = createPearsonCorrelationTable(identifier, analysis);
          content.appendChild(table);  
      } else if (/Kendall's rank correlation tau/.test(method)) {
          // Create a Kendall correlation table
          table = createKendallCorrelationTable(identifier, analysis);
          content.appendChild(table);  
      } else if (/Spearman's rank correlation rho/.test(method)) {
          // Create a Spearman correlation table
          table = createSpearmanCorrelationTable(identifier, analysis);
          content.appendChild(table);  
      } else if (/Chi-squared test/.test(method)) {
          // Create a Chi-squared table
          table = createChiSquaredTable(identifier, analysis);
          content.appendChild(table);  
      } else if (/Wilcoxon/.test(method)) {
          // Create a Wilcoxon table
          table = createWilcoxonTable(identifier, analysis);
          content.appendChild(table);  
      } else if (/Fisher's Exact Test/.test(method)) {
          // Create a Fisher's Exact test table
          table = createFisherExactTestTable(identifier, analysis);
          content.appendChild(table);  
      } else if (/One-way analysis of means/.test(method)) {
          // Create a one-way analysis of means table
          table = createOneWayAnalysisOfMeansTable(identifier, analysis);
          content.appendChild(table);  
      } else if (/ANOVA/.test(method)) {
          // Create an ANOVA table
          if ("groups" in analysis) {
            var tables = createANOVATables(identifier, analysis);
            
            for (table in tables) {
              content.appendChild(tables[table]);    
            }

          } else {
            table = createANOVATable(identifier, analysis);
            content.appendChild(table);
          }
          
      } else if (/Linear regression/.test(method)) {
          // Create linear regression tables (coefficient and model table)
          tables = createLinearModelTables(identifier, analysis);
          
          content.appendChild(tables[0]);

          // Add insert table button
          var insert_table_button = document.createElement("button");
          insert_table_button.onclick = insertStatisticsTable;
          insert_table_button.innerHTML = "Insert table";
          insert_table_button.className = "insert_table_button";
          insert_table_button.id = "table$" + identifier;
          content.appendChild(insert_table_button);

          content.appendChild(tables[1]);
      } else if (/Linear mixed model/.test(method)) {
          // Create linear regression tables (coefficient and model table)
          tables = createLinearMixedModelTables(identifier, analysis);
          
          for (var i in tables) {
            content.appendChild(tables[i]);  
          }
      } else if (/Descriptives/.test(method)) {
          
          // Create multiple tables; one for each variable
      	  table = createDescriptivesTable(identifier, analysis);
          content.appendChild(table);  

      } else {
          console.log("not supported")
      }



      // Add collapse button and content
      div_analysis.appendChild(button);
      div_analysis.appendChild(content);

      div_analyses.appendChild(div_analysis);
  }

  // Add collapse functions to the buttons
  var coll = document.getElementsByClassName("collapsible");
  var i;

  for (i = 0; i < coll.length; i++) {
    coll[i].addEventListener("click", function() {
      this.classList.toggle("active");
      var content = this.nextElementSibling;
      if (content.style.maxHeight){
        content.style.maxHeight = null;
      } else {
        content.style.maxHeight = content.scrollHeight + "px";
      } 
    });
  }

  // Make the main div visible
  var main = document.getElementById("main");
  main.style.display = "block";
}

function createStatisticsTable() {
  var table = document.createElement("table");
  table.className = "statistics_table";

  return(table)
}

function createGroupRow(name) {
  // Create elements
  var row = document.createElement("tr");
  var cell = document.createElement("th");
  var span_group = document.createElement("span");
  var span_group_name = document.createElement("span");

  // Set classes for CSS purposes
  row.className = "row_group";

  // Set column span to 2 since it needs to cover the statistics name and value cells
  cell.colSpan = "2";

  // Set spans
  span_group.innerHTML = "Group: ";
  span_group_name.innerHTML = name;

  // Set cell content
  cell.appendChild(span_group);
  cell.appendChild(span_group_name);

  // Add cell to row
  row.appendChild(cell);

  return(row)
}

function createNameRow(id, name) {
  // Create elements
  var row_name = document.createElement("tr");
  var cell_name = document.createElement("th");
  var insert_link = document.createElement("a");

  // Set classes for CSS purposes
  row_name.className = "statistics_row_name";
  insert_link.className = "insert_link";

  // Set column span to 2 since it needs to cover the statistics name and value cells
  cell_name.colSpan = "2";

  // Set the id we use to figure out what was clicked on
  insert_link.id = id;

  // Set the displayed text
  insert_link.innerHTML = name;

  // Disable the link
  insert_link.href = "javascript:void(0);";

  // Set the function to insert statistics
  insert_link.onclick = insertStatistics;

  // Append children
  cell_name.appendChild(insert_link);
  row_name.appendChild(cell_name);

  return(row_name)
}

function createNameRowWithoutLink(name) {
  // Create elements
  var row_name = document.createElement("tr");
  var cell_name = document.createElement("th");
  
  // Set classes for CSS purposes
  row_name.className = "statistics_row_name";

  // Set column span to 2 since it needs to cover the statistics name and value cells
  cell_name.colSpan = "2";

  // Set cell content
  cell_name.innerHTML = name;

  // Add cell to row
  row_name.appendChild(cell_name);

  return(row_name)
}

function createStatisticsRow(id, statistic_name, statistic_value, extra) {
  // Create elements
  var row_statistic = document.createElement("tr");
  var cell_statistic_name = document.createElement("td");
  var cell_statistic_value = document.createElement("td");
  var insert_link_statistic = document.createElement("a");

  // Set name and value classes for CSS purposes
  cell_statistic_name.className = "statistics_cell_name";
  cell_statistic_value.innerHTML = "statistics_cell_value";

  // Create the insert link
  insert_link_statistic.className = "insert_link";
  insert_link_statistic.href = "javascript:void(0);";
  insert_link_statistic.onclick = insertStatistics;
  insert_link_statistic.id = id + "$" + statistic_name;
  
  // Determine the statistic name
  // Handle exceptions such as CIs, tau's, etc.
  if (statistic_name == "CI_lower") {
    insert_link_statistic.innerHTML = extra * 100 + "% CI lower";
  } else if (statistic_name == "CI_upper") {
    insert_link_statistic.innerHTML = extra * 100 + "% CI upper";
  } else if (statistic_name == "df1") {
    insert_link_statistic.innerHTML = "numerator df";
  } else if (statistic_name == "df2") {
    insert_link_statistic.innerHTML = "denominator df";
  } else if (statistic_name == "tau") {
  	insert_link_statistic.innerHTML = "r<sub>&tau;<sub>";
  } else if (statistic_name == "rho") {
  	insert_link_statistic.innerHTML = "r<sub>S</sub>";
  } else if (statistic_name == "X-squared") {
  	insert_link_statistic.innerHTML = "&chi;²";
  } else if (statistic_name == "R-squared") {
  	insert_link_statistic.innerHTML = "R²";
  } else if (statistic_name == "adj-R-squared") {
  	insert_link_statistic.innerHTML = "adj. R²";
  } else {
    insert_link_statistic.innerHTML = statistic_name;
  }

  // Set the statistics value
  if (statistic_name == "df1" | statistic_name == "df2") {
    cell_statistic_value.innerHTML = formatNumber(statistic_value, "df");
  } else {
    cell_statistic_value.innerHTML = formatNumber(statistic_value, statistic_name);  
  }
  
  // Append children
  cell_statistic_name.appendChild(insert_link_statistic);
  row_statistic.appendChild(cell_statistic_name);
  row_statistic.appendChild(cell_statistic_value);

  return(row_statistic)
}

// Specific statistics table functions

function createTTestTable (identifier, analysis) {
  // Get the t-test statistics
  var statistics = analysis["statistics"];

  // Create a single statistics table
  var table = createStatisticsTable()

  // Create name row
  var row_name = createNameRow(identifier, analysis.data_name);

  // Create statistics rows
  // t-value, df, and p
  var row_test_statistic = createStatisticsRow(identifier, "t", statistics.statistic.value);
  var row_df = createStatisticsRow(identifier, "df", statistics.df);
  var row_p = createStatisticsRow(identifier, "p", statistics.p);

  // Add rows to table
  table.appendChild(row_name);
  table.appendChild(row_test_statistic);
  table.appendChild(row_df);
  table.appendChild(row_p);

  // CIs
  if ("CI" in statistics) {
    var row_CI_lower = createStatisticsRow(identifier, "CI_lower", statistics.CI.lower, statistics.CI.level);
    var row_CI_upper = createStatisticsRow(identifier, "CI_upper", statistics.CI.upper, statistics.CI.level);

    table.appendChild(row_CI_lower);
    table.appendChild(row_CI_upper);
  }

  return(table)
}

function createPearsonCorrelationTable(identifier, analysis) {
  // Get the correlation statistics
  var statistics = analysis["statistics"];

  // Create a single statistics table
  var table = createStatisticsTable()

  // Create name row
  var row_name = createNameRow(identifier, analysis.data_name);

  // Create statistics rows
  // correlation, t-value, df, and p
  var row_estimate = createStatisticsRow(identifier, "r", statistics.estimate);
  var row_test_statistic = createStatisticsRow(identifier, "t", statistics.statistic.value);
  var row_df = createStatisticsRow(identifier, "df", statistics.df);
  var row_p = createStatisticsRow(identifier, "p", statistics.p);

  // Add rows to table
  table.appendChild(row_name);
  table.appendChild(row_estimate);
  table.appendChild(row_test_statistic);
  table.appendChild(row_df);
  table.appendChild(row_p);

  // CIs
  if ("CI" in statistics) {
    var row_CI_lower = createStatisticsRow(identifier, "CI_lower", statistics.CI.lower, statistics.CI.level);
    var row_CI_upper = createStatisticsRow(identifier, "CI_upper", statistics.CI.upper, statistics.CI.level);

    table.appendChild(row_CI_lower);
    table.appendChild(row_CI_upper);
  }

  return(table)
}

function createKendallCorrelationTable(identifier, analysis) {
  // Get the correlation statistics
  var statistics = analysis["statistics"];

  // Create a single statistics table
  var table = createStatisticsTable()

  // Create name row
  var row_name = createNameRow(identifier, analysis.data_name);

  // Create statistics rows
  // correlation, z-value, p
  var row_estimate = createStatisticsRow(identifier, "tau", statistics.estimate);
  var row_test_statistic = createStatisticsRow(identifier, "z", statistics.statistic.value);
  var row_p = createStatisticsRow(identifier, "p", statistics.p);

  // Add rows to table
  table.appendChild(row_name);
  table.appendChild(row_estimate);
  table.appendChild(row_test_statistic);
  table.appendChild(row_p);

  return(table)
}

function createSpearmanCorrelationTable(identifier, analysis) {
  // Get the correlation statistics
  var statistics = analysis["statistics"];

  // Create a single statistics table
  var table = createStatisticsTable()

  // Create name row
  var row_name = createNameRow(identifier, analysis.data_name);

  // Create statistics rows
  // correlation, S-value, p
  var row_estimate = createStatisticsRow(identifier, "rho", statistics.estimate);
  var row_test_statistic = createStatisticsRow(identifier, "S", statistics.statistic.value);
  var row_p = createStatisticsRow(identifier, "p", statistics.p);

  // Add rows to table
  table.appendChild(row_name);
  table.appendChild(row_estimate);
  table.appendChild(row_test_statistic);
  table.appendChild(row_p);

  return(table)
}

function createChiSquaredTable(identifier, analysis) {

  // Get the chi-squared statistics
  var statistics = analysis["statistics"];

  // Create a single statistics table
  var table = createStatisticsTable()

  // Create name row
  var row_name = createNameRow(identifier, analysis.data_name);

  // Create statistics rows
  // Chi-square, df, p
  var row_test_statistic = createStatisticsRow(identifier, "X-squared", statistics.statistic.value);
  var row_df = createStatisticsRow(identifier, "df", statistics.df);
  var row_p = createStatisticsRow(identifier, "p", statistics.p);
  
  // Add rows to table
  table.appendChild(row_name);
  table.appendChild(row_test_statistic);
  table.appendChild(row_df);
  table.appendChild(row_p);

  return(table)
}

function createWilcoxonTable(identifier, analysis) {
  // Get the Wilcox statistics
  var statistics = analysis["statistics"];

  // Create a single statistics table
  var table = createStatisticsTable()

  // Create name row
  var row_name = createNameRow(identifier, analysis.data_name);

  // Create statistics rows
  // Test statistic, p
  var row_test_statistic = createStatisticsRow(identifier, statistics.statistic.name, statistics.statistic.value);
  var row_p = createStatisticsRow(identifier, "p", statistics.p);

  // Add rows to table
  table.appendChild(row_name);
  table.appendChild(row_test_statistic);
  table.appendChild(row_p);

  return(table)
}

function createFisherExactTestTable(identifier, analysis) {
  // Get the Fisher statistics
  var statistics = analysis["statistics"];

  // Create a single statistics table
  var table = createStatisticsTable()

  // Create name row
  var row_name = createNameRow(identifier, analysis.data_name);

  table.appendChild(row_name);

  // Create statistics rows
  // Check if there is an estimate
  if ("estimate" in statistics) {
    var row_estimate = createStatisticsRow(identifier, "OR", statistics.estimate);

    table.appendChild(row_estimate);
  }

  // p
  var row_p = createStatisticsRow(identifier, "p", statistics.p);
  
  table.appendChild(row_p);

  // CIs
  if ("CI" in statistics) {
    var row_CI_lower = createStatisticsRow(identifier, "CI_lower", statistics.CI.lower, statistics.CI.level);
    var row_CI_upper = createStatisticsRow(identifier, "CI_upper", statistics.CI.upper, statistics.CI.level);

    table.appendChild(row_CI_lower);
    table.appendChild(row_CI_upper);
  }

  return(table)
}

function createOneWayAnalysisOfMeansTable(identifier, analysis) {
  // Get the statistics
  var statistics = analysis["statistics"];

  // Create the table
  var table = createStatisticsTable()

  // Create name row
  var row_name = createNameRow(identifier, analysis.data_name);

  // Create statistic rows
  // Test statistic, num df, den df, p
  var row_test_statistic = createStatisticsRow(identifier, "F", statistics.statistic.value);
  var row_df1 = createStatisticsRow(identifier, "df1", statistics.dfs.numerator_df);
  var row_df2 = createStatisticsRow(identifier, "df2", statistics.dfs.denominator_df);
  var row_p = createStatisticsRow(identifier, "p", statistics.p);

  // Add rows to the table
  table.appendChild(row_name);
  table.appendChild(row_test_statistic);
  table.appendChild(row_df1);
  table.appendChild(row_df2);
  table.appendChild(row_p);

  return(table)
}

function createANOVATable(identifier, analysis, group) {
  // Create the table
  var table = createStatisticsTable()

  // Create factor rows
  var coefficients = analysis.coefficients;

  for (var i in coefficients) {

    var coefficient = coefficients[i];
    
    if (typeof group !== 'undefined') {
    	var identifier_coefficient = identifier + "$" + group + "$" + coefficient.name;
    } else {
    	var identifier_coefficient = identifier + "$" + coefficient.name;	
    }

    // Create name row
    if (coefficient.name != "Residuals") {
      var row_name = createNameRow(identifier_coefficient, coefficient.name);
    } else {
      var row_name = createNameRowWithoutLink(coefficient.name);
    }
    
    // Add statistics rows
    // SS, df, MS
    var row_SS = createStatisticsRow(identifier_coefficient, "SS", coefficient.statistics.SS);
    var row_df = createStatisticsRow(identifier_coefficient, "df", coefficient.statistics.df);
    var row_MS = createStatisticsRow(identifier_coefficient, "MS", coefficient.statistics.MS);
    
    // Add rows
    table.appendChild(row_name);
    table.appendChild(row_SS);
    table.appendChild(row_df);
    table.appendChild(row_MS);

    // Add t and p-value, if this is not the Residual coefficient
    if (coefficient.name != "Residuals") {
      // Statistic row
      var row_test_statistic = createStatisticsRow(identifier_coefficient, "F", coefficient.statistics.statistic.value);
      table.appendChild(row_test_statistic);

      // p row
      var row_p = createStatisticsRow(identifier_coefficient, "p", coefficient.statistics.p);
      table.appendChild(row_p);
    }

  }

  return(table)
}

function createANOVATables(identifier, analysis) {
  var tables = [];

  for (var group in analysis.groups) {
  	var group_name = analysis.groups[group].name;
    var table = createANOVATable(identifier, analysis.groups[group], group_name);
    
    // Add a caption
    var caption = document.createElement("caption");
    caption.innerHTML = "Error: " + analysis.groups[group].name;
    caption.style.textAlign = "left";
    table.appendChild(caption);

    tables.push(table);
  }

  return(tables)
}

function createLinearModelTables(identifier, analysis) {
    
    // Create the coefficients table
    var table_coefficients = createStatisticsTable()

    // Add a caption
    var caption_coefficients = document.createElement("caption");
    caption_coefficients.innerHTML = "<span>Table:</span> Coefficients";
    caption_coefficients.style.textAlign = "left";
    table_coefficients.appendChild(caption_coefficients);

    // Create coefficient rows
    var coefficients = analysis.coefficients;

    for (var i in coefficients) {

        var coefficient = coefficients[i];
        var identifier_coefficient = identifier + "$" + coefficient.name;

        // Create name row
  		  var row_name = createNameRow(identifier_coefficient, coefficient.name);

        // Estimate row
        var row_estimate = createStatisticsRow(identifier_coefficient, "b", coefficient.statistics.estimate);
        var row_SE = createStatisticsRow(identifier_coefficient, "SE", coefficient.statistics.SE);
        var row_test_statistic = createStatisticsRow(identifier_coefficient, "t", coefficient.statistics.statistic.value);
        var row_df = createStatisticsRow(identifier_coefficient, "df", coefficient.statistics.df);
     	  var row_p = createStatisticsRow(identifier_coefficient, "p", coefficient.statistics.p);

        // Add rows to the coefficient table
        table_coefficients.appendChild(row_name);
        table_coefficients.appendChild(row_estimate);
        table_coefficients.appendChild(row_SE);
        table_coefficients.appendChild(row_test_statistic);
        table_coefficients.appendChild(row_df);
        table_coefficients.appendChild(row_p); 
    }

    // Create the model fit table
    var table_model = createStatisticsTable()

    // Add a caption
    var caption_model = document.createElement("caption");
    caption_model.innerHTML = "<span>Table:</span> Model fit";
    caption_model.style.textAlign = "left";
    table_model.appendChild(caption_model);

    var model = analysis.model;
    var identifier_model = identifier + "$model";
    
    // Add name row
    var row_model_name = createNameRow(identifier_model, "Model");

    // Add statistics rows
    var row_test_statistic = createStatisticsRow(identifier_model, "F", model.statistics.statistic.value);
    var row_df1 = createStatisticsRow(identifier_model, "df1", model.statistics.dfs.numerator_df);
    var row_df2 = createStatisticsRow(identifier_model, "df2", model.statistics.dfs.denominator_df);
    var row_p = createStatisticsRow(identifier_model, "p", model.statistics.p);
    var row_r_squared = createStatisticsRow(identifier_model, "R-squared", model.statistics.r_squared);
    var row_adj_r_squared = createStatisticsRow(identifier_model, "adj-R-squared", model.statistics.adjusted_r_squared);

    // Add rows to the model fit table
    table_model.appendChild(row_model_name);
    table_model.appendChild(row_test_statistic);
    table_model.appendChild(row_df1);
    table_model.appendChild(row_df2);
    table_model.appendChild(row_p);
    table_model.appendChild(row_r_squared);
    table_model.appendChild(row_adj_r_squared);

    // Combine tables
    var tables = [];
    tables[0] = table_coefficients;
    tables[1] = table_model

    return(tables)
}

function createLinearMixedModelTables(identifier, analysis) {
    
    // Create the random effects variances table
    var table_random_effects_variances = createStatisticsTable()

    // Add a caption
    var caption_random_effects_variances = document.createElement("caption");
    caption_random_effects_variances.innerHTML = "<span>Table: </span>Random Effects Variances";
    caption_random_effects_variances.style.textAlign = "left";
    table_random_effects_variances.appendChild(caption_random_effects_variances);

    // Get the random effects groups
    var groups = analysis.random_effects.groups;

    // Loop over each group to get the variances
    for (var i in groups) {
      // Get the group statistics
      var group = groups[i]

      var row_group_name = createGroupRow(group.name);
      table_random_effects_variances.appendChild(row_group_name);

      // Get the variances of the terms
      var variances = group.variances;

      for (var j in variances) {
        // Get the term
        var term = variances[j];

        // Set the identifier
        var identifier_term = identifier + "$RE$" + i + "$" + group.name + "$" + j + "$" + term.name;

        // Create name row
        if (group.name != "Residual") {
          var row_name = createNameRowWithoutLink(term.name);
          table_random_effects_variances.appendChild(row_name);
        }

        // Create coefficient statistics rows
        var row_variance = createStatisticsRow(identifier_term, "var", term.statistics.var);
        var row_SD = createStatisticsRow(identifier_term, "SD", term.statistics.SD);
        
        // Add rows
        table_random_effects_variances.appendChild(row_variance);
        table_random_effects_variances.appendChild(row_SD);
      }
    }

    // Create the random effects correlations table, if there are any
    var random_correlations_found = false;
    var table_random_effects_correlations = createStatisticsTable()

    // Add a caption
    var caption_random_effects_correlations = document.createElement("caption");
    caption_random_effects_correlations.innerHTML = "<span>Table: </span>Random Effects Correlations";
    caption_random_effects_correlations.style.textAlign = "left";
    table_random_effects_correlations.appendChild(caption_random_effects_correlations);

    // Get the random effects groups
    var groups = analysis.random_effects.groups;

    // Loop over each group to get the correlations, if there are any
    for (var i in groups) {
      // Get the group statistics
      var group = groups[i]

      if ("correlations" in group) {
        var random_correlations_found = true;

        var row_group_name = createNameRowWithoutLink(group.name);
        table_random_effects_correlations.appendChild(row_group_name);

        // Get the correlations of the terms
        var random_correlations = group.correlations;

        for (var j in random_correlations) {
          // Get the correlation pair
          var pair = random_correlations[j];

          // Create a name for the pair
          var pair_name = pair.names[0] + " - " + pair.names[1]

          // Set the identifier
          var identifier_term = identifier + "$RE$" + i + "$" + group.name + "$" + j + "$" + pair_name;

          // Create name row
          var row_name = createNameRowWithoutLink(pair_name);
          table_random_effects_correlations.appendChild(row_name);

          // Create coefficient statistics rows
          var row_r = createStatisticsRow(identifier_term, "r", pair.statistics.r);
          
          // Add row
          table_random_effects_correlations.appendChild(row_r);
        }
      }
    }

    // Create the fixed effects coefficients table
    var table_fixed_effects_coefficients = createStatisticsTable()

    // Add a caption
    var caption_fixed_effects_coefficients = document.createElement("caption");
    caption_fixed_effects_coefficients.innerHTML = "<span>Table: </span>Fixed Effects Coefficients";
    caption_fixed_effects_coefficients.style.textAlign = "left";
    table_fixed_effects_coefficients.appendChild(caption_fixed_effects_coefficients);

    // Get the coefficient statistics
    var coefficients = analysis.fixed_effects.coefficients;

    // Create coefficient rows
    for (var i in coefficients) {

        var coefficient = coefficients[i];

        // Set the identifier
        var identifier_coefficient = identifier + "$FE$" + i + "$" + coefficient.name;

        // Create name row
        var row_name = createNameRow(identifier_coefficient, coefficient.name);

        // Create coefficient statistics rows
        var row_estimate = createStatisticsRow(identifier_coefficient, "b", coefficient.statistics.estimate);
        var row_SE = createStatisticsRow(identifier_coefficient, "SE", coefficient.statistics.SE);
        var row_test_statistic = createStatisticsRow(identifier_coefficient, "t", coefficient.statistics.statistic.value);
        
        // Add row
        table_fixed_effects_coefficients.appendChild(row_name);
        table_fixed_effects_coefficients.appendChild(row_estimate);
        table_fixed_effects_coefficients.appendChild(row_SE);
        table_fixed_effects_coefficients.appendChild(row_test_statistic);

        // Optionally, add a df and p-value row
        if ("df" in coefficient.statistics) {
          var row_df = createStatisticsRow(identifier_coefficient, "df", coefficient.statistics.df);  
          table_fixed_effects_coefficients.appendChild(row_df);
        }
        
        if ("p" in coefficient.statistics) {
          var row_p = createStatisticsRow(identifier_coefficient, "p", coefficient.statistics.p);  
          table_fixed_effects_coefficients.appendChild(row_p); 
        }
    }

    // Create the fixed effects correlations table, if there are any
    var fixed_correlations_found = false;
    var table_fixed_effects_correlations = createStatisticsTable()

    // Add a caption
    var caption_fixed_effects_correlations = document.createElement("caption");
    caption_fixed_effects_correlations.innerHTML = "<span>Table: </span>Fixed Effects Correlations";
    caption_fixed_effects_correlations.style.textAlign = "left";
    table_fixed_effects_correlations.appendChild(caption_fixed_effects_correlations);

    // Get the coefficient correlations
    if ("correlations" in analysis.fixed_effects) {
      var fixed_correlations_found = true;
      var fixed_correlations = analysis.fixed_effects.correlations;

      // Loop over each correlation
      for (var i in fixed_correlations) {
        // Get the term
        var pair = fixed_correlations[i];
        var pair_name = pair.names[0] + " - " + pair.names[1]
        
        // Set the identifier
        var identifier_term = identifier + "$FE$" + i + "$" + pair_name;

        // Create name row
        var row_name = createNameRowWithoutLink(pair_name);
        table_fixed_effects_correlations.appendChild(row_name);

        // Create coefficient statistics rows
        var row_r = createStatisticsRow(identifier_term, "r", pair.statistics.r);
        
        // Add row
        table_fixed_effects_correlations.appendChild(row_r);
      }
    }

    // Combine tables
    var tables = [];
    tables.push(table_random_effects_variances);
    if (random_correlations_found) {
      tables.push(table_random_effects_correlations);
    }
    tables.push(table_fixed_effects_coefficients);
    if (fixed_correlations_found) {
      tables.push(table_fixed_effects_correlations);
    }

    return(tables)
}

function createDescriptivesTable(identifier, analysis) {

  // Create a single statistics table
  var table = createStatisticsTable()

  // Check if there are any groups
  if ("groups" in analysis) {
    var groups = analysis.groups;

    // Loop over the groups
    for (var i in groups) {
      var group = groups[i];
      var statistics = group.statistics;
      var identifier_group = identifier + "$" + group.name;

      // Create name row
      var row_name = createNameRow(identifier_group, analysis.name + " - " + group.name);
      table.appendChild(row_name);

      // Loop over the statistics and turn each one into a row
      var statistics_names = Object.keys(statistics);

      for (var name in statistics_names) {
        var row_statistic = createStatisticsRow(identifier_group, statistics_names[name], statistics[statistics_names[name]]);
        table.appendChild(row_statistic);
      }
    }
  } else {
      var statistics = analysis.statistics;

      // Create name row
      var row_name = createNameRow(identifier, analysis.name);
      table.appendChild(row_name);

      // Loop over the statistics and turn each one into a row
      var statistics_names = Object.keys(statistics);

      for (var name in statistics_names) {
        var row_statistic = createStatisticsRow(identifier, statistics_names[name], statistics[statistics_names[name]]);
        table.appendChild(row_statistic);
      }
  }
  
  return(table)
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
    var lower = formatNumber(statistics.CI.lower);
    var upper = formatNumber(statistics.CI.upper);

    output = output + ", " + statistics.CI.level * 100 + "% CI [" + lower + 
      ", " + upper + "]"
  }

  return(output)
}

function createPearsonCorrelationLine(statistics) {
  var r, df, p, output;

  r = formatNumber(statistics.estimate, "r");
  df = formatNumber(statistics.df, "df");
  p = formatNumber(statistics.p, "p");

  if (p == "< .001") {
    output = "<i>r</i>(" + df + ") = " + r + ", <i>p</i> " + p;
  } else {
    output = "<i>r</i>(" + df + ") = " + r + ", <i>p</i> = " + p;
  }

  if ("CI" in statistics) {
    var lower = formatNumber(statistics.CI.lower);
    var upper = formatNumber(statistics.CI.upper);

    output = output + ", " + statistics.CI.level * 100 + "% CI [" + lower + 
      ", " + upper + "]"
  }

  return(output) 
}

function createKendallCorrelationLine(statistics) {
  var tau, p, output;

  tau = formatNumber(statistics.estimate, "r");
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

  rho = formatNumber(statistics.estimate, "r");
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
    var OR = formatNumber(statistics.estimate);

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
  df1 = formatNumber(statistics.dfs.numerator_df, "df");
  df2 = formatNumber(statistics.dfs.denominator_df, "df");
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

function createANOVALine(statistics, statistics_residuals) {
  var statistic, df1, df2, p, output;

  statistic = formatNumber(statistics.statistic.value);
  df1 = formatNumber(statistics.df, "df");
  df2 = formatNumber(statistics_residuals.df, "df");
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

function createLinearModelCoefficientLine(statistics) {
  var estimate, SE, statistic, df, p, output;

  estimate = formatNumber(statistics.estimate);
  SE = formatNumber(statistics.SE);
  statistic = formatNumber(statistics.statistic.value);
  df = formatNumber(statistics.df, "df");
  p = formatNumber(statistics.p, "p");

  if (p == "< .001") {
    output = "<i>b</i> = " + estimate + ", <i>SE</i> = " + SE + ", <i>t</i>(" + 
      df + ") = " + statistic + ", <i>p</i> " + p;
  } else {
    output = "<i>b</i> = " + estimate + ", <i>SE</i> = " + SE + ", <i>t</i>(" + 
      df + ") = " + statistic + ", <i>p</i> = " + p;
  }

  return(output) 
}

function createLinearModelModelFitLine(statistics) {
  var r_squared, statistic, df1, df2, p, output;

  r_squared = formatNumber(statistics.r_squared);
  statistic = formatNumber(statistics.statistic.value);
  df1 = formatNumber(statistics.dfs.numerator_df, "df");
  df2 = formatNumber(statistics.dfs.denominator_df, "df");
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

function createLinearMixedModelFixedEffectLine(statistics) {
  var b, SE, t, df, p, output;

  estimate = formatNumber(statistics.estimate);
  SE = formatNumber(statistics.SE);
  statistic = formatNumber(statistics.statistic.value);
  
  if ("df" in statistics) {
    df = formatNumber(statistics.df, "df");
    p = formatNumber(statistics.p, "p");

    if (p == "< .001") {
      output = "<i>b</i> = " + estimate + ", <i>SE</i> = " + SE + ", <i>t</i>(" + 
        df + ") = " + statistic + ", <i>p</i> " + p;
    } else {
      output = "<i>b</i> = " + estimate + ", <i>SE</i> = " + SE + ", <i>t</i>(" + 
        df + ") = " + statistic + ", <i>p</i> = " + p;
    }
  } else {
    if (p == "< .001") {
      output = "<i>b</i> = " + estimate + ", <i>SE</i> = " + SE + ", <i>t</i> = " + statistic;
    } else {
      output = "<i>b</i> = " + estimate + ", <i>SE</i> = " + SE + ", <i>t</i> = " + statistic;
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
