// Work in progress ----------------------------------------------------

function createTTestWordTable(body, identifier, analysis) {
  var statistics = analysis.statistics;
  var data = [
    ["", "t", "df", "p", "95% CI"],
    ["name", "t-value", "df", "p-value", "CIs"],
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
  content_control_t.insertHtml(
    formatNumber(statistics.statistic.value),
    "Replace"
  );

  content_control_df.tag = identifier + "$df";
  content_control_df.title = identifier + "$df";
  content_control_df.insertHtml(formatNumber(statistics.df, "df"), "Replace");

  content_control_p.tag = identifier + "$p";
  content_control_p.title = identifier + "$p";
  content_control_p.insertHtml(formatNumber(statistics.p, "p"), "Replace");

  content_control_CI.tag = identifier + "$CI";
  content_control_CI.title = identifier + "$CI";
  content_control_CI.insertHtml(formatCIs(statistics.CI), "Replace");

  return table;
}

function createRegressionWordTable(body, identifier, analysis) {
  var statistics = analysis.statistics;
  var data = [["", "b", "SE", "t", "df", "p"]];

  var coefficients = analysis.coefficients;

  for (i in coefficients) {
    data.push([
      "coefficient",
      "b-value",
      "SE-value",
      "t-value",
      "df-value",
      "p-value",
    ]);
  }

  var table = body.insertTable(
    data.length,
    data[0].length,
    Word.InsertLocation.end,
    data
  );

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
    content_control_t.insertHtml(
      formatNumber(statistics.statistic.value),
      "Replace"
    );

    content_control_df.tag = identifier + "$" + name + "$df";
    content_control_df.title = identifier + "$" + name + "$df";
    content_control_df.insertHtml(formatNumber(statistics.df, "df"), "Replace");

    content_control_p.tag = identifier + "$" + name + "$p";
    content_control_p.title = identifier + "$" + name + "$p";
    content_control_p.insertHtml(formatNumber(statistics.p, "p"), "Replace");
  }

  return table;
}

// Debugging
var testAnalysis =
  '{"t_test_one_sample":{"method":"One Sample t-test", "description":"A simple t-test that I ran to test tidystats out.", "name":"cox$call_parent","statistics":{"estimate":25.775,"SE":1.064,"statistic":{"name":"t","value":24.2248},"df":199,"p":1.4581E-61,"CI":{"level":0.95,"lower":24.0167,"upper":"Inf"}},"alternative":{"direction":"greater","mean":0},"package":{"name":"stats","version":"3.6.1"},"notes":"A one-sample t-test on call_parent"}}';

function test() {
  var analysisDiv = document.getElementById("analyses-container");

  testAnalysis = JSON.parse(testAnalysis);

  var test2 = createAnalysis(testAnalysis, "t_test_one_sample");

  analysisDiv.appendChild(test2);
}

var instantly_load_data = false;

function instantlyLoadData() {
  console.log("Instantly loading data");

  var text =
    '{"m":{"method":"Linear mixed model","REML_criterion_at_convergence":1743.6283,"convergence_code":0,"random_effects":{"N":180,"groups":[{"name":"Subject","N":18,"variances":[{"name":"(Intercept)","statistics":{"var":611.8976,"SD":24.7366}},{"name":"Days","statistics":{"var":35.0811,"SD":5.9229}}],"correlations":[{"names":["(Intercept)","Days"],"statistics":{"r":0.0656}}]},{"name":"Residual","variances":[{"statistics":{"var":654.9408,"SD":25.5918}}]}]},"fixed_effects":{"coefficients":[{"name":"(Intercept)","statistics":{"estimate":251.4051,"SE":6.8238,"df":17.0052,"statistic":{"name":"t","value":36.8425},"p":1.1582E-17}},{"name":"Days","statistics":{"estimate":10.4673,"SE":1.546,"df":16.9953,"statistic":{"name":"t","value":6.7707},"p":3.273E-6}}],"correlations":[{"names":["(Intercept)","Days"],"statistics":{"r":-0.1375}}]},"package":{"name":"lme4","version":"1.1-21"}},"fm":{"method":"Linear mixed model","REML_criterion_at_convergence":2705.5037,"convergence_code":0,"random_effects":{"N":648,"groups":[{"name":"Consumer:Product","N":324,"variances":[{"name":"(Intercept)","statistics":{"var":3.1622,"SD":1.7783}}]},{"name":"Consumer","N":81,"variances":[{"name":"(Intercept)","statistics":{"var":0.3756,"SD":0.6129}}]},{"name":"Residual","variances":[{"statistics":{"var":1.6675,"SD":1.2913}}]}]},"fixed_effects":{"coefficients":[{"name":"(Intercept)","statistics":{"estimate":5.849,"SE":0.2843,"df":322.3361,"statistic":{"name":"t","value":20.5742},"p":1.1733E-60}},{"name":"Gender2","statistics":{"estimate":-0.2443,"SE":0.2606,"df":79,"statistic":{"name":"t","value":-0.9375},"p":0.3514}},{"name":"Information2","statistics":{"estimate":0.1605,"SE":0.2029,"df":320.0004,"statistic":{"name":"t","value":0.791},"p":0.4296}},{"name":"Product2","statistics":{"estimate":-0.8272,"SE":0.3453,"df":339.5108,"statistic":{"name":"t","value":-2.3953},"p":0.0171}},{"name":"Product3","statistics":{"estimate":0.1481,"SE":0.3453,"df":339.5108,"statistic":{"name":"t","value":0.429},"p":0.6682}},{"name":"Product4","statistics":{"estimate":0.2963,"SE":0.3453,"df":339.5108,"statistic":{"name":"t","value":0.858},"p":0.3915}},{"name":"Information2:Product2","statistics":{"estimate":0.2469,"SE":0.287,"df":320.0004,"statistic":{"name":"t","value":0.8605},"p":0.3902}},{"name":"Information2:Product3","statistics":{"estimate":0.2716,"SE":0.287,"df":320.0004,"statistic":{"name":"t","value":0.9465},"p":0.3446}},{"name":"Information2:Product4","statistics":{"estimate":-0.358,"SE":0.287,"df":320.0004,"statistic":{"name":"t","value":-1.2477},"p":0.2131}}],"correlations":[{"names":["(Intercept)","Gender2"],"statistics":{"r":-0.4526}},{"names":["(Intercept)","Information2"],"statistics":{"r":-0.3569}},{"names":["(Intercept)","Product2"],"statistics":{"r":-0.6074}},{"names":["(Intercept)","Product3"],"statistics":{"r":-0.6074}},{"names":["(Intercept)","Product4"],"statistics":{"r":-0.6074}},{"names":["(Intercept)","Information2:Product2"],"statistics":{"r":0.2523}},{"names":["(Intercept)","Information2:Product3"],"statistics":{"r":0.2523}},{"names":["(Intercept)","Information2:Product4"],"statistics":{"r":0.2523}},{"names":["Gender2","Information2"],"statistics":{"r":-1.1265E-16}},{"names":["Gender2","Product2"],"statistics":{"r":1.3608E-14}},{"names":["Gender2","Product3"],"statistics":{"r":1.3608E-14}},{"names":["Gender2","Product4"],"statistics":{"r":1.3608E-14}},{"names":["Gender2","Information2:Product2"],"statistics":{"r":7.9653E-17}},{"names":["Gender2","Information2:Product3"],"statistics":{"r":7.9653E-17}},{"names":["Gender2","Information2:Product4"],"statistics":{"r":7.9653E-17}},{"names":["Information2","Product2"],"statistics":{"r":0.2938}},{"names":["Information2","Product3"],"statistics":{"r":0.2938}},{"names":["Information2","Product4"],"statistics":{"r":0.2938}},{"names":["Information2","Information2:Product2"],"statistics":{"r":-0.7071}},{"names":["Information2","Information2:Product3"],"statistics":{"r":-0.7071}},{"names":["Information2","Information2:Product4"],"statistics":{"r":-0.7071}},{"names":["Product2","Product3"],"statistics":{"r":0.5}},{"names":["Product2","Product4"],"statistics":{"r":0.5}},{"names":["Product2","Information2:Product2"],"statistics":{"r":-0.4155}},{"names":["Product2","Information2:Product3"],"statistics":{"r":-0.2077}},{"names":["Product2","Information2:Product4"],"statistics":{"r":-0.2077}},{"names":["Product3","Product4"],"statistics":{"r":0.5}},{"names":["Product3","Information2:Product2"],"statistics":{"r":-0.2077}},{"names":["Product3","Information2:Product3"],"statistics":{"r":-0.4155}},{"names":["Product3","Information2:Product4"],"statistics":{"r":-0.2077}},{"names":["Product4","Information2:Product2"],"statistics":{"r":-0.2077}},{"names":["Product4","Information2:Product3"],"statistics":{"r":-0.2077}},{"names":["Product4","Information2:Product4"],"statistics":{"r":-0.4155}},{"names":["Information2:Product2","Information2:Product3"],"statistics":{"r":0.5}},{"names":["Information2:Product2","Information2:Product4"],"statistics":{"r":0.5}},{"names":["Information2:Product3","Information2:Product4"],"statistics":{"r":0.5}}]},"package":{"name":"lme4","version":"1.1-21"}}}';

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
  var analysis = data[identifier];
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
  });
}
