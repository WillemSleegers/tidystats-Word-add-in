function createAnalyses(analyses) {
  var analysesContainer, analysisContainer;

  // Grab the analyses container to add each analysis to
  analysesContainer = document.getElementById("analyses-container");

  // Reset analyses in case a file was already loaded in
  while (analysesContainer.firstChild) {
      analysesContainer.removeChild(analysesContainer.firstChild);
  }

  // Loop over all the analyses and create an analysis element
  for (var identifier in analyses) {
      analysisContainer = createAnalysis(analyses, identifier);
      analysesContainer.appendChild(analysisContainer);
  }

  // Make the main div visible
  document.getElementById("app-main").style.display = "block";
}

function createAnalysis (analyses, identifier) {
  var analysis, analysisContainer, attrs;

  analysis = analyses[identifier];
  
  // Create an analysis container
  analysisContainer = createContainer("container", 1);
  
  // Add the identifier row
  analysisContainer = addIdentifierRow(analysisContainer, identifier);

  // Create an analysis content container
  contentContainer = createContainer("container", 2);

  // Add the method row
  contentContainer = addRow(contentContainer, false, "Method", 
    analysis.method);

  // Add a description row, if there is one
  if ("description" in analysis) {
    contentContainer = addRow(contentContainer, false, "Description", 
      analysis.description);
  }

  // Add the statistics
  // Handle exceptions for each analysis

  // Start with creating an attributes dictionary that will contain the 
  // information to figure out what the user wants to insert
  attrs = {};
  attrs["identifier"] = identifier;

  // Add statistics
  if ("statistics" in analysis) {
    contentContainer = addStatisticsRows(contentContainer, true, 3, 
      analysis.statistics, true, attrs);
  }

  // Add models
  if ("models" in analysis) {
    contentContainer = addModelsRows(contentContainer, 3, 
      analysis.models, attrs);
  }

  // Add effects
  if ("effects" in analysis) {
    // Add random effects
    contentContainer = addRandomEffectsRows(contentContainer, 3, 
      analysis.effects.random_effects, attrs);

    // Add fixed effects
    contentContainer = addFixedEffectsRows(contentContainer, 3, 
      analysis.effects.fixed_effects, attrs);
  }

  // Add groups
  if ("groups" in analysis) {
    contentContainer = addGroupsRows(contentContainer, 3, analysis.groups, 
      attrs);
  }

  // Add terms
  if ("terms" in analysis) {
    contentContainer = addTermsRows(contentContainer, 3, analysis.terms, attrs);
  }

  // Add children to table
  analysisContainer.appendChild(contentContainer);

  return(analysisContainer);
}


function formatName(x, extra) {
  var name;

  // Set extra to '' if no type is provided
  extra = extra || '';

  if (x == "CI_lower") {
    name = extra * 100 + "% CI lower";
  } else if (x == "CI_upper") {
    name = extra * 100 + "% CI upper";
  } else if (x == "df_numerator") {
    name = "num. df";
  } else if (x == "df_denominator") {
    name = "den. df";
  } else if (x == "df_null") {
    name = "null df";
  } else if (x == "df_residual") {
    name = "residual df";
  } else if (x == "cor") {
    name = "r";
  } else if (x == "tau") {
    name = "r<sub>&tau;<sub>";
  } else if (x == "rho") {
    name = "r<sub>S</sub>";
  } else if (x == "X-squared") {
    name = "&chi;²";
  } else if (x == "r_squared") {
    name = "R²";
  } else if (x == "adjusted_r_squared") {
    name = "adj. R²";
  } else if (x == "ges") {
    name = "η²<sub>G</sub>";
  } else if (x == "deviance_null") {
    name = "null deviance";
  } else if (x == "deviance_residual") {
    name = "residual deviance";
  } else {
    name = x;
  }

  return(name)
}

function formatNumber(x, type) {
  var number;

  // Set type to '' if no type is provided
  type = type || '';

  if (x == "Inf") {
    number = "&infin;";
  } else if (x == "-Inf") {
    number = "-&infin;"; // Would be nice to solve this issue with .includes and .replace at some point
  } else {
    number = parseFloat(x);

    const integers = ["df", "df_numerator", "df_denominator", "df_null",
      "df_residual", "N", "n", "missing"];
    const omitZero = ["r", "cor", "tau", "rho", "r_squared", 
      "adjusted_r_squared"];

    if (integers.includes(type)) {
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
      } else if (number == 1) {
          number = number.toFixed(2);
      } else {
          number = number.toFixed(2);
          number = number.toString();
          number = number.substr(1);
      }
    } else if (omitZero.includes(type)) {
      number = number.toFixed(2);
      if (number < 0) {
        number = number.toString();
        number = number.slice(0, 1) + number.slice(2);  
      } else if (number == 1) {
        number = number.toFixed(2);
      } else {
        number = number.toString();
        number = number.substr(1);
      }
    } else {
      if (number >= 1e10 || number <= 1e-10) {
        number = number.toPrecision(3);
      } else {
        number = number.toFixed(2);
      }
    }
  }                

  return(number)
}

function formatCIs(CIs) {
  var text = "[" + formatNumber(CIs.CI_lower) + ", " + formatNumber(CIs.CI_upper) + "]";
  return(text)
}

/* Add row functions */

function createContainer (className, level) {
  var container;

  container = document.createElement("div");
  container.classList.add(className);
  container.classList.add("level" + level);

  return(container)
}

function addRow(element, isParent, name, value) {
  var row, label, valueLabel;

  // Create the row div
  row = document.createElement("div");
  row.classList.add("row");

  label = document.createElement("div");

  if (typeof value !== 'undefined') {
    valueLabel = document.createElement("div");  
  }

  // Make it a parent or child
  if (isParent) {
    row.classList.add("parent");
    row.classList.add("active");
    
    if (typeof value !== 'undefined') {
      label.className = "parent-name";
      valueLabel.className = "parent-value";
    } else {
      label.className = "parent-title";
    }

    // Add the chevron
    row = addChevron(row);
  } else {
    row.classList.add("child");
    label.className = "child-name";

    if (typeof value !== 'undefined') {
      valueLabel.className = "child-value";
    }
  }

  // Set the children labels
  label.innerHTML = name;
  row.appendChild(label);

  if (typeof value !== 'undefined') {
    valueLabel.innerHTML = value;
    row.appendChild(valueLabel);
  }
  
  element.appendChild(row);

  return(element)
}

function addIdentifierRow(element, identifier) {
  var row, label;

  // Create the row div
  row = document.createElement("div");
  row.classList.add("row");
  row.classList.add("parent");
  row.classList.add("identifier");

  // Add the chevron
  row = addChevron(row);

  // Add the identifier label
  label = document.createElement("div");
  label.className = "identifier-label";
  label.innerHTML = identifier;
  row.appendChild(label);

  // Add the row to the element
  element.appendChild(row);

  return(element)
}

function addStatisticRow(element, name, value, attrs, extra) {
  var row, label, valueLabel, insertButton;

  // Create the row div
  row = document.createElement("div");
  row.classList.add("row");
  row.classList.add("child");

  // Add the children labels
  label = document.createElement("div");
  label.className = "child-name";
  label.innerHTML = formatName(name, extra);
  row.appendChild(label);

  valueLabel = document.createElement("div");
  valueLabel.className = "child-value";
  valueLabel.innerHTML = formatNumber(value, name);
  row.appendChild(valueLabel);

  // Add an insert button
  // Set single attribute to true to indicate the user wants to insert a single
  // statistic
  attrs["single"] = true;
  row = addInsertButton(row, attrs);

  element.appendChild(row);

  return(element)
}

function addStatisticsRows(element, isParent, level, statistics, attrs) {
  var statisticsContainer, row, label;

  // Add statistics header row
  row = document.createElement("div");
  row.classList.add("row");

  if (isParent) {
    row.classList.add("parent");
    row.classList.add("active");

    row = addChevron(row);
  } else {
    row.classList.add("child");
  }
  
  label = document.createElement("div");
  label.classList.add("statistics-name");
  label.innerHTML = "Statistics:";
  row.appendChild(label);

  // Set single attribute to false to indicate the user wants to insert multiple
  // statistics
  attrs["single"] = false;
  row = addInsertButton(row, attrs);

  element.appendChild(row);

  // Loop over each statistic, create a row, and add it to the element
  statisticsContainer = createContainer("container", level);

  for (var statistic in statistics) {
    // Handle exception cases: named statistics, multiple dfs, CIs
    if (typeof statistics[statistic] == "object") {
      if ("name" in statistics[statistic]) {
        var statisticsAttrs = {...attrs};
        statisticsAttrs["statistic"] = statistic;

        statisticsContainer = addStatisticRow(
          statisticsContainer, 
          statistics[statistic].name,
          statistics[statistic].value,
          statisticsAttrs
        );  
      } else if (statistic == "CI") {
        for (CI in statistics["CI"]) {
          if (CI != "CI_level") {
            var attrsCI = {...attrs};  
            attrsCI["statistic"] = CI;

            statisticsContainer = addStatisticRow(
              statisticsContainer, 
              CI, 
              statistics[statistic][CI], 
              attrsCI, 
              statistics[statistic].CI_level
            );
          }
        }
      } else if (statistic == "dfs") {
        for (df in statistics["dfs"]) {
          var attrsDf = {...attrs};  
          attrsDf["statistic"] = df;

          statisticsContainer = addStatisticRow(
            statisticsContainer, 
            df,
            statistics[statistic][df],
            attrsDf
          );
        }
      }
    } else {
      var statisticsAttrs = {...attrs};
      statisticsAttrs["statistic"] = statistic;

      statisticsContainer = addStatisticRow(
        statisticsContainer, 
        statistic, 
        statistics[statistic], 
        statisticsAttrs
      );  
    }
  }

  element.appendChild(statisticsContainer);

  return(element)
}

function addTermsRows(element, level, terms, attrs) {
  var term, termsContainer, termContainer, termAttrs;

  // Add a parent row
  console.log(element);
  element = addRow(element, true, "Terms:");
  console.log("This is called");

  // Create a new container
  termsContainer = createContainer("container", level);

  // Loop over the terms
  for (var t in terms) {
    term = terms[t];
    console.log(term);

    // Add name row
    termsContainer = addRow(termsContainer, true, "Name", 
      term.name);

    // Create a new container
    termContainer = createContainer("container", level);

    // Add statistics
    termAttrs = {...attrs};
    termAttrs["term"] = term.name;
    termContainer = addStatisticsRows(termContainer, false, level + 1, 
      term.statistics, termAttrs);  

    termsContainer.appendChild(termContainer);
  }

  element.appendChild(termsContainer);

  return(element)
}

function addPairsRows(element, level, pairs, attrs) {
  var pair, pairsContainer, pairContainer, pairAttrs;

  // Add a parent row
  element = addRow(element, true, "Pairs:");

  // Create a new container
  pairsContainer = createContainer("container", level);

  // Loop over the terms
  for (var p in pairs) {
    pair = pairs[p];

    // Add name row
    name1 = pair.names[0];
    name2 = pair.names[1];
    pairsContainer = addRow(pairsContainer, true, "Name", 
      name1 + " - " + name2);

    // Create a new container
    pairContainer = createContainer("container", level);

    pairAttrs = {...attrs};
    pairAttrs["pair1"] = pair.names[0];
    pairAttrs["pair2"] = pair.names[1];
    pairContainer = addStatisticsRows(pairContainer, false, level + 1, 
      pair.statistics, pairAttrs); 

    pairsContainer.appendChild(pairContainer); 
  }

  element.appendChild(pairsContainer);

  return(element)
}

function addGroupsRows(element, level, groups, attrs) {
  var group, groupsContainer, groupContainer, groupAttrs;

  // Add a parent row
  element = addRow(element, true, "Groups:");

  // Create a new container
  groupsContainer = createContainer("container", level);

  // Loop over the groups
  for (var g in groups) {
    group = groups[g];
    console.log(group);

    // Add name row
    groupsContainer = addRow(groupsContainer, true, "Name", 
      group.name);

    // Add statistics
    if ("statistics" in group) {
      groupAttrs = {...attrs};
      groupAttrs["group"] = group.name;

      // Create a new container
      groupContainer = createContainer("container", level);

      groupContainer = addStatisticsRows(groupContainer, false, level + 1, 
        group.statistics, groupAttrs);
    }

    // Add terms
    if ("terms" in group) {
      groupAttrs = {...attrs};
      groupAttrs["group"] = group.name;
      
      // Create a new container
      groupContainer = createContainer("container", level);

      groupContainer = addTermsRows(groupContainer, level + 1, 
        group.terms, groupAttrs);
    }

    if ("pairs" in group) {
      groupAttrs = {...attrs};
      groupAttrs["group"] = group.name;

      // Create a new container
      groupContainer = createContainer("container", level);

      groupContainer = addPairsRows(groupContainer, level + 1, 
        group.pairs, groupAttrs);
    }

    groupsContainer.appendChild(groupContainer);
  }
    
  element.appendChild(groupsContainer);

  return(element)
}

function addRandomEffectsRows(element, level, random_effects, attrs) {
  var randomContainer, statisticsAttrs, groupsAttrs;

  attrs["effect"] = "random_effect";

  // Add a parent row
  element = addRow(element, true, "Random effects:");

  // Create a new container
  randomContainer = createContainer("random-container", level);

  // Add statistics
  statisticsAttrs = {...attrs};
  randomContainer = addStatisticsRows(randomContainer, true, level + 1, 
    random_effects.statistics, statisticsAttrs);

  // Add groups
  groupsAttrs = {...attrs};
  randomContainer = addGroupsRows(randomContainer, level + 1, 
    random_effects.groups, groupsAttrs);

  element.appendChild(randomContainer);

  return(element)
}

function addFixedEffectsRows(element, level, fixed_effects, attrs) {
  var fixedContainer, termsAttrs;

  attrs["effect"] = "fixed_effect";

  // Add a parent row
  element = addRow(element, true, "Fixed effects:");

  // Create a new container
  fixedContainer = createContainer("container", level);

  // Add terms
  termsAttrs = {...attrs};
  fixedContainer = addTermsRows(fixedContainer, level + 1, fixed_effects.terms, 
    termsAttrs);
  
  // Add pairs
  if ("pairs" in fixed_effects) {
    var pairsAttrs = {...attrs};
    fixedContainer = addPairsRows(fixedContainer, level + 1, 
      fixed_effects.pairs, pairsAttrs);  
  }

  element.appendChild(fixedContainer);

  return(element)
}

function addModelsRows(element, level, models, attrs) {
  var model, modelsContainer, modelAttrs;

  // Add a parent row
  element = addRow(element, true, "Models:");

  // Create a new container
  modelsContainer = createContainer("container", level);

  // Loop over the models
  for (var m in models) {
    model = models[m];

    // Add name row
    modelsContainer = addRow(modelsContainer, false, "Name", model.name);

    // Add statistics
    modelAttrs = {...attrs};
    modelAttrs["model"] = model.name;
    modelsContainer = addStatisticsRows(modelsContainer, true, level + 1, 
      model.statistics, modelAttrs);  
  }

  element.appendChild(modelsContainer);

  return(element)
}

/* Row helper functions */ 

function addChevron(row) {
  var chevron, chevronRight, chevronDown;

  chevron = document.createElement("div");
  chevron.className = "chevron";
  chevron.addEventListener("click", collapse);

  chevronRight = document.createElement("img");
  chevronRight.className = "chevron-right";
  chevronRight.src = "assets/chevron-right.svg"

  chevronDown = document.createElement("img");
  chevronDown.className = "chevron-down";
  chevronDown.src = "assets/chevron-down.svg"

  chevron.appendChild(chevronRight);
  chevron.appendChild(chevronDown);

  row.appendChild(chevron);

  return(row);
}

function addInsertButton(row, attrs) {
  var insertButton;

  insertButton = document.createElement("div");
  insertButton.className = "insert-button";
  insertButton.innerHTML = "+";
  insertButton.onclick = insertStatistics;

  for (var key in attrs) {
    // console.log("key: " + key);
    // console.log("value: " + attrs[key]);
    insertButton.setAttribute(key, attrs[key]);
  }

  row.appendChild(insertButton);

  return(row);
}

/* Event functions */

function collapse() {
  var parent, content;

  parent = this.parentElement;
  content = parent.nextElementSibling;
  
  if (parent.classList.contains("identifier")) {
    if (content.style.display != "block") {
      content.style.display = "block";
    } else {
      content.style.display = "none";
    }  
  } else {
    if (content.style.display != "none") {
      content.style.display = "none";
    } else {
      content.style.display = "block";
    }
  }

  parent.classList.toggle("active"); 
}