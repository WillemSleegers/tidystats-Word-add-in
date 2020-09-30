function insertStatistic() {
  console.log("Inserting single statistic");
  
  var element = this;
  var attributes = {};

  attributes["identifier"] = element.getAttribute("identifier");

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

  // Insert statistic in Word
  insert(attributes);
}

function insertStatistics() {
  console.log("Inserting multiple statistics");
  
  // Get statistics row that was clicked on
  var row = this;
  
  // Get the statistics rows
  var rows = row.nextSibling.children;
  
  // Determine which statistics are selected
  var selectedStatistics = [];
  
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].getElementsByClassName("checkbox-selected").length) {
      selectedStatistics.push(rows[i].getAttribute("statistic"));
    }
  }
  
  // Determine attributes
  var element = rows[1];
  var attributes = {};
  
  attributes["identifier"] = element.getAttribute("identifier");

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
  
  attributes["statistics"] = selectedStatistics;
  
  // Insert statistics in Word
  insert(attributes);
}


function toggleStatistic() {
  event.stopPropagation();
  console.log("Toggling statistic");
  console.log(this);
  this.firstChild.classList.toggle("checkbox-selected");
  
  // If the statistic is confidence interval, also toggle its paired value
  if (this.parentElement.getAttribute("statistic") == "CI_lower") {
    this.parentElement.nextSibling.getElementsByClassName("checkbox")[0].classList.toggle("checkbox-selected");
  } else if (this.parentElement.getAttribute("statistic") == "CI_upper") {
    this.parentElement.previousSibling.getElementsByClassName("checkbox")[0].classList.toggle("checkbox-selected");
  }
  
  // If the statistic is numerator or denominator df, also toggle its paired value
  if (this.parentElement.getAttribute("statistic") == "df_numerator") {
    this.parentElement.nextSibling.getElementsByClassName("checkbox")[0].classList.toggle("checkbox-selected");
  } else if (this.parentElement.getAttribute("statistic") == "df_denominator") {
    this.parentElement.previousSibling.getElementsByClassName("checkbox")[0].classList.toggle("checkbox-selected");
  }
}

function toggleToggles() {
  event.stopPropagation();
  
  // Get the gear element and toggle its active state
  this.classList.toggle("gear-active");
  
  // Get rows
  var rows = this.parentElement.nextSibling;
  
  // Find all toggles
  var toggles = rows.getElementsByClassName("checkbox-container");
  
  // Loop over the toggles and toggle their visibility
  for (var i = 0; i < toggles.length; i++) {
    var toggle = toggles[i];
    
    if (toggle.style.display == "flex") {
      toggle.style.display = "none";
    } else {
      toggle.style.display = "flex";
    }
  }
}

function test() {
  console.log("test");
}

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
    // Add name row
    if ("name" in analysis) {
      contentContainer = addRow(contentContainer, false, "Name", 
        analysis.name);  
    }

    contentContainer = addStatisticsRows(contentContainer, true, 3, 
      analysis.statistics, attrs);
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
    contentContainer = addTermsRows(contentContainer, true, 3, analysis.terms, attrs);
  }

  // Add children to table
  analysisContainer.appendChild(contentContainer);

  return(analysisContainer);
}


function formatName(x, extra) {
  var name;

  // Set extra to '' if no type is provided
  extra = extra || '';
  
  switch (x) {
    case "CI_lower":
      name = extra * 100 + "% CI lower";
      break;
    case "CI_upper":
      name = extra * 100 + "% CI upper";
      break;
    case "df_numerator":
      name = "num. df";
      break;
    case "df_denominator":
      name = "den. df";
      break;
    case "df_null":
      name = "null df";
      break;
    case "df_residual":
      name = "residual df";
      break;
    case "cor":
      name = "r";
      break;
    case "tau":
      name = "r<sub>&tau;</sub>";
      break;
    case "rho":
      name = "r<sub>S</sub>";
      break;
    case "X-squared":
      name = "&chi;²";
      break;
    case "r_squared":
      name = "R²";
      break;
    case "adjusted_r_squared":
      name = "adj. R²";
      break;
    case "ges":
      name = "η²<sub>G</sub>";
      break;
    case "deviance_null":
      name = "null deviance";
      break;
    case "deviance_residual":
      name = "residual deviance";
      break;
    case "BF_01":
      name = "BF<sub>01</sub>";
      break;
    case "BF_10":
      name = "BF<sub>10</sub>";
      break;
    case "mean":
      name = "M";
      break;
    case "mean difference":
      name = "M<sub>diff</sub>";
      break;
    case "difference in location":
      name = "Mdn<sub>diff</sub>";
      break;
    case "odds ratio":
      name = "OR"
      break;
    default:
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
      if (number == 0) {
        number = number.toFixed(0);
      } else if (Math.abs(number) >= 1e8 || Math.abs(number) <= 0.00001) {
        number = number.toExponential(2);
      } else if (Math.abs(number) <= 0.001) {
        number = number.toPrecision(2);
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
  row.classList.add("statistic");
  
  // Add insert statistics functionality
  row = addInsertClick(row, attrs);

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
  row = addCheckbox(row, attrs);

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

  // Check if there is more than 1 statistic in statistics
  // If so, add settings and make the row clickable to insert multiple statistics
  if (Object.keys(statistics).length > 1) {
    row.classList.add("statistics");
    row = addSettings(row);
    row.onclick = insertStatistics;  
  }
  
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

function addTermsRows(element, isParent, level, terms, attrs) {
  var term, termsContainer, termContainer, termAttrs;

  // Add a parent row
  element = addRow(element, isParent, "Terms:");

  // Create a new container
  termsContainer = createContainer("container", level);

  // Loop over the terms
  for (var t in terms) {
    term = terms[t];

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

function addPairsRows(element, isParent, level, pairs, attrs) {
  var pair, pairsContainer, pairContainer, pairAttrs;

  // Add a parent row
  element = addRow(element, isParent, "Pairs:");

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

    // Add name row
    groupsContainer = addRow(groupsContainer, true, "Name", 
      group.name);
      
    // Create a new container
    groupContainer = createContainer("container", level + 1);

    // Add statistics
    if ("statistics" in group) {
      groupAttrs = {...attrs};
      groupAttrs["group"] = group.name;
    
      groupContainer = addStatisticsRows(groupContainer, true, level + 2, 
        group.statistics, groupAttrs);
    }
    
    if ("terms" in group || "pairs" in group) {
      // Add terms
      if ("terms" in group) { 
        groupAttrs = {...attrs};
        groupAttrs["group"] = group.name;
        
        groupContainer = addTermsRows(groupContainer, true, level + 2, 
          group.terms, groupAttrs);
      }
      
      // Add pairs
      if ("pairs" in group) {
        groupAttrs = {...attrs};
        groupAttrs["group"] = group.name;
  
        groupContainer = addPairsRows(groupContainer, true, level + 2, 
          group.pairs, groupAttrs);
      }
    }
    
    groupsContainer.appendChild(groupContainer);
    
    element.appendChild(groupsContainer);
  }

  return(element)
}

function addRandomEffectsRows(element, level, random_effects, attrs) {
  var randomContainer, statisticsAttrs, groupsAttrs;

  attrs["effect"] = "random_effects";

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

  attrs["effect"] = "fixed_effects";

  // Add a parent row
  element = addRow(element, true, "Fixed effects:");

  // Create a new container
  fixedContainer = createContainer("container", level);

  // Add terms
  termsAttrs = {...attrs};
  fixedContainer = addTermsRows(fixedContainer, true, level + 1, fixed_effects.terms, 
    termsAttrs);
  
  // Add pairs
  if ("pairs" in fixed_effects) {
    var pairsAttrs = {...attrs};
    fixedContainer = addPairsRows(fixedContainer, true, level + 1, 
      fixed_effects.pairs, pairsAttrs);  
  }

  element.appendChild(fixedContainer);

  return(element)
}

function addModelsRows(element, level, models, attrs) {
  var model, modelsContainer, modelContainer, modelAttrs;
  
  // Add a parent row
  element = addRow(element, true, "Models:");

  // Create a new container
  modelsContainer = createContainer("container", level);

  // Loop over the models
  for (var m in models) {
    model = models[m];
    
    // Add name row
    modelsContainer = addRow(modelsContainer, true, "Name", model.name);
    
    // Create a new container
    modelContainer = createContainer("container", level);

    // Add statistics
    modelAttrs = {...attrs};
    modelAttrs["model"] = model.name;
    modelContainer = addStatisticsRows(modelContainer, false, level + 1, 
      model.statistics, modelAttrs);
      
    modelsContainer.appendChild(modelContainer); 
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
  chevronRight.src = "assets/icons/chevron-right.svg"

  chevronDown = document.createElement("img");
  chevronDown.className = "chevron-down";
  chevronDown.src = "assets/icons/chevron-down.svg"

  chevron.appendChild(chevronRight);
  chevron.appendChild(chevronDown);

  row.appendChild(chevron);

  return(row);
}

function addInsertClick(row, attrs) {

  row.onclick = insertStatistic;

  for (var key in attrs) {
    // console.log("key: " + key);
    // console.log("value: " + attrs[key]);
    row.setAttribute(key, attrs[key]);
  }
  
  return(row);
}

function addSettings(row) {
  var gear;
  
  gear = document.createElement("div");
  gear.className = "gear-container";
  gear = addGear(gear);
  
  gear.onclick = toggleToggles;
  
  row.appendChild(gear);
  
  return(row);
}

function addCheckbox(row, attrs) {
  var div, checkbox, checkmark;

  div = document.createElement("div");
  div.className = "checkbox-container";
  div.onclick = toggleStatistic;

  checkbox = document.createElement("div");
  checkbox.className = "checkbox";
  checkbox.classList.add("checkbox-selected");
  checkbox = addCheckmark(checkbox);
  
  div.appendChild(checkbox);

  row.appendChild(div);

  return(row);
}

/* Event functions */

function collapse() {
  event.stopPropagation();
  
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




// SVG functions

function addCheckmark(element) {
  var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  var path = document.createElementNS("http://www.w3.org/2000/svg", 'path');
  
  svg.setAttribute('width', '1em');
  svg.setAttribute('height', '1em');
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute("class", "checkmark");
  svg.setAttribute("aria-hidden", "true");
  
  path.setAttribute('d', 'M13.854 3.646a.5.5 0 0 1 0 .708l-7 7a.5.5 0 0 1-.708 0l-3.5-3.5a.5.5 0 1 1 .708-.708L6.5 10.293l6.646-6.647a.5.5 0 0 1 .708 0z');
  path.setAttribute('fill-rule', 'evenodd');
  
  svg.appendChild(path);
  element.appendChild(svg);
  
  return(element);
}

function addGear(element) {
  var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  var path = document.createElementNS("http://www.w3.org/2000/svg", 'path');
  
  svg.setAttribute('width', '1em');
  svg.setAttribute('height', '1em');
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute("class", "gear");
  svg.setAttribute("aria-hidden", "true");
  
  path.setAttribute('d', 'M9.405 1.05c-.413-1.4-2.397-1.4-2.81 0l-.1.34a1.464 1.464 0 0 1-2.105.872l-.31-.17c-1.283-.698-2.686.705-1.987 1.987l.169.311c.446.82.023 1.841-.872 2.105l-.34.1c-1.4.413-1.4 2.397 0 2.81l.34.1a1.464 1.464 0 0 1 .872 2.105l-.17.31c-.698 1.283.705 2.686 1.987 1.987l.311-.169a1.464 1.464 0 0 1 2.105.872l.1.34c.413 1.4 2.397 1.4 2.81 0l.1-.34a1.464 1.464 0 0 1 2.105-.872l.31.17c1.283.698 2.686-.705 1.987-1.987l-.169-.311a1.464 1.464 0 0 1 .872-2.105l.34-.1c1.4-.413 1.4-2.397 0-2.81l-.34-.1a1.464 1.464 0 0 1-.872-2.105l.17-.31c.698-1.283-.705-2.686-1.987-1.987l-.311.169a1.464 1.464 0 0 1-2.105-.872l-.1-.34zM8 10.93a2.929 2.929 0 1 0 0-5.86 2.929 2.929 0 0 0 0 5.858z');
  path.setAttribute('fill-rule', 'evenodd');
  
  svg.appendChild(path);
  element.appendChild(svg);
  
  return(element);
  
}