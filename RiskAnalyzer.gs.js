function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Risk Analysis')
    .addItem('Calculate Sum of Range of Cost', 'showCostSidebar')
    .addItem('Calculate Self-Build Risk', 'showRiskSidebar')
    .addItem('Input Data', 'showFinancialSidebar')
    .addItem('Input OEM Cost', 'showOEMCostSidebar') // New item for OEM Cost
    .addItem('Calculate Risks', 'calculateRisks')
    .addToUi();
}

function showCostSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('CostSidebar')
    .setTitle('Cost Data Entry')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showRiskSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('RiskSidebar')
    .setTitle('Self-Build Risk Analysis')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showFinancialSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('FinancialSidebar')
    .setTitle('Financial Data Input')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// New function to show OEM Cost Sidebar
function showOEMCostSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('OEMCostSidebar')
    .setTitle('OEM Cost Data Entry')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function setupSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Input');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Data Input');
  } else {
    sheet.clear();
  }
  
  // Risk Factors and Weights
  sheet.getRange('A1').setValue('Risk Factor');
  sheet.getRange('B1').setValue('Weight');
  sheet.getRange('C1').setValue('Build Scores');
  var riskFactors = ['Cost', 'Time', 'Quality', 'Supply Chain', 'Scalability', 'Expertise', 'Market Risk', 'Technology Risk'];
  var weights = [0.2, 0.15, 0.15, 0.1, 0.1, 0.1, 0.1, 0.1];
  
  for (var i = 0; i < riskFactors.length; i++) {
    sheet.getRange(i + 2, 1).setValue(riskFactors[i]);
    sheet.getRange(i + 2, 2).setValue(weights[i]);
  }
  
  // Cost Attributes
  sheet.getRange('E1').setValue('Attributes');
  sheet.getRange('F1').setValue('Cost / month ($)');
  var attributes = ['Land and Buildings', 'Machinery and Equipment', 'Technology and Software', 'Raw Materials', 'Labor Costs', 'Utilities', 'Maintenance and Repairs', 'Quality Control', 'Overhead Costs', 'Inventory Costs', 'Depreciation', 'Logistics and Distribution', 'Licenses and Permits', 'Environmental Compliance', 'Health and Safety', 'Interest and Loan Payments', 'Property and Casualty Insurance', 'Liability Insurance', 'Product Development', 'Process Improvement', 'Supplier Relationships', 'Supply Chain Management'];
  var costs = [10000, 6000, 500, 150000, 0, 1000, 1250, 0, 2000, 500, 0, 69778, 260, 678, 556.2, 7880, 7000, 1633, 11230.02, 12352, 1322, 4562];
  
  for (var j = 0; j < attributes.length; j++) {
    sheet.getRange(j + 2, 5).setValue(attributes[j]);
    sheet.getRange(j + 2, 6).setValue(costs[j]);
  }
  
  // Output Cells
  sheet.getRange('I1').setValue('Total Cost');
  sheet.getRange('I2').setValue('Build Average Rating');
  sheet.getRange('I3').setValue('Recommendation');
  sheet.getRange('G1').setValue('Beta Value');
  sheet.getRange('G2').setValue('OEM Cost'); // Cell for OEM Cost

  // Formula to calculate the total cost dynamically
  calculateSumAndDisplay('F2:F23', 'I1');
}

function onEdit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Input');
  if (e.range.getSheet().getName() === 'Data Input' && e.range.getColumn() === 6) {
    calculateSumAndDisplay('F2:F23', 'I1');
  }
}

function processRiskData(formData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Input');
  var riskFactors = ['cost', 'time', 'quality', 'supplyChain', 'scalability', 'expertise', 'marketRisk', 'technologyRisk'];
  var buildScores = [];

  for (var i = 0; i < riskFactors.length; i++) {
    var score = parseInt(formData[riskFactors[i]]);
    buildScores.push(score);
  }

  for (var j = 0; j < buildScores.length; j++) {
    sheet.getRange(j + 2, 3).setValue(buildScores[j]);
  }

  calculateRiskScores();
}

function calculateRiskScores() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Input');
  var weights = [0.2, 0.15, 0.15, 0.1, 0.1, 0.1, 0.1, 0.1];
  var buildTotal = 0;

  for (var i = 0; i < weights.length; i++) {
    var buildScore = sheet.getRange(i + 2, 3).getValue();
    buildTotal += buildScore * weights[i];
  }

  sheet.getRange('I2').setValue(buildTotal);

  var totalCost = sheet.getRange('I1').getValue(); // Total cost is dynamically calculated in cell I1

  var recommendation;
  if (buildTotal < 2) {
    recommendation = 'Strongly Recommended to Build Your Own Business';
  } else if (buildTotal < 3) {
    recommendation = 'Recommended to Build Your Own Business';
  } else if (buildTotal < 4) {
    recommendation = 'Neutral';
  } else if (buildTotal < 5) {
    recommendation = 'Not Recommended to Build Your Own Business';
  } else {
    recommendation = 'Strongly Not Recommended to Build Your Own Business';
  }

  sheet.getRange('I3').setValue(recommendation);
}

function calculateSumAndDisplay(selectedRange, displayRange) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange(selectedRange);
  var values = range.getValues();
  
  var sum = values.reduce(function(acc, row) {
    return acc + row.reduce(function(rowAcc, cell) {
      return rowAcc + (isNaN(cell) ? 0 : cell);
    }, 0);
  }, 0);

  sheet.getRange(displayRange).setValue(sum);
  return sum;
}

function getActiveRange() {
  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveRange().getA1Notation();
  return range;
}

function fetchData(companyTicker, startDate, endDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Input');

  // Fetch company data
  var companyDataFormula = '=GOOGLEFINANCE("' + companyTicker + '", "close", DATE(' + startDate.split('-')[0] + ',' + startDate.split('-')[1] + ',' + startDate.split('-')[2] + '), DATE(' + endDate.split('-')[0] + ',' + endDate.split('-')[1] + ',' + endDate.split('-')[2] + '), "DAILY")';
  sheet.getRange("A26").setFormula(companyDataFormula);

  // Fetch S&P 500 data
  var sp500DataFormula = '=GOOGLEFINANCE("INDEXSP:.INX", "close", DATE(' + startDate.split('-')[0] + ',' + startDate.split('-')[1] + ',' + startDate.split('-')[2] + '), DATE(' + endDate.split('-')[0] + ',' + endDate.split('-')[1] + ',' + endDate.split('-')[2] + '), "DAILY")';
  sheet.getRange("D26").setFormula(sp500DataFormula);

  // Wait for data to be fetched
  SpreadsheetApp.flush();
}

function calculateRisks() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var companyTicker = sheet.getRange("A1").getValue();
    var startDate = sheet.getRange("A2").getValue();
    var endDate = sheet.getRange("A3").getValue();

    fetchData(companyTicker, startDate, endDate);

    var companyDataRange = sheet.getRange("A27:A");
    var companyDataValues = companyDataRange.getValues();
    var sp500DataRange = sheet.getRange("D27:D");
    var sp500DataValues = sp500DataRange.getValues();

    var returns = [];
    var marketReturns = [];
    var n = companyDataValues.length;

    for (var i = 1; i < n - 1; i++) {
      var companyReturn = (companyDataValues[i][0] - companyDataValues[i - 1][0]) / companyDataValues[i - 1][0];
      var marketReturn = (sp500DataValues[i][0] - sp500DataValues[i - 1][0]) / sp500DataValues[i - 1][0];
      returns.push(companyReturn);
      marketReturns.push(marketReturn);
    }

    var covariance = 0;
    var marketVariance = 0;

    for (var j = 0; j < returns.length; j++) {
      covariance += (returns[j] - average(returns)) * (marketReturns[j] - average(marketReturns));
      marketVariance += Math.pow(marketReturns[j] - average(marketReturns), 2);
    }

    covariance /= returns.length - 1;
    marketVariance /= marketReturns.length - 1;

    var beta = covariance / marketVariance;
    sheet.getRange('G1').setValue(beta); // Store beta in the appropriate cell

    // OEM Cost Input
    var oemCost = sheet.getRange('G2').getValue(); // Assuming G2 is the cell where OEM cost is input

    var unsystematicRisk = calculateUnsystematicRisk(returns, beta, marketReturns);
    var systematicRisk = beta * unsystematicRisk;
    var totalRisk = systematicRisk + unsystematicRisk;

    var riskComparison = '';
    if (totalRisk < oemCost) {
      riskComparison = 'Lower than OEM Cost';
    } else {
      riskComparison = 'Higher than OEM Cost';
    }

    // Display the calculated risks and recommendation
    sheet.getRange('I4').setValue(systematicRisk);
    sheet.getRange('I5').setValue(unsystematicRisk);
    sheet.getRange('I6').setValue(totalRisk);
    sheet.getRange('I7').setValue(riskComparison);
    
  } catch (error) {
    Logger.log('Error in calculateRisks: ' + error.message);
    SpreadsheetApp.getUi().alert('Error calculating risks: ' + error.message);
  }
}

function calculateUnsystematicRisk(returns, beta, marketReturns) {
  var unsystematicVariance = 0;

  for (var i = 0; i < returns.length; i++) {
    var expectedReturn = beta * marketReturns[i];
    var deviation = returns[i] - expectedReturn;
    unsystematicVariance += Math.pow(deviation, 2);
  }

  return Math.sqrt(unsystematicVariance / (returns.length - 1));
}

function average(array) {
  var sum = array.reduce(function(acc, val) {
    return acc + val;
  }, 0);
  return sum / array.length;
}

// New function to process OEM Cost data
function processOEMCostData(formData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Input');
  var oemCost = parseFloat(formData.oemCost);
  sheet.getRange('G2').setValue(oemCost);
}

