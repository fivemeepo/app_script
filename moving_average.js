function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  // Define the range for the input columns
  var lastColumn = sheet.getLastColumn();
  var inputRange = sheet.getRange("A1:D");

  // Check if the edited range is within the input range
  if (
    range.getRow() <= inputRange.getLastRow() &&
    range.getColumn() <= inputRange.getLastColumn()
  ) {
    // Clear any existing charts
    var charts = sheet.getCharts();
    for (var i = 0; i < charts.length; i++) {
      sheet.removeChart(charts[i]);
    }

    // Get the data from the input range
    var dataRange = sheet.getRange(
      inputRange.getRow(),
      inputRange.getColumn(),
      inputRange.getLastRow() - inputRange.getRow() + 1,
      inputRange.getLastColumn() - inputRange.getColumn() + 1
    );

    dataRange.setNumberFormat("0.00,, \" mil\"");
    var firstColumnRange = sheet.getRange(
      inputRange.getRow(),
      inputRange.getColumn(),
      inputRange.getLastRow() - inputRange.getRow() + 1,
      1
    );
    firstColumnRange.setNumberFormat("yyyy/mm/dd");

    // Calculate column C using the specified formula
    var columnCFormula = "=ARRAYFORMULA(AVERAGE(OFFSET($B$2, ROW(B2)-ROW($B$2), 0, 7, 1)))";
    var columnCRange = sheet.getRange(
      inputRange.getRow()+1,
      inputRange.getColumn() + 2,
      inputRange.getLastRow() - inputRange.getRow() + 1,
      1
    );
    columnCRange.setFormula(columnCFormula);

    // Set the title for the first row of column C
    var columnCTitle = "Moving Average-7D";
    var columnCTitleCell = sheet.getRange(inputRange.getRow(), inputRange.getColumn() + 2);
    columnCTitleCell.setValue(columnCTitle);

    // Calculate column D using the specified formula
    var columnDFormula = "=ARRAYFORMULA(AVERAGE(OFFSET($B$2, ROW(B2)-ROW($B$2), 0, 30, 1)))";
    var columnDRange = sheet.getRange(
      inputRange.getRow()+1,
      inputRange.getColumn() + 3,
      inputRange.getLastRow() - inputRange.getRow() + 1,
      1
    );
    columnDRange.setFormula(columnDFormula);

    // Set the title for the first row of column D
    var columnDTitle = "Moving Average-30D";
    var columnDTitleCell = sheet.getRange(inputRange.getRow(), inputRange.getColumn() + 3);
    columnDTitleCell.setValue(columnDTitle);


    // Get the title for the chart from cell B1
    var chartTitle = sheet.getRange("B1").getValue();

    // Get the series titles from the first row
    var seriesTitles = sheet
      .getRange(1, inputRange.getColumn() + 1, 1, inputRange.getLastColumn() - inputRange.getColumn())
      .getValues()[0];

    // Create a new chart
    var chartBuilder = sheet.newChart();
    chartBuilder
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataRange)
      .setPosition(5, 7, 0, 0)
      .setOption("title", chartTitle)
      .setOption("vAxes", {
        0: {
          format: "0.00,, \"mil\"",
          viewWindow: {
            min: 13000000, // 14 million
            max: 22000000, // 23 million
          },
        },
      })
      .setOption("series", getSeriesOptions(seriesTitles));

    // Insert the chart into the sheet
    sheet.insertChart(chartBuilder.build());
  }
}

function getSeriesOptions(seriesTitles) {
  var seriesOptions = {};
  for (var i = 0; i < seriesTitles.length; i++) {
    seriesOptions[i] = {
      labelInLegend: seriesTitles[i],
    };
  }
  return seriesOptions;
}

