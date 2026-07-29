function processResponses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Demo_Data");
  const processedSheet = ss.getSheetByName("Processed_Data");

  const rawData = rawSheet.getDataRange().getValues();

  processedSheet.clearContents();

  processedSheet.getRange(1,1,1,10).setValues([[
    "Timestamp","Analyst Name","Scenario Type","Priority","Source",
    "Error Present","Requires Escalation","Recommended Queue",
    "Confidence Level","Review Flag"
  ]]);

  for (let i = 1; i < rawData.length; i++) {

    const row = rawData[i];

    const timestamp = row[0];
    const analyst = row[1];
    const scenario = row[2];
    const priority = row[3];
    const source = row[4];
    const errorFlag = row[5];
    const escalationFlag = row[6];

    let queue = "";
    let confidence = "";
    let review = "";

    if (escalationFlag === "Yes") {
      queue = "Escalation Queue";
      confidence = "High";
      review = "Immediate Review";
    } else if (priority === "High" && errorFlag === "Yes") {
      queue = "Priority Validation Queue";
      confidence = "High";
      review = "Review Required";
    } else if (priority === "High") {
      queue = "Priority Queue";
      confidence = "Medium";
      review = "Monitor";
    } else if (errorFlag === "Yes") {
      queue = "Quality Check Queue";
      confidence = "Medium";
      review = "Review Required";
    } else {
      queue = "Standard Queue";
      confidence = "High";
      review = "No Action";
    }

    processedSheet.appendRow([
      timestamp, analyst, scenario, priority, source,
      errorFlag, escalationFlag, queue, confidence, review
    ]);
  }
}
