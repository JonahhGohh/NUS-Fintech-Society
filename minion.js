const XLSX = require("xlsx");

const workbook = XLSX.readFile("blockchain_attendee_report.csv");
const workbook1 = XLSX.readFile("blockchain_attendance.csv");
const workbook2 = XLSX.readFile("modelInterpretability_attendee_report.csv");
const workbook3 = XLSX.readFile("modelInterpretability_attendance.csv");
const workbook4 = XLSX.readFile("naturalLanguage_attendee_report.csv");
const workbook5 = XLSX.readFile("naturalLanguage_attendance.csv");
const workbook6 = XLSX.readFile("web3_attendee_report.csv");
const workbook7 = XLSX.readFile("web3_attendance.csv");

const sheet_name_list = workbook.SheetNames;
const sheet_name_list1 = workbook1.SheetNames;
const sheet_name_list2 = workbook2.SheetNames;
const sheet_name_list3 = workbook3.SheetNames;
const sheet_name_list4 = workbook4.SheetNames;
const sheet_name_list5 = workbook5.SheetNames;
const sheet_name_list6 = workbook6.SheetNames;
const sheet_name_list7 = workbook7.SheetNames;
const blockchain_attendees_report = XLSX.utils.sheet_to_json(
  workbook.Sheets[sheet_name_list[0]],
  {
    raw: false,
  }
);
const blockchain_final_report = XLSX.utils.sheet_to_json(
  workbook1.Sheets[sheet_name_list1[0]],
  {
    raw: false,
  }
);
const modelInterpretability_attendees_report = XLSX.utils.sheet_to_json(
  workbook2.Sheets[sheet_name_list2[0]],
  {
    raw: false,
  }
);
const modelInterpretability_final_report = XLSX.utils.sheet_to_json(
  workbook3.Sheets[sheet_name_list3[0]],
  {
    raw: false,
  }
);
const naturalLanguage_attendees_report = XLSX.utils.sheet_to_json(
  workbook4.Sheets[sheet_name_list4[0]],
  {
    raw: false,
  }
);
const naturalLanguage_final_report = XLSX.utils.sheet_to_json(
  workbook5.Sheets[sheet_name_list5[0]],
  {
    raw: false,
  }
);
const web3_attendees_report = XLSX.utils.sheet_to_json(
  workbook6.Sheets[sheet_name_list6[0]],
  {
    raw: false,
  }
);
const web3_final_report = XLSX.utils.sheet_to_json(
  workbook7.Sheets[sheet_name_list7[0]],
  {
    raw: false,
  }
);

for (const row of blockchain_final_report) {
  const email = row["Email"];
  row["Attendance"] = "No";
  for (const row1 of blockchain_attendees_report) {
    if (row1["Email"] == email) {
      row["Attendance"] = "Yes";
      break;
    }
  }
}

for (const row of modelInterpretability_final_report) {
  const email = row["Email"];
  row["Attendance"] = "No";
  for (const row1 of modelInterpretability_attendees_report) {
    if (row1["Email"] == email) {
      row["Attendance"] = "Yes";
      break;
    }
  }
}

for (const row of naturalLanguage_final_report) {
  const email = row["Email"];
  row["Attendance"] = "No";
  for (const row1 of naturalLanguage_attendees_report) {
    if (row1["Email"] == email) {
      row["Attendance"] = "Yes";
      break;
    }
  }
}

for (const row of web3_final_report) {
  const email = row["Email"];
  row["Attendance"] = "No";
  for (const row1 of web3_attendees_report) {
    if (row1["Email"] == email) {
      row["Attendance"] = "Yes";
      break;
    }
  }
}

const naturalLanguage_sheet = XLSX.utils.json_to_sheet(
  naturalLanguage_final_report
);
const modelInterpretability_sheet = XLSX.utils.json_to_sheet(
  modelInterpretability_final_report
);
const blockchain_sheet = XLSX.utils.json_to_sheet(naturalLanguage_final_report);
const web3_sheet = XLSX.utils.json_to_sheet(web3_final_report);
/* Add the worksheet to the workbook */
const naturalLanguage_workbook = XLSX.utils.book_new();
const modelInterpretability_workbook = XLSX.utils.book_new();
const blockchain_workbook = XLSX.utils.book_new();
const web3_workbook = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(
  naturalLanguage_workbook,
  naturalLanguage_sheet,
  "attendance"
);

XLSX.utils.book_append_sheet(
  modelInterpretability_workbook,
  modelInterpretability_sheet,
  "attendance"
);

XLSX.utils.book_append_sheet(
  blockchain_workbook,
  blockchain_sheet,
  "attendance"
);

XLSX.utils.book_append_sheet(web3_workbook, web3_sheet, "attendance");

XLSX.writeFile(
  naturalLanguage_workbook,
  "naturalLanguage_final_attendance.csv"
);
XLSX.writeFile(
  modelInterpretability_workbook,
  "modelInterpretability_final_attendance.csv"
);
XLSX.writeFile(blockchain_workbook, "blockchain_final_attendance.csv");
XLSX.writeFile(web3_workbook, "web3_final_attendance.csv");
