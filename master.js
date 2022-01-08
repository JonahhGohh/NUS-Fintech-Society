const XLSX = require("xlsx");

const workbook = XLSX.readFile("attendance-master.csv");

const sheet_name_list = workbook.SheetNames;
const jsonObj = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]], {
  raw: false,
});

const naturalLanguage = [];
const modelInterpretability = [];
const blockchain = [];
const web3 = [];

const workshops = [
  "Natural Language Processing in Fintech",
  "Model Interpretability in Risk Analytics",
  "Introduction to Blockchain",
  "Web2 to Web3 : The new Internet - Blockchain",
];

for (const row of jsonObj) {
  const signUpString =
    row[
      "Which Introductory Workshop(s) are you interested in registering for?"
    ];
  if (signUpString.includes(workshops[0])) {
    naturalLanguage.push({ ...row });
  }

  if (signUpString.includes(workshops[1])) {
    modelInterpretability.push({ ...row });
  }

  if (signUpString.includes(workshops[2])) {
    blockchain.push({ ...row });
  }

  if (signUpString.includes(workshops[3])) {
    web3.push({ ...row });
  }
}

const naturalLanguage_sheet = XLSX.utils.json_to_sheet(naturalLanguage);
const modelInterpretability_sheet = XLSX.utils.json_to_sheet(
  modelInterpretability
);
const blockchain_sheet = XLSX.utils.json_to_sheet(blockchain);
const web3_sheet = XLSX.utils.json_to_sheet(web3);
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

XLSX.writeFile(naturalLanguage_workbook, "naturalLanguage_attendance.csv");
XLSX.writeFile(
  modelInterpretability_workbook,
  "modelInterpretability_attendance.csv"
);
XLSX.writeFile(blockchain_workbook, "blockchain_attendance.csv");
XLSX.writeFile(web3_workbook, "web3_attendance.csv");
