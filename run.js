const fs = require("fs");
const XLSX = require("xlsx");
const { Document, Packer, Paragraph, TextRun, PageBreak } = require("docx");

// Load Excel
const workbook = XLSX.readFile("students.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet);

// Config
const subjectCodeToFind = "101";
const absentRolls = ["204050", "202122","202129"]; // Update as needed
const groupSize = 200;

// Filter students with target subject in Excel order
const filtered = data.filter((row) =>
  Object.values(row).includes(parseInt(subjectCodeToFind))
);

const allRolls = filtered
  .filter((row) => row.roll !== undefined && row.roll !== null)
  .map((row) => row.roll.toString());

// Step: Build groups of 200 present students
const groups = [];
let presentCount = 0;
let i = 0;
let currentGroup = [];
let currentGroupAbsent = [];

while (i < allRolls.length) {
  const roll = allRolls[i];
  const isAbsent = absentRolls.includes(roll);

  currentGroup.push(roll);
  if (isAbsent) {
    currentGroupAbsent.push(roll);
  } else {
    presentCount++;
  }

  if (presentCount === groupSize) {
    groups.push({
      fullGroup: currentGroup.slice(), // keep original order including absent
      absents: currentGroupAbsent.slice(),
    });
    // Reset for next group
    currentGroup = [];
    currentGroupAbsent = [];
    presentCount = 0;
  }

  i++;
}

// Push last group if some left
if (currentGroup.length > 0) {
  groups.push({
    fullGroup: currentGroup,
    absents: currentGroupAbsent,
  });
}

// Compress roll ranges for each group (excluding absents)
const sections = [];

groups.forEach((group, index) => {
  const { fullGroup, absents } = group;
  const present = fullGroup.filter((r) => !absents.includes(r));

  const rollRanges = [];
  let i = 0;
  while (i < present.length) {
    const start = present[i];
    let end = start;
    let count = 1;

    while (
      i + 1 < present.length &&
      parseInt(present[i + 1]) === parseInt(present[i]) + 1
    ) {
      end = present[i + 1];
      count++;
      i++;
    }

    if (start === end) {
      rollRanges.push(${start});
    } else {
      rollRanges.push(${start}---${end}=${count});
    }

    i++;
  }

  const rollRangeText = rollRanges.join(", ");
  const absentText = absents.length ? absents.join(", ") : "0";

  const children = [
    new Paragraph({
      children: [
        new TextRun({
          text: Group ${index + 1},
          bold: true,
          size: 28,
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 200 },
      children: [new TextRun(Roll Range: ${rollRangeText})],
    }),
    new Paragraph({
      children: [new TextRun(Absent: ${absentText})],
    }),
  ];

  if (index < groups.length - 1) {
    children.push(new Paragraph({ children: [new PageBreak()] }));
  }

  sections.push({ children });
});

// Build DOCX
const doc = new Document({
  creator: "Top Sheet Generator",
  title: "Student Top Sheet",
  description: "200-present-per-group layout",
  sections,
});

// Save file
Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("TopSheet.docx", buffer);
  console.log("✅ DOCX file created: TopSheet.docx");
});
