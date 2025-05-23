const express = require("express");
const multer = require("multer");
const cors = require("cors");
const fs = require("fs");
const xlsx = require("exceljs");
const csvParser = require("csv-parser");
const iconv = require("iconv-lite");
const path = require("path");

const uploadDir = "uploads";
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir);
}

const app = express();
app.use(cors());
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Multer storage configuration
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname);
  },
});

const upload = multer({ storage: storage });

// Process multiple CSV files and track attendance
app.post("/download", upload.array("files"), async (req, res) => {
  if (!req.files || req.files.length === 0) {
    return res.status(400).json({ message: "No files uploaded" });
  }

  const attendanceCount = {};

  for (const file of req.files) {
    const filepath = path.join(uploadDir, file.originalname);
    console.log(`Processing file: ${filepath}`);

    try {
      await processAttendanceFile(filepath, attendanceCount);
    } catch (error) {
      console.error("Error processing file:", error);
      return res
        .status(500)
        .json({ message: "Failed to process file", error: error.message });
    }
  }

  console.log("Final attendance count:", attendanceCount);

  const workbook = new xlsx.Workbook();
  const worksheet = workbook.addWorksheet("Attendance Summary");

  worksheet.columns = [
    { header: "Name", key: "name", width: 30 },
    { header: "Days Present", key: "daysPresent", width: 15 },
  ];

  Object.entries(attendanceCount).forEach(([name, count]) => {
    worksheet.addRow({ name, daysPresent: count });
  });

  res.setHeader(
    "Content-Disposition",
    'attachment; filename="attendance_summary.xlsx"'
  );
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );

  await workbook.xlsx.write(res);
  res.end();

  req.files.forEach((file) => {
    fs.unlink(file.path, (err) => {
      if (err) console.error("Error deleting file:", err);
    });
  });
});

async function processAttendanceFile(filepath, attendanceCount) {
  return new Promise((resolve, reject) => {
    const uniqueAttendeesInFile = new Set();
    let inParticipantsSection = false;
    let isActivitiesSection = false;
    let headerSkipped = false;

    fs.createReadStream(filepath)
      .pipe(iconv.decodeStream("utf16le"))
      .pipe(
        csvParser({
          trim: true,
          skip_empty_lines: true,
          separator: "\t",
          quote: '"',
          headers: false,
        })
      )
      .on("data", (row) => {
        console.log("Row Data:", row);

        if (!headerSkipped && row[0] && row[0].includes("Name") && row[6] && row[6].includes("Role")) {
          headerSkipped = true;
          inParticipantsSection = true;
          return;
        }

        if (inParticipantsSection) {
          let fullName = row[0] || ""; // Name is in the first column
          let role = row[6] || ""; // Role is in the 7th column
          fullName = fullName.trim();
          role = role.trim().toLowerCase();

          if (role === "presenter" || role === "organizer") {
            const name = fullName.split('(')[0].trim();
            if (name && !uniqueAttendeesInFile.has(name)) {
              uniqueAttendeesInFile.add(name);
              if (attendanceCount[name]) {
                attendanceCount[name]++;
              } else {
                attendanceCount[name] = 1;
              }
            }
          }
        }
      })
      .on("end", () => {
        console.log("Processed file:", filepath);
        resolve();
      })
      .on("error", (err) => {
        console.error("Error processing file:", err);
        reject(err);
      });
  });
}



// Server setup
const port = 5000;
app.listen(port, () => console.log(`Server running on port ${port}`));
