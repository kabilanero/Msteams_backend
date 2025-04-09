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

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname);
  },
});

const upload = multer({ storage: storage });

// POST route for uploading and processing files
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
      return res.status(500).json({
        message: "Failed to process file",
        error: error.message,
      });
    }
  }

  // Create Excel file
  const workbook = new xlsx.Workbook();
  const worksheet = workbook.addWorksheet("Attendance Summary");

  worksheet.columns = [
    { header: "Name", key: "name", width: 30 },
    { header: "Email", key: "email", width: 30 },
    { header: "Days Present", key: "daysPresent", width: 15 },
  ];

  Object.values(attendanceCount).forEach(({ name, email, count }) => {
    worksheet.addRow({
      name,
      email,
      daysPresent: count,
    });
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

  // Delete uploaded files
  req.files.forEach((file) => {
    fs.unlink(file.path, (err) => {
      if (err) console.error("Error deleting file:", err);
    });
  });
});

// Helper function to process each file
async function processAttendanceFile(filepath, attendanceCount) {
  return new Promise((resolve, reject) => {
    const uniqueAttendeesInFile = new Set();
    let inParticipantsSection = false;
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
        if (
          !headerSkipped &&
          row[0]?.includes("Name") &&
          row[6]?.includes("Role")
        ) {
          headerSkipped = true;
          inParticipantsSection = true;
          return;
        }

        if (inParticipantsSection) {
          const fullName = row[0]?.trim() || "";
          const email = row[4]?.trim().toLowerCase() || "";
          const role = row[6]?.trim().toLowerCase() || "";

          if (
            (role === "presenter" || role === "organizer") &&
            email &&
            !fullName.toLowerCase().includes("unverified")
          ) {
            const cleanName = fullName.replace(/\s*\(Guest\)/i, "").trim();
            const uniqueKey = email; // key based on email to prevent duplicates

            if (!uniqueAttendeesInFile.has(uniqueKey)) {
              uniqueAttendeesInFile.add(uniqueKey);

              if (attendanceCount[uniqueKey]) {
                attendanceCount[uniqueKey].count += 1;
              } else {
                attendanceCount[uniqueKey] = {
                  name: cleanName,
                  email,
                  count: 1,
                };
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

// Server start
const port = 5000;
app.listen(port, () => console.log(`Server running on port ${port}`));
