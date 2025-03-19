const express = require("express");
const multer = require("multer");
const cors = require("cors");
const fs = require("fs");
const xlsx = require("exceljs");
const csvParser = require("csv-parser");
const iconv = require('iconv-lite');

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
    cb(null, "uploads");
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname);
  },
});

const upload = multer({ storage });

// Process multiple CSV files and track attendance
app.post("/download", upload.array("files"), async (req, res) => {
  if (!req.files || req.files.length === 0) {
    return res.status(400).json({ message: "No files uploaded" });
  }

  // Object to store attendance counts for each student
  const attendanceCount = {};

  // Process each uploaded file
  for (const file of req.files) {
    const filepath = file.path;
    console.log(`Processing file: ${file.originalname}`);
    
    // Process this CSV file
    await processAttendanceFile(filepath, attendanceCount);
  }

  console.log("Final attendance count:", attendanceCount);

  // Create XLSX with the compiled attendance data
  const workbook = new xlsx.Workbook();
  const worksheet = workbook.addWorksheet("Attendance Summary");
  
  // Define columns
  worksheet.columns = [
    { header: "Name", key: "name", width: 30 },
    { header: "Days Present", key: "daysPresent", width: 15 }
  ];

  // Add rows from the attendance data
  Object.entries(attendanceCount).forEach(([name, count]) => {
    worksheet.addRow({ name, daysPresent: count });
  });

  // Set response headers for file download
  res.setHeader(
    "Content-Disposition",
    'attachment; filename="attendance_summary.xlsx"'
  );
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );

  // Write the file to response
  await workbook.xlsx.write(res);
  res.end();

  // Clean up uploaded files
  req.files.forEach(file => {
    fs.unlink(file.path, (err) => {
      if (err) console.error("Error deleting file:", err);
    });
  });
});

async function processAttendanceFile(filepath, attendanceCount) {
    return new Promise((resolve, reject) => {
      // Set to track unique attendees in this file
      const uniqueAttendeesInFile = new Set();
  
      fs.createReadStream(filepath)
        .pipe(iconv.decodeStream('utf-8'))
        .pipe(csvParser({
          trim: false,
          quote: '"',
          ltrim: true,
          rtrim: true,
          skip_empty_lines: true
        }))
        .on("data", (row) => {
          let fullName = row["Name"] || "";
          fullName = fullName.replace(/^"|"$/g, '').trim();
          
          let role = row["Role"] ? row["Role"].trim().toLowerCase() : "";
  
          // Only count if attendee AND name not already processed in this file
          if (role === "attendee" && fullName && !uniqueAttendeesInFile.has(fullName)) {
            uniqueAttendeesInFile.add(fullName);
            
            // Increment the attendance count
            if (attendanceCount[fullName]) {
              attendanceCount[fullName]++;
            } else {
              attendanceCount[fullName] = 1;
            }
          }
        })
        .on("end", () => {
          resolve();
        });
    });
  }
  
// Server setup
const port = 5000;
app.listen(port, () => console.log(`Server running on port ${port}`));
