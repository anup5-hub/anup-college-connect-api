import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import ExcelJS from "exceljs";
import pg from "pg";

dotenv.config();

const { Pool } = pg;
const app = express();
app.use(cors());
app.use(express.json());

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: { rejectUnauthorized: false }
});

async function start() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS queries (
      id SERIAL PRIMARY KEY,
      query_number TEXT,
      student_name TEXT NOT NULL,
      email TEXT,
      category TEXT,
      question TEXT NOT NULL,
      status TEXT DEFAULT 'New',
      created_at TIMESTAMPTZ DEFAULT NOW()
    )
  `);

  app.get("/health", (req, res) => res.json({ ok: true }));

  app.post("/queries", async (req, res) => {
    if (req.header("x-api-key") !== process.env.API_KEY) {
      return res.status(401).json({ error: "Unauthorized" });
    }

    const { queryNumber, studentName, email, category, question, status = "New" } = req.body;

    const result = await pool.query(
      `INSERT INTO queries
       (query_number, student_name, email, category, question, status)
       VALUES ($1, $2, $3, $4, $5, $6)
       RETURNING *`,
      [queryNumber, studentName, email, category, question, status]
    );

    res.status(201).json(result.rows[0]);
  });

  app.get("/export", async (req, res) => {
    if (req.query.key !== process.env.API_KEY) {
      return res.status(401).send("Unauthorized");
    }

    const result = await pool.query("SELECT * FROM queries ORDER BY created_at DESC");
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Student Queries");

    sheet.columns = [
      { header: "Query Number", key: "query_number", width: 16 },
      { header: "Student Name", key: "student_name", width: 25 },
      { header: "Email", key: "email", width: 30 },
      { header: "Category", key: "category", width: 20 },
      { header: "Question", key: "question", width: 50 },
      { header: "Status", key: "status", width: 16 },
      { header: "Created Date", key: "created_at", width: 24 }
    ];

    sheet.addRows(result.rows);
    sheet.getRow(1).font = { bold: true };

    res.setHeader(
      "Content-Disposition",
      "attachment; filename=Anup-College-Queries.xlsx"
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    await workbook.xlsx.write(res);
    res.end();
  });

  app.listen(process.env.PORT || 3000, () => {
    console.log("Server running");
  });
}

start().catch(console.error);
