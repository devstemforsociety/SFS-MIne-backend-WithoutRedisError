import { Request, RequestHandler, Response } from "express";
import { INVALID_SESSION_MSG } from "../../utils/constants";
import { db } from "../../db/connection";
import { z } from "zod";
import ExcelJS from "exceljs";

export const getAdminStudents: RequestHandler = async (
  req: Request,
  res: Response,
) => {
  try {
    const adminAuth = req.auth["ADMIN"];
    if (!adminAuth) {
      res.status(401).json({
        error: INVALID_SESSION_MSG,
      });
      return;
    }
    const students = await db.query.userTable.findMany({
      columns: {
        hash: false,
        salt: false,
        createdAt: false,
        updatedAt: false,
      },
    });
    res.json({ data: students });
  } catch (error) {
    console.log("🚀 ~ getAdminStudents ~ error:", error);
    res.status(500).json({
      error: "Server error in fetching student details",
    });
  }
};

export const getAdminStudent: RequestHandler = async (
  req: Request,
  res: Response,
) => {
  try {
    const adminAuth = req.auth["ADMIN"];
    if (!adminAuth) {
      res.status(401).json({
        error: INVALID_SESSION_MSG,
      });
      return;
    }
    const { studentId: studentIdUnsafe } = req.params;
    const studentId = z.string().uuid().safeParse(studentIdUnsafe);
    if (!studentId.success) {
      res.status(400).json({
        error: "Invalid student ID",
      });
      return;
    }
    const student = await db.query.userTable.findFirst({
      columns: {
        hash: false,
        salt: false,
        createdAt: false,
        updatedAt: false,
      },
      with: {
        enrolments: {
          with: {
            training: true,
          },
        },
      },
      where(fields, operators) {
        return operators.eq(fields.id, studentId.data);
      },
    });
    res.json({ data: student ?? {} });
  } catch (error) {
    console.log("🚀 ~ getAdminStudents ~ error:", error);
    res.status(500).json({
      error: "Server error in fetching student details",
    });
  }
};

export const exportAdminStudents: RequestHandler = async (
  req: Request,
  res: Response,
) => {
  try {
    const adminAuth = req.auth["ADMIN"];
    if (!adminAuth) {
      res.status(401).json({ error: INVALID_SESSION_MSG });
      return;
    }

    const students = await db.query.userTable.findMany({
      columns: {
        hash: false,
        salt: false,
        updatedAt: false,
      },
    });

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "SFS Admin";
    const sheet = workbook.addWorksheet("Students");

    sheet.columns = [
      { header: "First Name", key: "firstName", width: 18 },
      { header: "Last Name",  key: "lastName",  width: 18 },
      { header: "Email",      key: "email",      width: 32 },
      { header: "Mobile",     key: "mobile",     width: 16 },
      { header: "Joined On",  key: "createdAt",  width: 22 },
    ];

    const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1A56DB" } };
    headerRow.alignment = { vertical: "middle", horizontal: "center" };
    headerRow.height = 20;

    students.forEach((s) => {
      sheet.addRow({
        firstName: s.firstName ?? "",
        lastName:  s.lastName  ?? "",
        email:     s.email     ?? "",
        mobile:    s.mobile    ?? "",
        createdAt: s.createdAt ? new Date(s.createdAt).toLocaleDateString("en-IN") : "",
      });
    });

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        row.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: rowNumber % 2 === 0 ? "FFF3F4F6" : "FFFFFFFF" },
        };
      }
      row.border = { bottom: { style: "thin", color: { argb: "FFE5E7EB" } } };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="students-${Date.now()}.xlsx"`);
    res.send(buffer);
  } catch (error) {
    console.log("🚀 ~ exportAdminStudents ~ error:", error);
    res.status(500).json({ error: "Server error in exporting students" });
  }
};
