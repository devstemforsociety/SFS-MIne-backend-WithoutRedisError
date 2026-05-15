import { Request, RequestHandler, Response } from "express";
import { db } from "../../db/connection";
import { INVALID_SESSION_MSG } from "../../utils/constants";
import {
  trainingEnrolmentTable,
  trainingTable,
  userTable,
} from "../../db/schema";
import { eq, inArray, sql } from "drizzle-orm";
import ExcelJS from "exceljs";

export const getStudents: RequestHandler = async (
  req: Request,
  res: Response,
) => {
  try {
    const partnerAuth = req.auth?.["PARTNER"];
    if (!partnerAuth) {
      res.status(401).json({
        error: INVALID_SESSION_MSG,
      });
      return;
    }
    const trainingIds = await db
      .select({ trainingId: trainingTable.id })
      .from(trainingTable)
      .where(eq(trainingTable.createdBy, partnerAuth.id));
    if (trainingIds.length === 0) {
      res.json({ data: [] });
      return;
    }
    const students = await db
      .select({
        firstName: userTable.firstName,
        lastName: userTable.lastName,
        email: userTable.email,
        id: userTable.id,
        mobile: userTable.mobile,
      })
      .from(trainingEnrolmentTable)
      .where(
        inArray(
          trainingEnrolmentTable.trainingId,
          trainingIds?.length !== 0
            ? trainingIds.map((tra) => tra.trainingId)
            : sql`false`,
        ),
      )
      .leftJoin(userTable, eq(trainingEnrolmentTable.userId, userTable.id))
      .groupBy(
        trainingEnrolmentTable.userId,
        userTable.firstName,
        userTable.lastName,
        userTable.email,
        userTable.id,
        userTable.mobile,
      );
      console.log("🚀 ~ getStudents ~ students:", students);
    res.json({ data: students });
  } catch (error) {
    console.log("🚀 ~ getStudents ~ error:", error);
    res.status(500).json({
      error: "Server error in fetching student details",
    });
  }
};

const buildStudentWorkbook = async (students: { firstName: string | null; lastName: string | null; email: string | null; mobile: string | null; }[]) => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "SFS Partner";
  const sheet = workbook.addWorksheet("Students");

  sheet.columns = [
    { header: "First Name", key: "firstName", width: 18 },
    { header: "Last Name",  key: "lastName",  width: 18 },
    { header: "Email",      key: "email",     width: 32 },
    { header: "Mobile",     key: "mobile",    width: 16 },
  ];

  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF059669" } };
  headerRow.alignment = { vertical: "middle", horizontal: "center" };
  headerRow.height = 20;

  students.forEach((s) => {
    sheet.addRow({
      firstName: s.firstName ?? "",
      lastName:  s.lastName  ?? "",
      email:     s.email     ?? "",
      mobile:    s.mobile    ?? "",
    });
  });

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      row.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: rowNumber % 2 === 0 ? "FFF0FDF4" : "FFFFFFFF" },
      };
    }
    row.border = { bottom: { style: "thin", color: { argb: "FFD1FAE5" } } };
  });

  return workbook;
};

export const exportStudents: RequestHandler = async (
  req: Request,
  res: Response,
) => {
  try {
    const partnerAuth = req.auth?.["PARTNER"];
    if (!partnerAuth) {
      res.status(401).json({ error: INVALID_SESSION_MSG });
      return;
    }

    const trainingIds = await db
      .select({ trainingId: trainingTable.id })
      .from(trainingTable)
      .where(eq(trainingTable.createdBy, partnerAuth.id));

    if (trainingIds.length === 0) {
      // Return empty workbook
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet("Students").columns = [
        { header: "First Name", key: "firstName", width: 18 },
        { header: "Last Name",  key: "lastName",  width: 18 },
        { header: "Email",      key: "email",     width: 32 },
        { header: "Mobile",     key: "mobile",    width: 16 },
      ];
      const buf = await wb.xlsx.writeBuffer();
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="students-${Date.now()}.xlsx"`);
      res.send(buf);
      return;
    }

    const students = await db
      .select({
        firstName: userTable.firstName,
        lastName:  userTable.lastName,
        email:     userTable.email,
        mobile:    userTable.mobile,
      })
      .from(trainingEnrolmentTable)
      .where(inArray(trainingEnrolmentTable.trainingId, trainingIds.map((t) => t.trainingId)))
      .leftJoin(userTable, eq(trainingEnrolmentTable.userId, userTable.id))
      .groupBy(
        trainingEnrolmentTable.userId,
        userTable.firstName,
        userTable.lastName,
        userTable.email,
        userTable.mobile,
      );

    const workbook = await buildStudentWorkbook(students as any);
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="students-${Date.now()}.xlsx"`);
    res.send(buffer);
  } catch (error) {
    console.log("🚀 ~ exportStudents ~ error:", error);
    res.status(500).json({ error: "Server error in exporting students" });
  }
};

export const exportTrainingStudents: RequestHandler = async (
  req: Request,
  res: Response,
) => {
  try {
    const partnerAuth = req.auth?.["PARTNER"];
    if (!partnerAuth) {
      res.status(401).json({ error: INVALID_SESSION_MSG });
      return;
    }

    const { z } = await import("zod");
    const { trainingId: trainingIdUnsafe } = req.params;
    const trainingId = z.string().uuid().safeParse(trainingIdUnsafe);
    if (!trainingId.success) {
      res.status(400).json({ error: "Invalid training ID" });
      return;
    }

    // Ensure this training belongs to the partner
    const training = await db.query.trainingTable.findFirst({
      columns: { id: true, title: true },
      with: {
        enrolments: {
          columns: {
            id: true,
            paidOn: true,
            completedOn: true,
          },
          with: {
            user: {
              columns: {
                firstName: true,
                lastName: true,
                email: true,
                mobile: true,
              },
            },
          },
        },
      },
      where(fields, operators) {
        return operators.and(
          operators.eq(fields.createdBy, partnerAuth.id),
          operators.eq(fields.id, trainingId.data),
        );
      },
    });

    if (!training) {
      res.status(404).json({ error: "Training not found" });
      return;
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "SFS Partner";
    const sheet = workbook.addWorksheet("Enrolled Students");

    sheet.columns = [
      { header: "First Name",    key: "firstName",   width: 18 },
      { header: "Last Name",     key: "lastName",    width: 18 },
      { header: "Email",         key: "email",       width: 32 },
      { header: "Mobile",        key: "mobile",      width: 16 },
      { header: "Paid On",       key: "paidOn",      width: 18 },
      { header: "Completed On",  key: "completedOn", width: 18 },
    ];

    const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF059669" } };
    headerRow.alignment = { vertical: "middle", horizontal: "center" };
    headerRow.height = 20;

    training.enrolments.forEach((enr) => {
      sheet.addRow({
        firstName:   enr.user?.firstName  ?? "",
        lastName:    enr.user?.lastName   ?? "",
        email:       enr.user?.email      ?? "",
        mobile:      enr.user?.mobile     ?? "",
        paidOn:      enr.paidOn      ? new Date(enr.paidOn).toLocaleDateString("en-IN")      : "Pending",
        completedOn: enr.completedOn ? new Date(enr.completedOn).toLocaleDateString("en-IN") : "In Progress",
      });
    });

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        row.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: rowNumber % 2 === 0 ? "FFF0FDF4" : "FFFFFFFF" },
        };
      }
      row.border = { bottom: { style: "thin", color: { argb: "FFD1FAE5" } } };
    });

    const safeTitle = (training.title ?? "training").replace(/[^a-z0-9]/gi, "-").toLowerCase();
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${safeTitle}-students-${Date.now()}.xlsx"`);
    res.send(buffer);
  } catch (error) {
    console.log("🚀 ~ exportTrainingStudents ~ error:", error);
    res.status(500).json({ error: "Server error in exporting training students" });
  }
};
