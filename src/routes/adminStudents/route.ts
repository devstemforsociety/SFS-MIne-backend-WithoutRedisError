import { Router } from "express";
import { exportAdminStudents, getAdminStudent, getAdminStudents } from "./controller";
import { requireAuthToken } from "../../middleware";

const adminStudentsRouter = Router();

adminStudentsRouter.get("/", requireAuthToken("ADMIN"), getAdminStudents);
adminStudentsRouter.get(
  "/:studentId",
  requireAuthToken("ADMIN"),
  getAdminStudent,
);

adminStudentsRouter.get("/export", requireAuthToken("ADMIN"), exportAdminStudents);

export default adminStudentsRouter;
