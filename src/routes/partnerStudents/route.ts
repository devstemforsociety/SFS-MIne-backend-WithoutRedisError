import { Router } from "express";
import { requireAuthToken } from "../../middleware";
import { exportStudents, exportTrainingStudents, getStudents } from "./controller";

const partnerStudentsRouter = Router();

partnerStudentsRouter.get("/", requireAuthToken("PARTNER"), getStudents);
partnerStudentsRouter.get("/export", requireAuthToken("PARTNER"), exportStudents);
partnerStudentsRouter.get("/:trainingId/students/export", requireAuthToken("PARTNER"), exportTrainingStudents);

export default partnerStudentsRouter;
