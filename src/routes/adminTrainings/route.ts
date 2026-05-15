import { Router } from "express";
import { approveTraining, exportAdminTrainings, getTraining, getTrainings } from "./controller";
import { requireAuthToken } from "../../middleware";

const adminTrainingRouter = Router();

adminTrainingRouter.get("/", requireAuthToken("ADMIN"), getTrainings);
adminTrainingRouter.get("/:trainingId", requireAuthToken("ADMIN"), getTraining);
adminTrainingRouter.post(
  "/:trainingId/decision",
  requireAuthToken("ADMIN"),
  approveTraining,
);

adminTrainingRouter.get("/export", requireAuthToken("ADMIN"), exportAdminTrainings);

export default adminTrainingRouter;
