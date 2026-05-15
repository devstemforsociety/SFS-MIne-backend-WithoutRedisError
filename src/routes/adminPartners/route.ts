import { Router } from "express";
import { requireAuthToken } from "../../middleware";
import {
  approvePartner,
  exportAdminPartners,
  getAdminPartner,
  getAdminPartners,
} from "./controller";

const adminPartnersRouter = Router();

adminPartnersRouter.get("/", requireAuthToken("ADMIN"), getAdminPartners);
adminPartnersRouter.get(
  "/:partnerId",
  requireAuthToken("ADMIN"),
  getAdminPartner,
);
adminPartnersRouter.post(
  "/:partnerId/decision",
  requireAuthToken("ADMIN"),
  approvePartner,
);
adminPartnersRouter.get(
  "/:partnerId/accounts/verify",
  requireAuthToken("ADMIN"),
);

adminPartnersRouter.get("/export", requireAuthToken("ADMIN"), exportAdminPartners);

export default adminPartnersRouter;
