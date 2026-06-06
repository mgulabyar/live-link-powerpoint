declare const PowerPoint: any;
declare const Office: any;

import React, { useState, useEffect } from "react";
import { Box, Typography, Button, Paper, IconButton, List, CircularProgress, Alert, Snackbar } from "@mui/material";
import { Refresh, Layers, Description, DeleteOutline } from "@mui/icons-material";
import { getLinkDetails, getBulkLinkDetails, deleteLinkData } from "../services/api";
import { deletePPTShape, PPTLinkedItem } from "../utils/officeHelpers";

interface ActiveConnectionsProps {
  pptLinks: PPTLinkedItem[];
  onLinkSuccess: () => void | Promise<void>;
}

const truncateText = (text: string, maxLength: number): string => {
  if (!text) return "";
  return text.length > maxLength ? text.substring(0, maxLength - 3) + "..." : text;
};

const ActiveConnections: React.FC<ActiveConnectionsProps> = ({ pptLinks, onLinkSuccess }) => {
  const [refreshingId, setRefreshingId] = useState<string | null>(null);
  const [globalUpdating, setGlobalUpdating] = useState<boolean>(false);
  const [alertMessage, setAlertMessage] = useState<{
    text: string;
    severity: "success" | "error";
  } | null>(null);
  
 
  const [dbLinksMap, setDbLinksMap] = useState<Map<string, any>>(new Map());

  const safePptLinks = pptLinks || [];
  const hasLinks = safePptLinks.length > 0;

  useEffect(() => {
    checkLinksReachability();
  }, [pptLinks]);

  const checkLinksReachability = async () => {
    if (safePptLinks.length === 0) {
      setDbLinksMap(new Map());
      return;
    }
    try {
      const ids = safePptLinks.map((item) => item.id);
      const response = await getBulkLinkDetails(ids);
      if (response.success && response.data) {
        const linksMap = new Map<string, any>();
        response.data.forEach((dbLink: any) => {
          linksMap.set(dbLink.linkId, dbLink);
        });
        setDbLinksMap(linksMap); 
      }
    } catch (e) {
      console.error("Failed to check reachability:", e);
    }
  };

  const getRawBase64 = (base64String: string): string => {
    if (base64String && base64String.includes("base64,")) {
      return base64String.split("base64,")[1];
    }
    return base64String;
  };

  const handleRefreshLink = async (item: PPTLinkedItem) => {
    setRefreshingId(item.id);
    setAlertMessage(null);

    try {
      const response = await getLinkDetails(item.id);
      if (!response.success || !response.data) {
        throw new Error("Connection data not found in database.");
      }

      const linkData = response.data;
      const rawBase64 = getRawBase64(linkData.dataSnapshot);

      await PowerPoint.run(async (context: any) => {
        const slides = context.presentation.getSelectedSlides();
        slides.load("items");
        await context.sync();

        let shapeUpdated = false;
        for (const slide of slides.items) {
          const shapes = slide.shapes;
          shapes.load("items/id");
          await context.sync();

          const targetShape = shapes.items.find((s: any) => s.id === item.shapeId);
          if (targetShape) {
            targetShape.load(["left", "top"]);
            await context.sync();

            const targetShapeProperties = {
              left: targetShape.left,
              top: targetShape.top,
            };

            const existingIds = new Set(shapes.items.map((s: any) => s.id));

            targetShape.delete();
            await context.sync();

            await new Promise<void>((resolve, reject) => {
              Office.context.document.setSelectedDataAsync(
                rawBase64,
                { coercionType: Office.CoercionType.Image },
                (asyncResult: any) => {
                  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    reject(new Error(asyncResult.error.message));
                  } else {
                    resolve();
                  }
                }
              );
            });

            const updatedShapes = slide.shapes;
            updatedShapes.load("items/id");
            await context.sync();

            const newImage = updatedShapes.items.find((s: any) => !existingIds.has(s.id));
            if (newImage) {
              newImage.left = targetShapeProperties.left;
              newImage.top = targetShapeProperties.top;

              newImage.tags.add("LIVE_LINK_ID", item.id);
              newImage.tags.add("EXCEL_FILE_ID", item.excelFileId);
              newImage.tags.add("EXCEL_FILE_NAME", item.excelFileName);
              newImage.tags.add("EXCEL_SHEET_NAME", item.sheetName);
              newImage.tags.add("EXCEL_RANGE_ADDRESS", item.rangeAddress);
              newImage.tags.add("TYPE", item.type);

              await context.sync();
              shapeUpdated = true;
              break;
            } else {
              throw new Error("Failed to detect refreshed shape.");
            }
          }
        }
        if (!shapeUpdated) throw new Error("Linked shape not found on current slide.");
      });

      setAlertMessage({ text: "Updated successfully!", severity: "success" });
      onLinkSuccess();
    } catch (err: any) {
      console.error(err);
      setAlertMessage({ text: err.message || "Failed to update connection.", severity: "error" });
    } finally {
      setRefreshingId(null);
    }
  };

  const handleUpdateAllLinks = async () => {
    if (safePptLinks.length === 0) return;
    setGlobalUpdating(true);
    setAlertMessage(null);

    try {
      const linkIds = safePptLinks.map((item) => item.id);

      const response = await getBulkLinkDetails(linkIds);
      if (!response.success || !response.data) {
        throw new Error("Failed to retrieve bulk updates.");
      }

      const dbLinksMapLocal = new Map<string, any>();
      response.data.forEach((dbLink: any) => {
        dbLinksMapLocal.set(dbLink.linkId, dbLink);
      });

      await PowerPoint.run(async (context: any) => {
        const slides = context.presentation.getSelectedSlides();
        slides.load("items");
        await context.sync();

        let updatedCount = 0;

        for (const slide of slides.items) {
          const shapes = slide.shapes;
          shapes.load("items/id");
          await context.sync();

          for (const item of safePptLinks) {
            const dbLink = dbLinksMapLocal.get(item.id);
            if (!dbLink) continue;

            const targetShape = shapes.items.find((s: any) => s.id === item.shapeId);
            if (targetShape) {
              targetShape.load(["left", "top"]);
              await context.sync();

              const targetShapeProperties = {
                left: targetShape.left,
                top: targetShape.top,
              };

              const existingIds = new Set(shapes.items.map((s: any) => s.id));

              targetShape.delete();
              await context.sync();

              const rawBase64 = getRawBase64(dbLink.dataSnapshot);

              await new Promise<void>((resolve, reject) => {
                Office.context.document.setSelectedDataAsync(
                  rawBase64,
                  { coercionType: Office.CoercionType.Image },
                  (asyncResult: any) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                      reject(new Error(asyncResult.error.message));
                    } else {
                      resolve();
                    }
                  }
                );
              });

              const updatedShapes = slide.shapes;
              updatedShapes.load("items/id");
              await context.sync();

              const newImage = updatedShapes.items.find((s: any) => !existingIds.has(s.id));
              if (newImage) {
                newImage.left = targetShapeProperties.left;
                newImage.top = targetShapeProperties.top;

                newImage.tags.add("LIVE_LINK_ID", item.id);
                newImage.tags.add("EXCEL_FILE_ID", item.excelFileId);
                newImage.tags.add("EXCEL_FILE_NAME", item.excelFileName);
                newImage.tags.add("EXCEL_SHEET_NAME", item.sheetName);
                newImage.tags.add("EXCEL_RANGE_ADDRESS", item.rangeAddress);
                newImage.tags.add("TYPE", item.type);

                updatedCount++;
                await context.sync();
              }
            }
          }
        }

        await context.sync();
        setAlertMessage({
          text: `Successfully updated ${updatedCount} connections.`,
          severity: "success",
        });
      });

      onLinkSuccess();
    } catch (err: any) {
      console.error("Bulk update error:", err);
      setAlertMessage({
        text: err.message || "Failed to update bulk connections.",
        severity: "error",
      });
    } finally {
      setGlobalUpdating(false);
    }
  };

  const handleBreakPPTLink = async (shapeId: string, linkId: string) => {
    try {
      await deletePPTShape(shapeId);
      await deleteLinkData(linkId);
      onLinkSuccess();
    } catch (e) {
      console.error("Error breaking PPT link:", e);
    }
  };

  return (
    <Box sx={{ px: 0.5 }}>
      {hasLinks ? (
        <Box
          sx={{ display: "flex", alignItems: "center", justifyContent: "space-between", mb: 1.5 }}
        >
          <Box sx={{ display: "flex", alignItems: "center", gap: 0.8 }}>
            <Layers sx={{ fontSize: 16, color: "#605E5C" }} />
            <Typography
              variant="caption"
              sx={{
                fontWeight: 800,
                color: "#605E5C",
                fontFamily: "Segoe UI, Arial",
                fontSize: "11px",
              }}
            >
              ACTIVE CONNECTIONS ({safePptLinks.length})
            </Typography>
          </Box>
          <Button
            size="small"
            disabled={globalUpdating}
            startIcon={
              globalUpdating ? (
                <CircularProgress size={10} color="inherit" />
              ) : (
                <Refresh sx={{ fontSize: 12 }} />
              )
            }
            onClick={handleUpdateAllLinks}
            sx={{
              fontSize: "11px",
              fontWeight: 700,
              p: 0,
              color: "#0078d4",
              textTransform: "none",
              fontFamily: "Segoe UI, Arial",
            }}
          >
            Update All
          </Button>
        </Box>
      ) : (
        <Box
          sx={{
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            justifyContent: "center",
            py: 2,
            textAlign: "center",
          }}
        >
          <Layers sx={{ fontSize: 24, color: "#A19F9D", mb: 0.8 }} />
          <Typography
            variant="caption"
            sx={{
              fontSize: "11px",
              fontWeight: 800,
              color: "#605E5C",
              fontFamily: "Segoe UI, Arial",
              mb: 0.4,
            }}
          >
            ACTIVE CONNECTIONS (0)
          </Typography>
          <Typography
            sx={{
              fontSize: "11px",
              color: "#605E5C",
              fontFamily: "Segoe UI, Arial",
            }}
          >
            No active slide links detected.
          </Typography>
        </Box>
      )}

      {hasLinks && (
        <List sx={{ p: 0 }}>
          {safePptLinks.map((item) => {
            // Retrieve matched DB record dynamically to get the custom name [1]
            const dbItem = dbLinksMap.get(item.id);
            const isReachable = !!dbItem;
            const customComponentName = dbItem?.componentName || ""; 

            return (
              <Paper
                key={item.id}
                elevation={0}
                sx={{
                  p: 1.5,
                  mb: 1.2,
                  border: "1px solid #EDEBE9",
                  borderRadius: "6px",
                  bgcolor: "#FAFAFA",
                  "&:hover": { borderColor: "#0078d4", bgcolor: "#FFFFFF" },
                  transition: "0.2s",
                }}
              >
                <Box
                  sx={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "flex-start",
                    mb: 1,
                  }}
                >
                  <Box sx={{ display: "flex", gap: 1 }}>
                    <Description sx={{ fontSize: 18, color: "#0078d4", mt: 0.1 }} />
                    <Box>
                      {/* PRIORITIZED: Display the custom componentName on card header, fallback to technical name [1] */}
                      <Typography
                        sx={{
                          fontSize: "12px",
                          fontWeight: 700,
                          color: "#323130",
                          lineHeight: 1.2,
                          fontFamily: "Segoe UI, Arial",
                        }}
                      >
                        {customComponentName || `${item.type} Connection`}
                      </Typography>

                      {/* UNIFIED SINGLE-LINE METADATA FLOW WRAPPED IN A DENSE FLEXBOX ROW [5] */}
                      <Box
                        sx={{
                          display: "flex",
                          gap: 0.8,
                          mt: 0.6,
                          flexWrap: "wrap",
                          alignItems: "center",
                        }}
                      >
                        <Typography
                          sx={{
                            fontSize: "10.5px",
                            color: "#605E5C",
                            fontFamily: "Segoe UI, Arial",
                          }}
                        >
                          File: {truncateText(item.excelFileName, 12)}
                        </Typography>
                        <Typography
                          sx={{
                            fontSize: "10.5px",
                            color: "#605E5C",
                            fontFamily: "Segoe UI, Arial",
                          }}
                        >
                          Sheet: {item.sheetName}
                        </Typography>
                        <Box sx={{ bgcolor: "#E1DFDD", px: 0.8, py: 0.1, borderRadius: "3px" }}>
                          <Typography
                            sx={{
                              fontSize: "10.5px",
                              fontWeight: 700,
                              color: "#323130",
                              fontFamily: "Segoe UI, Arial",
                            }}
                          >
                            Range: {item.rangeAddress}
                          </Typography>
                        </Box>
                      </Box>
                    </Box>
                  </Box>
                  <IconButton
                    size="small"
                    color="error"
                    disabled={refreshingId === item.id || globalUpdating}
                    onClick={() => handleBreakPPTLink(item.shapeId, item.id)}
                    sx={{ p: 0.5 }}
                  >
                    <DeleteOutline sx={{ fontSize: 16 }} />
                  </IconButton>
                </Box>

                <Button
                  variant="outlined"
                  fullWidth
                  size="small"
                  disabled={refreshingId === item.id || globalUpdating || !isReachable}
                  startIcon={
                    refreshingId === item.id ? (
                      <CircularProgress size={14} color="inherit" />
                    ) : (
                      <Refresh sx={{ fontSize: 12 }} />
                    )
                  }
                  onClick={() => handleRefreshLink(item)}
                  sx={{
                    height: "28px",
                    fontSize: "11px",
                    textTransform: "none",
                    fontWeight: 700,
                    borderColor: "#D1D1D1",
                    color: "#323130",
                    bgcolor: "#FFFFFF",
                    fontFamily: "Segoe UI, Arial",
                    "&:hover": { borderColor: "#0078d4", color: "#0078d4" },
                  }}
                >
                  {refreshingId === item.id ? "Updating..." : "Update"}
                </Button>
              </Paper>
            );
          })}
        </List>
      )}

      {/* DYNAMIC AUTO-DISSOLVING 2-SECOND BOTTOM TOAST [1] */}
      <Snackbar
        open={alertMessage !== null}
        autoHideDuration={2000} 
        onClose={() => setAlertMessage(null)}
        anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
      >
        {alertMessage ? (
          <Alert
            onClose={() => setAlertMessage(null)}
            severity={alertMessage.severity}
            sx={{ width: "100%", fontSize: "13px", fontFamily: "Segoe UI, Arial" }}
          >
            {alertMessage.text}
          </Alert>
        ) : undefined}
      </Snackbar>
    </Box>
  );
};

export default ActiveConnections;