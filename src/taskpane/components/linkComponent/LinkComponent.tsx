
/* global Office, PowerPoint */
declare const Office: any;
declare const PowerPoint: any;

import React, { useEffect, useState } from "react";
import {
  Box, Typography, Button, Paper, MenuItem, TextField, CircularProgress, Alert,
} from "@mui/material";
import { Link as LinkIcon, AddBox } from "@mui/icons-material";
import { getDistinctWorkbooks, getLinksByWorkbook } from "../services/api";

interface ExcelComponent {
  id: string; name: string; sheetName: string; rangeAddress: string;
  type: string; excelFileId: string; excelFileName: string; dataSnapshot: string;
}

const truncateText = (text: string, maxLength: number): string => {
  if (!text) return "";
  return text.length > maxLength ? text.substring(0, maxLength - 3) + "..." : text;
};

const getRawBase64 = (input: any): string => {
  if (!input) return "";
  let str = String(input);
  if (str.indexOf("base64,") !== -1) {
    str = str.split("base64,")[1];
  }
  return str.replace(/[^A-Za-z0-9+/=]/g, "");
};

const LinkComponent: React.FC<{ onLinkSuccess: () => void }> = ({ onLinkSuccess }) => {
  const [excelFiles, setExcelFiles] = useState<{ id: string; name: string }[]>([]);
  const [selectedFileName, setSelectedFileName] = useState<string>("");
  const [excelComponents, setExcelComponents] = useState<ExcelComponent[]>([]);
  const [selectedComponentId, setSelectedComponentId] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(false);
  
  const [fetchingFiles, setFetchingFiles] = useState<boolean>(false);
  const [fetchingComponents, setFetchingComponents] = useState<boolean>(false);
  const [statusMessage, setStatusMessage] = useState<{ text: string; severity: "success" | "error" | "info" } | null>(null);

  useEffect(() => { fetchExcelDocuments(); }, []);

  const fetchExcelDocuments = async () => {
    setFetchingFiles(true);
    setStatusMessage(null);
    try {
      const res = await getDistinctWorkbooks();
      console.log("[DEBUG] API Response:", res);

      if (res.success && res.data) {
        if (res.data.length === 0) {
          setStatusMessage({ 
            text: "Database connected, but 0 Excel workbooks found in DB. Please register a link from Excel first.", 
            severity: "info" 
          });
        } else {
          setExcelFiles(res.data.map((f: any) => ({ id: f._id, name: f._id })));
        }
      } else {
        setStatusMessage({ text: "Backend succeeded, but query did not return records.", severity: "error" });
      }
    } catch (e: any) { 
      console.error("Fetch Files Error:", e);
      setStatusMessage({ 
        text: `API Connection Failed: ${e.message || "Failed to reach backend"}. Ensure Node.js server is running on port 5000.`, 
        severity: "error" 
      });
    } finally {
      setFetchingFiles(false);
    }
  };

  const handleExcelFileChange = async (fileName: string) => {
    setSelectedFileName(fileName);
    setExcelComponents([]);
    setSelectedComponentId("");
    setFetchingComponents(true);
    try {
      const res = await getLinksByWorkbook(fileName);
      if (res.success) setExcelComponents(res.data.map((item: any) => ({
        id: item.linkId,
        name: `${item.type} on [${item.sheetName}] (${item.rangeAddress})`,
        sheetName: item.sheetName, rangeAddress: item.rangeAddress, type: item.type,
        excelFileId: item.excelFileId, excelFileName: item.excelFileName, dataSnapshot: item.dataSnapshot,
      })));
    } catch (e) { 
      console.error("Fetch Components Error:", e); 
      setStatusMessage({ text: "Failed to load workbook components.", severity: "error" });
    } finally {
      setFetchingComponents(false);
    }
  };

  const handleInsertLinkToSlide = async () => {
    if (!selectedComponentId) return;
    setLoading(true);
    setStatusMessage(null);

    try {
      const comp = excelComponents.find((c) => c.id === selectedComponentId);
      if (!comp || !comp.dataSnapshot) throw new Error("Image data not found.");

      const rawBase64 = getRawBase64(comp.dataSnapshot);

      await PowerPoint.run(async (context: any) => {
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items");
        await context.sync();

        let shapeToDelete: any = null;
        
        const insertOptions: any = {
            coercionType: Office.CoercionType.Image
        };

        const slides = context.presentation.getSelectedSlides();
        slides.load("items");
        await context.sync();
        const currentSlide = slides.items[0];

        const shapes = currentSlide.shapes;
        shapes.load("items/id");
        await context.sync();
        const existingIds = new Set(shapes.items.map((s: any) => s.id));

        if (selectedShapes.items.length > 0) {
          const activeShape = selectedShapes.items[0];
          activeShape.load(["left", "top", "tags"]);
          await context.sync();

          insertOptions.imageLeft = activeShape.left;
          insertOptions.imageTop = activeShape.top;

          const isLinked = activeShape.tags.getItemOrNullObject("LIVE_LINK_ID");
          await context.sync();
          
          if (isLinked.isNullObject) {
            shapeToDelete = activeShape;
          }
        }

        await new Promise<void>((resolve, reject) => {
            Office.context.document.setSelectedDataAsync(
                rawBase64,
                insertOptions,
                (res: any) => {
                    if (res.status === Office.AsyncResultStatus.Failed) reject(new Error(res.error.message));
                    else resolve();
                }
            );
        });

        const updatedShapes = currentSlide.shapes;
        updatedShapes.load("items/id");
        await context.sync();

        const newImageShape = updatedShapes.items.find((s: any) => !existingIds.has(s.id));

        if (newImageShape) {
            newImageShape.tags.add("LIVE_LINK_ID", comp.id);
            newImageShape.tags.add("EXCEL_FILE_ID", comp.excelFileId);
            newImageShape.tags.add("EXCEL_FILE_NAME", comp.excelFileName || "");
            newImageShape.tags.add("EXCEL_SHEET_NAME", comp.sheetName);
            newImageShape.tags.add("EXCEL_RANGE_ADDRESS", comp.rangeAddress);
            newImageShape.tags.add("TYPE", comp.type);
            
            if (shapeToDelete) {
              shapeToDelete.delete();
            }
            
            await context.sync();
            setStatusMessage({ text: "Success! Image inserted.", severity: "success" });
        } else {
            throw new Error("Could not trace newly inserted image shape.");
        }

        onLinkSuccess();
      });
    } catch (err: any) {
      console.error("Error:", err);
      setStatusMessage({ text: err.message || "Failed to insert.", severity: "error" });
    } finally {
      setLoading(false);
    }
  };

  return (
    <Box sx={{ px: 0.5 }}>
      <Typography variant="caption" sx={{ fontWeight: 800, mb: 1, display: "block", fontFamily: "Segoe UI, Arial", fontSize: "11px" }}>
        LINK NEW COMPONENT
      </Typography>
      
      <Paper elevation={0} sx={{ p: 2, border: "1px solid #D1D1D1", borderRadius: "8px" }}>
        <Box sx={{ textAlign: 'center', mb: 2 }}>
          <LinkIcon sx={{ color: "#0078d4", fontSize: 28, mb: 0.5 }} />
          <Typography sx={{ fontWeight: 800, fontSize: "16px", color: "#323130", fontFamily: "Segoe UI, Arial" }}>
            Live Connection
          </Typography>
        </Box>

        <TextField 
          select 
          fullWidth 
          size="small" 
          label="1. Choose Excel Document" 
          value={selectedFileName} 
          disabled={loading || fetchingFiles || fetchingComponents}
          sx={{
            mb: 1.5,
            "& .MuiOutlinedInput-root": { height: "42px", fontSize: "13px", fontFamily: "Segoe UI, Arial" },
            "& .MuiInputLabel-root": { fontSize: "13px", fontFamily: "Segoe UI, Arial" }
          }} 
          onChange={(e) => handleExcelFileChange(e.target.value)}
        >
          {fetchingFiles ? (
            <MenuItem disabled sx={{ fontSize: "13px", fontFamily: "Segoe UI, Arial" }}>
              Loading files...
            </MenuItem>
          ) : excelFiles.length > 0 ? (
            excelFiles.map((f) => (
              <MenuItem 
                key={f.id} 
                value={f.id}
                sx={{ fontSize: "13px", fontFamily: "Segoe UI, Arial", textAlign: "left" }}
              >
                {truncateText(f.name, 25)}
              </MenuItem>
            ))
          ) : (
            <MenuItem disabled sx={{ fontSize: "13px", fontFamily: "Segoe UI, Arial" }}>
              No registered Excel files found.
            </MenuItem>
          )}
        </TextField>

        <TextField 
          select 
          fullWidth 
          size="small" 
          label="2. Choose Live Component" 
          value={selectedComponentId} 
          disabled={excelComponents.length === 0 || loading || fetchingComponents}
          sx={{
            mb: 2,
            "& .MuiOutlinedInput-root": { height: "42px", fontSize: "13px", fontFamily: "Segoe UI, Arial" },
            "& .MuiInputLabel-root": { fontSize: "13px", fontFamily: "Segoe UI, Arial" }
          }} 
          onChange={(e) => setSelectedComponentId(e.target.value)}
        >
          {fetchingComponents ? (
            <MenuItem disabled sx={{ fontSize: "13px", fontFamily: "Segoe UI, Arial" }}>
              Loading components...
            </MenuItem>
          ) : excelComponents.length > 0 ? (
            excelComponents.map((c) => (
              <MenuItem 
                key={c.id} 
                value={c.id}
                sx={{ fontSize: "13px", fontFamily: "Segoe UI, Arial", textAlign: "left" }}
              >
                {truncateText(c.name, 35)}
              </MenuItem>
            ))
          ) : (
            <MenuItem disabled sx={{ fontSize: "13px", fontFamily: "Segoe UI, Arial" }}>
              No active links found.
            </MenuItem>
          )}
        </TextField>

        <Button 
          variant="contained" 
          fullWidth 
          onClick={handleInsertLinkToSlide} 
          disabled={loading || !selectedComponentId} 
          startIcon={loading ? <CircularProgress size={18} color="inherit" /> : <AddBox sx={{ fontSize: 18 }} />}
          sx={{
            height: "40px",
            bgcolor: "#0078d4",
            fontWeight: 700,
            textTransform: "none",
            fontSize: "13px",
            boxShadow: "none",
            fontFamily: "Segoe UI, Arial",
            "&:hover": { bgcolor: "#005a9e", boxShadow: "none" },
          }}
        >
          {loading ? "Linking..." : "Insert Link to Slide"}
        </Button>

        {statusMessage && (
          <Alert severity={statusMessage.severity} sx={{ mt: 2, fontSize: '11px', fontFamily: "Segoe UI, Arial", textAlign: "left" }}>
            {statusMessage.text}
          </Alert>
        )}
      </Paper>
    </Box>
  );
};

export default LinkComponent;