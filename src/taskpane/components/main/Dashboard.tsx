// declare const Office: any;

// import React, { useEffect, useState } from "react";
// import { Box, Typography, IconButton } from "@mui/material";
// import { Logout } from "@mui/icons-material";
// import LinkComponent from "../linkComponent/LinkComponent"; 
// import ActiveConnections from "../activeConnections/ActiveConnections"; 
// import { getPPTLinkedItems, PPTLinkedItem } from "../utils/officeHelpers"; 

// interface DashboardProps {
//   onLogout: () => void;
// }

// const LinkComponentSafe = LinkComponent as any;
// const ActiveConnectionsSafe = ActiveConnections as any;

// const Dashboard: React.FC<DashboardProps> = ({ onLogout }) => {
//   const [pptLinks, setPptLinks] = useState<PPTLinkedItem[]>([]);
//   const [resolvingFile, setResolvingFile] = useState<boolean>(true);

//   useEffect(() => {
//     loadPPTLinkedItems();
//   }, []);

//   const loadPPTLinkedItems = async () => {
//     setResolvingFile(true);
//     try {
//       const items = await getPPTLinkedItems();
//       setPptLinks(items);
//     } catch (e) {
//       console.error("Failed to load PPT linked items:", e);
//     } finally {
//       setResolvingFile(false);
//     }
//   };

//   return (
//     <Box
//       sx={{
//         height: "100vh",
//         display: "flex",
//         flexDirection: "column",
//         bgcolor: "#FFFFFF",
//         fontFamily: "Segoe UI, Arial, sans-serif",
//         overflow: "hidden",
//       }}
//     >
//       <Box
//         sx={{
//           position: "sticky",
//           top: 0,
//           zIndex: 100,
//           p: 2,
//           bgcolor: "#0078d4",
//           color: "#FFFFFF",
//           display: "flex",
//           alignItems: "center",
//           justifyContent: "space-between",
//         }}
//       >
//         <Box sx={{ width: 32 }} />
//         <Typography
//           sx={{
//             fontWeight: 700,
//             fontSize: "16px",
//             letterSpacing: "0.5px",
//             fontFamily: "Segoe UI, Arial",
//             textAlign: "center",
//           }}
//         >
//           EXCEL TO POWERPOINT
//         </Typography>
//         <IconButton onClick={onLogout} size="small" sx={{ color: "#FFFFFF" }}>
//           <Logout sx={{ fontSize: 18 }} />
//         </IconButton>
//       </Box>

//       <Box
//         sx={{
//           px: 2.5,
//           py: 1.5,
//           bgcolor: "#F3F2F1",
//           display: "flex",
//           alignItems: "center",
//           gap: 1,
//           borderBottom: "1px solid #EDEBE9",
//         }}
//       >
//         <Box sx={{ width: 8, height: 8, bgcolor: resolvingFile ? "#ffc83b" : "#107C10", borderRadius: "50%" }} />
//         <Typography sx={{ fontSize: "11px", fontWeight: 700, color: "#323130", fontFamily: "Segoe UI, Arial" }}>
//           {resolvingFile ? "SCANNING SLIDE CONTEXT..." : "SYNC ENGINE ACTIVE"}
//         </Typography>
//       </Box>

//       <Box
//         sx={{
//           p: 2,
//           flex: 1,
//           overflowY: "auto",
//           display: "flex",
//           flexDirection: "column",
//           gap: 2,
//           maxHeight: "95vh",
//           "&::-webkit-scrollbar": { display: "none" },
//           msOverflowStyle: "none",
//           scrollbarWidth: "none",
//         }}
//       >
//         <LinkComponentSafe onLinkSuccess={loadPPTLinkedItems} />

//         <ActiveConnectionsSafe pptLinks={pptLinks} onLinkSuccess={loadPPTLinkedItems} />
//       </Box>

//       <Box sx={{ p: 1.5, textAlign: "center", borderTop: "1px solid #EDEBE9" }}>
//         <Typography
//           sx={{ fontSize: "10px", color: "#A19F9D", fontWeight: 600, fontFamily: "Segoe UI, Arial" }}
//         >
//           Version 1.0.0 Stable
//         </Typography>
//       </Box>
//     </Box>
//   );
// };

// export default Dashboard;

declare const Office: any;

import React, { useEffect, useState } from "react";
import { Box, Typography, IconButton, CircularProgress } from "@mui/material";
import { Refresh } from "@mui/icons-material";
import LinkComponent from "../linkComponent/LinkComponent"; 
import ActiveConnections from "../activeConnections/ActiveConnections"; 
import { getPPTLinkedItems, PPTLinkedItem } from "../utils/officeHelpers"; // Exact relative path

interface DashboardProps {
  onLogout: () => void;
}

const LinkComponentSafe = LinkComponent as any;
const ActiveConnectionsSafe = ActiveConnections as any;

const Dashboard: React.FC<DashboardProps> = () => {
  const [pptLinks, setPptLinks] = useState<PPTLinkedItem[]>([]);
  const [resolvingFile, setResolvingFile] = useState<boolean>(true);

  useEffect(() => {
    loadPPTLinkedItems();
  }, []);

  const loadPPTLinkedItems = async () => {
    setResolvingFile(true);
    try {
      const items = await getPPTLinkedItems();
      setPptLinks(items);
    } catch (e) {
      console.error("Failed to load PPT linked items:", e);
    } finally {
      setResolvingFile(false);
    }
  };

  return (
    <Box
      sx={{
        height: "100vh",
        display: "flex",
        flexDirection: "column",
        bgcolor: "#FFFFFF",
        fontFamily: "Segoe UI, Arial, sans-serif",
        overflow: "hidden",
      }}
    >
      {/* Header Container */}
      <Box
        sx={{
          position: "sticky",
          top: 0,
          zIndex: 100,
          p: 2,
          bgcolor: "#0078d4",
          color: "#FFFFFF",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
        }}
      >
        <Box sx={{ width: 32 }} />
        <Typography
          sx={{
            fontWeight: 700,
            fontSize: "16px",
            letterSpacing: "0.5px",
            fontFamily: "Segoe UI, Arial",
            textAlign: "center",
          }}
        >
          EXCEL TO POWERPOINT
        </Typography>
        
        {/* Dynamic Refresh / Loading Indicator Button [1] */}
        <IconButton 
          onClick={loadPPTLinkedItems} 
          size="small" 
          sx={{ color: "#FFFFFF" }}
          disabled={resolvingFile}
        >
          {resolvingFile ? (
            <CircularProgress size={18} sx={{ color: "#FFFFFF" }} />
          ) : (
            <Refresh sx={{ fontSize: 20 }} />
          )}
        </IconButton>
      </Box>

      {/* Main Panel Content Area */}
      <Box
        sx={{
          p: 2,
          flex: 1,
          overflowY: "auto",
          display: "flex",
          flexDirection: "column",
          gap: 2,
          maxHeight: "95vh",
          "&::-webkit-scrollbar": { display: "none" },
          msOverflowStyle: "none",
          scrollbarWidth: "none",
        }}
      >
        <LinkComponentSafe onLinkSuccess={loadPPTLinkedItems} />

        <ActiveConnectionsSafe pptLinks={pptLinks} onLinkSuccess={loadPPTLinkedItems} />
      </Box>

      <Box sx={{ p: 1.5, textAlign: "center", borderTop: "1px solid #EDEBE9" }}>
        <Typography
          sx={{ fontSize: "10px", color: "#A19F9D", fontWeight: 600, fontFamily: "Segoe UI, Arial" }}
        >
          Version 1.0.0 Stable
        </Typography>
      </Box>
    </Box>
  );
};

export default Dashboard;