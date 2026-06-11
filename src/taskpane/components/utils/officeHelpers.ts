// declare const PowerPoint: any;
// export interface PPTLinkedItem {
//   id: string;
//   shapeId: string;
//   excelFileId: string;
//   excelFileName: string;
//   sheetName: string;
//   rangeAddress: string;
//   type: "Table" | "Chart";
// }

// export const getPPTLinkedItems = async (): Promise<PPTLinkedItem[]> => {
//   return await PowerPoint.run(async (context: any) => {
//     const presentation = context.presentation;
//     const slides = presentation.slides;
//     slides.load("items");
//     await context.sync();

//     const linkedItems: PPTLinkedItem[] = [];

//     for (const slide of slides.items) {
//       const shapes = slide.shapes;
//       shapes.load("items/id");
//       await context.sync();

//       for (const shape of shapes.items) {
//         shape.tags.load("items");
//       }

//       try {
//         await context.sync();
//       } catch (err) {
//         console.error("Failed to load shape tags resiliently:", err);
//         continue;
//       }

//       for (const shape of shapes.items) {
//         const tags = shape.tags;
//         if (!tags || !tags.items) continue;

//         let linkId = "";
//         let excelFileId = "";
//         let excelFileName = "";
//         let sheetName = "";
//         let rangeAddress = "";
//         let type: "Table" | "Chart" = "Table";
//         let isLinked = false;

//         for (const tag of tags.items) {
//           if (tag.key === "LIVE_LINK_ID") {
//             linkId = tag.value;
//             isLinked = true;
//           } else if (tag.key === "EXCEL_FILE_ID") {
//             excelFileId = tag.value;
//           } else if (tag.key === "EXCEL_FILE_NAME") {
//             excelFileName = tag.value;
//           } else if (tag.key === "EXCEL_SHEET_NAME") {
//             sheetName = tag.value;
//           } else if (tag.key === "EXCEL_RANGE_ADDRESS") {
//             rangeAddress = tag.value;
//           } else if (tag.key === "TYPE") {
//             type = tag.value as "Table" | "Chart";
//           }
//         }

//         if (isLinked && linkId) {
//           linkedItems.push({
//             id: linkId,
//             shapeId: shape.id,
//             excelFileId,
//             excelFileName,
//             sheetName,
//             rangeAddress,
//             type,
//           });
//         }
//       }
//     }

//     return linkedItems;
//   });
// };

// export const deletePPTShape = async (shapeId: string): Promise<void> => {
//   await PowerPoint.run(async (context: any) => {
//     const slides = context.presentation.slides;
//     slides.load("items");
//     await context.sync();

//     for (const slide of slides.items) {
//       const shapes = slide.shapes;
//       shapes.load("items/id");
//       await context.sync();

//       const targetShape = shapes.items.find((s: any) => s.id === shapeId);
//       if (targetShape) {
//         targetShape.delete();
//         await context.sync();
//         break;
//       }
//     }
//   });
// };

// export const clearPPTShapeTags = async (shapeId: string): Promise<void> => {
//   await PowerPoint.run(async (context: any) => {
//     const slides = context.presentation.slides;
//     slides.load("items");
//     await context.sync();

//     for (const slide of slides.items) {
//       const shapes = slide.shapes;
//       shapes.load("items");
//       await context.sync();

//       for (const shape of shapes.items) {
//         shape.load("id");
//       }
//       await context.sync();

//       const targetShape = shapes.items.find((s: any) => s.id === shapeId);
//       if (targetShape) {
//         targetShape.tags.delete("LIVE_LINK_ID");
//         targetShape.tags.delete("EXCEL_FILE_ID");
//         targetShape.tags.delete("EXCEL_FILE_NAME");
//         targetShape.tags.delete("EXCEL_SHEET_NAME");
//         targetShape.tags.delete("EXCEL_RANGE_ADDRESS");
//         targetShape.tags.delete("TYPE");

//         await context.sync();
//         break;
//       }
//     }
//   });
// };


declare const PowerPoint: any;

export interface PPTLinkedItem {
  id: string;
  shapeId: string;
  slideId: string; // Dynamic slide tracking ke liye property add ki hai
  excelFileId: string;
  excelFileName: string;
  sheetName: string;
  rangeAddress: string;
  type: "Table" | "Chart";
}

export const getPPTLinkedItems = async (): Promise<PPTLinkedItem[]> => {
  return await PowerPoint.run(async (context: any) => {
    const presentation = context.presentation;
    const slides = presentation.slides;
    slides.load("items");
    await context.sync();

    const linkedItems: PPTLinkedItem[] = [];

    for (const slide of slides.items) {
      slide.load("id"); // Slide ID ko load kiya taake sidebar inactive slides ke images track kare
      await context.sync();

      const shapes = slide.shapes;
      shapes.load("items/id");
      await context.sync();

      for (const shape of shapes.items) {
        shape.tags.load("items");
      }

      try {
        await context.sync();
      } catch (err) {
        console.error("Failed to load shape tags resiliently:", err);
        continue;
      }

      for (const shape of shapes.items) {
        const tags = shape.tags;
        if (!tags || !tags.items) continue;

        let linkId = "";
        let excelFileId = "";
        let excelFileName = "";
        let sheetName = "";
        let rangeAddress = "";
        let type: "Table" | "Chart" = "Table";
        let isLinked = false;

        for (const tag of tags.items) {
          if (tag.key === "LIVE_LINK_ID") {
            linkId = tag.value;
            isLinked = true;
          } else if (tag.key === "EXCEL_FILE_ID") {
            excelFileId = tag.value;
          } else if (tag.key === "EXCEL_FILE_NAME") {
            excelFileName = tag.value;
          } else if (tag.key === "EXCEL_SHEET_NAME") {
            sheetName = tag.value;
          } else if (tag.key === "EXCEL_RANGE_ADDRESS") {
            rangeAddress = tag.value;
          } else if (tag.key === "TYPE") {
            type = tag.value as "Table" | "Chart";
          }
        }

        if (isLinked && linkId) {
          linkedItems.push({
            id: linkId,
            shapeId: shape.id,
            slideId: slide.id, // Store loaded slide ID [1]
            excelFileId,
            excelFileName,
            sheetName,
            rangeAddress,
            type,
          });
        }
      }
    }

    return linkedItems;
  });
};

export const deletePPTShape = async (shapeId: string): Promise<void> => {
  await PowerPoint.run(async (context: any) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items/id");
      await context.sync();

      const targetShape = shapes.items.find((s: any) => s.id === shapeId);
      if (targetShape) {
        targetShape.delete();
        await context.sync();
        break;
      }
    }
  });
};

export const clearPPTShapeTags = async (shapeId: string): Promise<void> => {
  await PowerPoint.run(async (context: any) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      for (const shape of shapes.items) {
        shape.load("id");
      }
      await context.sync();

      const targetShape = shapes.items.find((s: any) => s.id === shapeId);
      if (targetShape) {
        targetShape.tags.delete("LIVE_LINK_ID");
        targetShape.tags.delete("EXCEL_FILE_ID");
        targetShape.tags.delete("EXCEL_FILE_NAME");
        targetShape.tags.delete("EXCEL_SHEET_NAME");
        targetShape.tags.delete("EXCEL_RANGE_ADDRESS");
        targetShape.tags.delete("TYPE");

        await context.sync();
        break;
      }
    }
  });
};