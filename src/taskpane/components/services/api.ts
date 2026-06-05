import axios from "axios";
// Is line ko update karein:
const API_BASE_URL = "https://live-link-backend.vercel.app/api/links";
// const API_BASE_URL = "http://localhost:5000/api/links";

const apiInstance = axios.create({
  baseURL: API_BASE_URL,
  headers: {
    "Content-Type": "application/json",
  },
});

export const getDistinctWorkbooks = async () => {
  const response = await apiInstance.get("/workbooks");
  return response.data;
};

export const getLinksByWorkbook = async (excelFileId: string) => {
  const response = await apiInstance.get(`/workbook/${encodeURIComponent(excelFileId)}`);
  return response.data;
};

export const getLinkDetails = async (linkId: string) => {
  const response = await apiInstance.get(`/${linkId}`);
  return response.data;
};

export const getBulkLinkDetails = async (linkIds: string[]) => {
  const response = await apiInstance.post("/bulk", { linkIds });
  return response.data;
};

export const deleteLinkData = async (linkId: string) => {
  const response = await apiInstance.delete(`/${linkId}`);
  return response.data;
};