import React from "react";
import ReactDOM from "react-dom/client";
import { createTheme, CssBaseline, ThemeProvider } from "@mui/material";
import ScanConverter from "./ScanConverter";

// Тема под фирменную палитру Сбера (зелёный #21A038). В sberAct компонент
// возьмёт их собственную MUI-тему — здесь она только для автономного dev-запуска.
const theme = createTheme({
  palette: {
    primary: { main: "#21A038", dark: "#158A2B", light: "#5FBF6E", contrastText: "#fff" },
    success: { main: "#0F8A2E" },
    error: { main: "#E53935" },
    background: { default: "#F4F8F5" },
  },
  shape: { borderRadius: 12 },
  typography: { fontFamily: "Segoe UI, Roboto, Arial, sans-serif" },
});

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <ScanConverter />
    </ThemeProvider>
  </React.StrictMode>,
);
