import { initStorage } from "./storageShim.js";
import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";

async function start() {
  await initStorage();
  ReactDOM.createRoot(document.getElementById("root")).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
}

start();
