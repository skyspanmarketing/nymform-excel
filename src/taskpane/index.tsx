// Entry point of the task pane.
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import { officeHost } from "./host";

const el = document.getElementById("root");
if (el) {
  createRoot(el).render(
    <StrictMode>
      <App host={officeHost} />
    </StrictMode>,
  );
}
