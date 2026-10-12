import React from "react";
import ReactDOM from "react-dom/client";
import { MsalProvider } from "@azure/msal-react";
import { EventType, type AuthenticationResult } from "@azure/msal-browser";
import { completeRedirect, msal } from "./auth/msal";
import { ThemeProvider } from "./theme/ThemeContext";
import { OperatorsProvider } from "./operators/OperatorsContext";
import { App } from "./App";
import "./styles.css";

msal
  .initialize()
  .then(completeRedirect)
  .then(() => {
    msal.addEventCallback((e) => {
      if (e.eventType === EventType.LOGIN_SUCCESS && e.payload) {
        msal.setActiveAccount((e.payload as AuthenticationResult).account);
      }
    });

    ReactDOM.createRoot(document.getElementById("root")!).render(
      <React.StrictMode>
        <ThemeProvider>
          <MsalProvider instance={msal}>
            <OperatorsProvider>
              <App />
            </OperatorsProvider>
          </MsalProvider>
        </ThemeProvider>
      </React.StrictMode>
    );
  });
