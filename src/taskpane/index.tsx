import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require, HTMLElement */

const title = "Contoso Task Pane Add-in";

// Use React.lazy for component lazy loading
const App = React.lazy(() => import("./components/App"));

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <React.Suspense fallback={<div style={{ display: "flex", justifyContent: "center", alignItems: "center", height: "100%" }}>Loading...</div>}>
        <App />
      </React.Suspense>
    </FluentProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(
      <FluentProvider theme={webLightTheme}>
        <React.Suspense fallback={<div style={{ display: "flex", justifyContent: "center", alignItems: "center", height: "100%" }}>Loading...</div>}>
          <NextApp />
        </React.Suspense>
      </FluentProvider>
    );
  });
}
