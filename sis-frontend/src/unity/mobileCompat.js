/**
 * Unity SIS — Mobile compatibility (pure JS, no JSX)
 * Place at: src/unity/mobileCompat.js
 *
 * import { MobileStyles, useIsMobile } from "./unity/mobileCompat.js";
 */
import { useState, useEffect } from "react";

const STYLE_ID = "uu-mobile-compat-css";

const MOBILE_CSS = 'html, body, #root { max-width: 100vw; overflow-x: hidden; -webkit-text-size-adjust: 100%; }\n.uu-module-root, .uu-module-root > div { height: 100%; min-height: 0; }\n.uu-main { min-height: 0; }\n@media (max-width: 767px) {\n  .uu-app-header { height: auto !important; min-height: 3rem; padding: 0.5rem 0.75rem !important; flex-wrap: wrap; gap: 0.5rem; }\n  .uu-app-header h1 { font-size: 0.95rem !important; }\n  .uu-app-header .uu-user-block { border-right: none !important; padding-right: 0 !important; }\n  .uu-sidebar-desktop { display: none !important; }\n  .uu-sidebar-drawer { position: fixed !important; inset: 0 auto 0 0; width: min(86vw, 320px) !important; max-width: 320px; z-index: 60; transform: translateX(-105%); transition: transform 0.22s ease; box-shadow: 8px 0 40px rgba(0,0,0,0.45); }\n  .uu-sidebar-drawer.uu-open { transform: translateX(0); }\n  .uu-sidebar-backdrop { position: fixed; inset: 0; background: rgba(0,0,0,0.55); z-index: 55; }\n  .uu-main { padding: 0.75rem !important; width: 100% !important; max-width: 100vw; }\n  .uu-table-scroll, .overflow-x-auto { overflow-x: auto !important; -webkit-overflow-scrolling: touch; max-width: 100%; }\n  .uu-table-scroll table, .overflow-x-auto table { min-width: 560px; }\n  .uu-kpi-grid, .grid.grid-cols-4, .grid.grid-cols-3 { grid-template-columns: repeat(2, minmax(0, 1fr)) !important; }\n  .grid.grid-cols-2 { grid-template-columns: repeat(1, minmax(0, 1fr)) !important; }\n  .uu-modal, [role="dialog"] { max-width: calc(100vw - 1rem) !important; margin: 0.5rem !important; }\n  button, .uu-touch { min-height: 40px; }\n  input, select, textarea { font-size: 16px !important; }\n  .uu-module-bar { flex-wrap: wrap; height: auto !important; min-height: 3rem; padding: 0.4rem 0.6rem !important; }\n  .uu-module-bar button { font-size: 0.65rem !important; padding: 0.35rem 0.6rem !important; }\n}\n@media (max-width: 480px) {\n  .uu-kpi-grid, .grid.grid-cols-4, .grid.grid-cols-3, .grid.grid-cols-2 { grid-template-columns: repeat(1, minmax(0, 1fr)) !important; }\n}';

export function useIsMobile(breakpoint) {
  var bp = typeof breakpoint === "number" ? breakpoint : 768;
  var [mobile, setMobile] = useState(function () {
    if (typeof window === "undefined") return false;
    return window.innerWidth < bp;
  });

  useEffect(
    function () {
      if (typeof window === "undefined") return undefined;
      function onResize() {
        setMobile(window.innerWidth < bp);
      }
      window.addEventListener("resize", onResize);
      onResize();
      return function () {
        window.removeEventListener("resize", onResize);
      };
    },
    [bp]
  );

  return mobile;
}

export function MobileStyles() {
  useEffect(function () {
    if (typeof document === "undefined") return undefined;
    var el = document.getElementById(STYLE_ID);
    if (!el) {
      el = document.createElement("style");
      el.id = STYLE_ID;
      el.setAttribute("data-uu", "mobile-compat");
      el.textContent = MOBILE_CSS;
      document.head.appendChild(el);
    }
    return undefined;
  }, []);

  return null;
}

export default { useIsMobile: useIsMobile, MobileStyles: MobileStyles };
