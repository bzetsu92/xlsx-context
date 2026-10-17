// Worker entry. Installs DOM globals before the core loads. Exposes the same public functions.
import "./worker-globals.js";
export * from "./index.js";
