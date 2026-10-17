// Minimal Node typing for the Node entry. Avoids pulling @types/node into the browser-safe core typecheck.
declare module "node:module" {
    export function createRequire(url: string): (id: string) => unknown;
}
