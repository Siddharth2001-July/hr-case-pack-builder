// Copies the Web SDK's WASM/worker assets into public/ so the app serves them itself
// (self-hosted: nothing is loaded from a Nutrient CDN at runtime).
import { cp, rm } from "node:fs/promises";
const src = new URL("../node_modules/@nutrient-sdk/viewer/dist/nutrient-viewer-lib", import.meta.url);
const dest = new URL("../public/nutrient-viewer-lib", import.meta.url);
await rm(dest, { recursive: true, force: true });
await cp(src, dest, { recursive: true });
console.log("Copied Web SDK assets to public/nutrient-viewer-lib");
