import XCF from "../src/main.js";
import { readFileSync, writeFileSync } from "node:fs";

const bytes = readFileSync(process.argv[2] || "test/samples/simple.xcf");

let pixel_data;
for (let i = 0; i < 10; i++) {
    const xcf = XCF.from_bytes(bytes);
    pixel_data = xcf.getLayerPixels(0);
}
writeFileSync("test/output.json", JSON.stringify(pixel_data));