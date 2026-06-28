export function appAssetUrl (path: string): string {
  const cleanPath = path.replace(/^\/+/, "");
  const baseUrl = typeof document === "undefined"
    ? new URL("../", globalThis.location.href).toString()
    : document.baseURI;

  return new URL(cleanPath, baseUrl).toString();
}
