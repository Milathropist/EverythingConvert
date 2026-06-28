export function appAssetUrl (path: string): string {
  const cleanPath = path.replace(/^\/+/, "");
  return new URL(`../${cleanPath}`, import.meta.url).toString();
}
