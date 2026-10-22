// Hands serialised bytes to the browser as a file download: a transient
// anchor pointed at an object URL, revoked once the download has had time
// to start. Shared by every product's download path (lib/downloadDocument,
// lib/presentation/download) so the mechanism exists exactly once.
export function triggerBlobDownload(bytes: Uint8Array, fileName: string, mimeType: string): void {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: mimeType }));
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}
