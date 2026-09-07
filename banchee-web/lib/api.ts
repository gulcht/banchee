export const API_BASE_URL =
  process.env.NEXT_PUBLIC_API_BASE_URL?.replace(/\/+$/, '') || 'http://localhost:8000';

export const BACKEND_INTERNAL_SECRET =
  process.env.NEXT_PUBLIC_BACKEND_INTERNAL_SECRET || '';

function getApiHeaders(): HeadersInit {
  const headers: Record<string, string> = {};
  if (BACKEND_INTERNAL_SECRET) {
    headers['x-backend-secret'] = BACKEND_INTERNAL_SECRET;
  }
  return headers;
}

export type SupportedDocType = 'sso' | 'pnd1' | 'pnd3' | 'pnd53' | 'statement';

export interface ConvertOptions {
  docType: SupportedDocType;
  format?: 'xlsx' | 'csv';
}

/**
 * Convert single or multiple PDF files for SSO, PND1, PND3, PND53, or Bank Statement via backend API.
 * - Single file: returns file directly (e.g. {docType}-MM-YYYY.xlsx or {docType}-MM-YYYY.csv)
 * - Multiple files: returns a ZIP archive (e.g. {docType}-YYYY.zip)
 */
export async function convertPdfDocuments(
  files: File[],
  docType: SupportedDocType = 'sso',
  format: 'xlsx' | 'csv' = 'xlsx'
): Promise<{ filename: string; isZip: boolean }> {
  if (!files || files.length === 0) {
    throw new Error('กรุณาเลือกไฟล์ PDF อย่างน้อย 1 ไฟล์');
  }

  const formData = new FormData();
  for (const file of files) {
    formData.append('files', file);
  }

  const url = `${API_BASE_URL}/convert/${encodeURIComponent(docType)}?format=${encodeURIComponent(format)}`;

  const response = await fetch(url, {
    method: 'POST',
    headers: getApiHeaders(),
    body: formData,
  });

  if (!response.ok) {
    let errorMessage = `การแปลงไฟล์ล้มเหลว (${response.status})`;
    try {
      const errorJson = (await response.json()) as { detail?: string };
      if (errorJson?.detail) {
        errorMessage = errorJson.detail;
      }
    } catch {
      // If response is not JSON, retain default error message
    }
    throw new Error(errorMessage);
  }

  // Parse filename from Content-Disposition header if available
  const contentDisposition = response.headers.get('content-disposition');
  let filename = files.length > 1 ? `${docType}.zip` : `${docType}-converted.${format}`;

  if (contentDisposition) {
    // Check filename*=UTF-8''... (RFC 5987)
    const utf8Match = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);
    if (utf8Match && utf8Match[1]) {
      filename = decodeURIComponent(utf8Match[1].trim().replace(/^["']|["']$/g, ''));
    } else {
      const normalMatch = contentDisposition.match(/filename="?([^";]+)"?/i);
      if (normalMatch && normalMatch[1]) {
        filename = normalMatch[1].trim();
      }
    }
  }

  const blob = await response.blob();
  const downloadUrl = window.URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = downloadUrl;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  window.URL.revokeObjectURL(downloadUrl);

  const isZip = filename.toLowerCase().endsWith('.zip');
  return { filename, isZip };
}

// Backward compatibility helper for SSO
export async function convertSsoPdfs(
  files: File[],
  format: 'xlsx' | 'csv' = 'xlsx'
): Promise<{ filename: string; isZip: boolean }> {
  return convertPdfDocuments(files, 'sso', format);
}

/**
 * Reconcile 12-month payroll documents (ภ.ง.ด. 1 vs สปส. 1-10) via backend API.
 */
export async function reconcilePayrollPdfs(
  files: File[],
  format: 'xlsx' | 'csv' = 'xlsx'
): Promise<{ filename: string }> {
  if (!files || files.length === 0) {
    throw new Error('กรุณาเลือกไฟล์ PDF สำหรับกระทบยอดอย่างน้อย 1 ไฟล์');
  }

  const formData = new FormData();
  for (const file of files) {
    formData.append('files', file);
  }

  const url = `${API_BASE_URL}/reconcile/payroll?format=${encodeURIComponent(format)}`;

  const response = await fetch(url, {
    method: 'POST',
    headers: getApiHeaders(),
    body: formData,
  });

  if (!response.ok) {
    let errorMessage = `การกระทบยอดล้มเหลว (${response.status})`;
    try {
      const errorJson = (await response.json()) as { detail?: string };
      if (errorJson?.detail) {
        errorMessage = errorJson.detail;
      }
    } catch {
      // If response is not JSON, retain default error message
    }
    throw new Error(errorMessage);
  }

  const contentDisposition = response.headers.get('content-disposition');
  let filename = `reconciliation-payroll.${format}`;

  if (contentDisposition) {
    const utf8Match = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);
    if (utf8Match && utf8Match[1]) {
      filename = decodeURIComponent(utf8Match[1].trim().replace(/^["']|["']$/g, ''));
    } else {
      const normalMatch = contentDisposition.match(/filename="?([^";]+)"?/i);
      if (normalMatch && normalMatch[1]) {
        filename = normalMatch[1].trim();
      }
    }
  }

  const blob = await response.blob();
  const downloadUrl = window.URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = downloadUrl;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  window.URL.revokeObjectURL(downloadUrl);

  return { filename };
}
