import type { WorkbookPass, WorkbookSettings } from './sheetTypes';

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(`/api/${path}`, {
    credentials: 'same-origin',
    ...init,
    headers: {
      ...(init?.body ? { 'Content-Type': 'application/json' } : {}),
      ...init?.headers,
    },
  });
  const body = await response.text();
  let result: any = {};
  if (body.trim()) {
    try {
      result = JSON.parse(body);
    } catch {
      throw new Error(`The server returned an invalid response (HTTP ${response.status}).`);
    }
  }
  if (!response.ok) {
    const deploymentHint = response.status >= 500
      ? ' Check Vercel logs and Google Sheets environment variables.'
      : '';
    throw new Error(`${result.error || `Request failed (HTTP ${response.status}).`}${deploymentHint}`);
  }
  if (!body.trim()) throw new Error(`The server returned an empty response (HTTP ${response.status}).`);
  return result as T;
}

export async function getAuthStatus() {
  return request<{ configured: boolean; authenticated: boolean }>('auth/status');
}

export async function login(password: string) {
  return request<{ authenticated: boolean }>('auth/login', {
    method: 'POST',
    body: JSON.stringify({ password }),
  });
}

export async function logout() {
  return request<{ authenticated: boolean }>('auth/logout', { method: 'POST' });
}

export async function loadSpreadsheet() {
  return request<{ settings: WorkbookSettings; passes: WorkbookPass[] }>('bootstrap');
}

export async function savePass(payload: Partial<WorkbookPass>, editingId: number | null) {
  const path = editingId ? `gate-passes/${editingId}` : 'gate-passes';
  return request<WorkbookPass>(path, {
    method: editingId ? 'PUT' : 'POST',
    body: JSON.stringify(payload),
  });
}

export async function deletePass(id: number) {
  return request<{ success: boolean }>(`gate-passes/${id}`, { method: 'DELETE' });
}

export async function duplicatePass(id: number) {
  return request<WorkbookPass>(`gate-passes/${id}/duplicate`, { method: 'POST' });
}

export async function saveSettings(settings: WorkbookSettings) {
  return request<WorkbookSettings>('settings', {
    method: 'PUT',
    body: JSON.stringify(settings),
  });
}
