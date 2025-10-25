/**
 * Minimal API client for the frontend to talk to the backend via HTTP only.
 *
 * Behavior:
 * - If REACT_APP_API_BASE is set, use absolute URLs: `${BASE}${path}`.
 * - Otherwise, use relative paths (e.g., `/api/...`) to work with dev proxy or same-origin deployments.
 */
import { getApiBase } from '../utils/runtimeConfig';

const base = (() => {
  try { return (getApiBase() || '').replace(/\/$/, ''); } catch { return ''; }
})();

const makeUrl = (path: string) => {
  const p = path.startsWith('/') ? path : `/${path}`;
  return base ? `${base}${p}` : p; // relative when base is empty
};

type Json = any;

export async function apiGet<T = Json>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(makeUrl(path), { method: 'GET', ...(init || {}) });
  if (!res.ok) throw new Error(`GET ${path} failed: ${res.status}`);
  return (await res.json()) as T;
}

export async function apiPost<T = Json>(path: string, body?: any, init?: RequestInit): Promise<T> {
  const res = await fetch(makeUrl(path), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...(init?.headers || {}) },
    body: body !== undefined ? JSON.stringify(body) : undefined,
    ...init
  });
  if (!res.ok) throw new Error(`POST ${path} failed: ${res.status}`);
  return (await res.json()) as T;
}

export async function apiPut<T = Json>(path: string, body?: any, init?: RequestInit): Promise<T> {
  const res = await fetch(makeUrl(path), {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json', ...(init?.headers || {}) },
    body: body !== undefined ? JSON.stringify(body) : undefined,
    ...init
  });
  if (!res.ok) throw new Error(`PUT ${path} failed: ${res.status}`);
  return (await res.json()) as T;
}

export async function apiDelete<T = Json>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(makeUrl(path), { method: 'DELETE', ...(init || {}) });
  if (!res.ok) throw new Error(`DELETE ${path} failed: ${res.status}`);
  return (await res.json()) as T;
}
