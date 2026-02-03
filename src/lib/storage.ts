const KEY_PREFIX = "dagrooster.values.v1.";

export function loadValues(dateISO: string): Record<string, string> {
  try {
    const raw = localStorage.getItem(KEY_PREFIX + dateISO);
    return raw ? (JSON.parse(raw) as Record<string, string>) : {};
  } catch {
    return {};
  }
}

export function saveValues(dateISO: string, values: Record<string, string>) {
  localStorage.setItem(KEY_PREFIX + dateISO, JSON.stringify(values));
}
