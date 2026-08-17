// Sahne vurgu rengi paleti. Manifest fingerprint.accent alanı bu anahtarlara bağlanır.

const accentRenkleri: Record<string, string> = {
  blue: '#2563eb',
  green: '#0d8a4b',
  orange: '#f59e0b',
  navy: '#153b75',
  emerald: '#16a34a',
  red: '#e5484d',
  yellow: '#fbbf24',
};

export function accentRenk(anahtar: string): string {
  return accentRenkleri[anahtar] ?? accentRenkleri.green;
}
