export function teklifVerUrl(productName: string): string {
  const message = `Merhaba, ${productName} Excel sistemi için teklif vermek istiyorum.`;
  return `https://wa.me/905393333303?text=${encodeURIComponent(message)}`;
}
