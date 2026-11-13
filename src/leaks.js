/**
 * Per alias: which sender domains wrote to it. The domain with the most mail is "home".
 * An alias is flagged when another domain is not trusted.
 */
export function leakReport(records, trusted) {
  const byAlias = new Map();
  for (const { alias, domain } of records) {
    const key = domain || '(unknown)';
    if (!byAlias.has(alias)) byAlias.set(alias, new Map());
    const counts = byAlias.get(alias);
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  return [...byAlias]
    .map(([alias, counts]) => {
      const sorted = [...counts].sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : 1));
      const home = sorted[0][0];
      const ok = new Set(trusted[alias] || []);
      const domains = sorted.map(([domain, count]) => ({ domain, count, trusted: domain === home || ok.has(domain) }));
      return {
        alias,
        total: sorted.reduce((sum, [, count]) => sum + count, 0),
        home,
        domains,
        flagged: domains.some(d => !d.trusted),
      };
    })
    .sort((a, b) => b.flagged - a.flagged || b.total - a.total);
}
