npm audit --json | node -e "
const data = [];
process.stdin.on('data', d => data.push(d));
process.stdin.on('end', () => {
  const audit = JSON.parse(data.join(''));
  const vulns = audit.vulnerabilities;
  
  const filtered = Object.entries(vulns)
    .filter(([_, v]) => ['critical', 'high'].includes(v.severity))
    .map(([name, v]) => ({
      name,
      severity: v.severity,
      title: v.via[0]?.title || v.via[0],
      url: v.via[0]?.url || 'N/A',
      fixAvailable: v.fixAvailable ? '✅ Fix available' : '❌ No fix yet'
    }));

  console.log('CRITICAL (' + filtered.filter(v => v.severity === 'critical').length + '):');
  filtered.filter(v => v.severity === 'critical').forEach(v => {
    console.log(' -', v.name, '|', v.title, '|', v.fixAvailable);
  });

  console.log('');
  console.log('HIGH (' + filtered.filter(v => v.severity === 'high').length + '):');
  filtered.filter(v => v.severity === 'high').forEach(v => {
    console.log(' -', v.name, '|', v.title, '|', v.fixAvailable);
  });
});
"