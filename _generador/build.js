const fs = require('fs');
const tpl = fs.readFileSync('dashboard.template.html', 'utf8');
const datos = fs.readFileSync('datos.json', 'utf8').replace(/</g, '\\u003c');
const body = tpl.replace('__DATOS__', () => datos);
const html = `<!doctype html>
<html lang="es">
<meta charset="utf-8">
${body}
</html>
`;
fs.writeFileSync('index.html', html);
console.log('index.html', (fs.statSync('index.html').size / 1024).toFixed(0) + ' KB');
