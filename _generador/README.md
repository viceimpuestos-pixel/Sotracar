# Generador del Panel Financiero Sotracar

`index.html` (en la carpeta principal) es **autocontenido**: trae los datos embebidos.
Para publicarlo en GitHub Pages basta con reemplazar ese único archivo en el repo
`viceimpuestos-pixel/Sotracar` y hacer commit.

## Regenerar cuando haya nuevos libros auxiliares

Requisitos: Node.js (ya instalado en este equipo).

1. Coloca los `.xlsx` del Libro Auxiliar en la carpeta principal (uno por año).
2. Si cambian los nombres de archivo, edítalos en el objeto `files` de `etl2.js`.
3. Desde esta carpeta `_generador/`:

   ```bash
   npm install xlsx@0.18.5
   node etl2.js      # lee los .xlsx -> escribe datos.json + imprime un resumen de control
   node build.js     # datos.json + dashboard.template.html -> index.html
   ```

4. Copia el `index.html` resultante a la carpeta principal y súbelo al repo.

## Archivos

| Archivo | Qué hace |
|---|---|
| `etl2.js` | Lee los libros auxiliares, mensualiza movimientos débito/crédito, arma P&L, balance encadenado entre años, ratios, detalle por cuenta y concentración por tercero. Escribe `datos.json`. |
| `dashboard.template.html` | Plantilla del panel (HTML + CSS + JS de gráficos). El marcador `__DATOS__` se reemplaza por `datos.json`. |
| `build.js` | Inyecta `datos.json` en la plantilla y produce `index.html`. |
| `datos.json` | Datos agregados (generado). |

## Metodología (resumen)

- **P&L**: ingresos = grupo 41; costo del servicio = clases 6 y 7; gastos de
  administración = grupo 51 (incluye depreciación 5160/5165 como gasto operativo);
  gastos financieros = cuenta 5305; otros egresos = resto del grupo 53;
  otros ingresos = grupo 42; impuesto = grupo 54.
  EBITDA = utilidad operacional + depreciación y amortización.
- **Balance**: saldos de fin de mes por cuenta, encadenados entre años (las cuentas
  sin movimiento en un año conservan su saldo de cierre anterior). Patrimonio se
  presenta como activo − pasivo. Los años 2020–2024 quedan marcados con `~` /
  "en consolidación" porque faltan saldos de apertura de activos fijos anteriores a 2020;
  el estado de resultados no se ve afectado por esto.
