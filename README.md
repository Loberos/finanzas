# Mis Finanzas

Formulario web para controlar tus finanzas personales en soles (S/) y dólares (US$), con presupuesto mensual y exportación a Excel.

## Qué hace

- Registra gastos e ingresos con fecha, monto, moneda, tipo de cambio, categoría, medio de pago y descripción.
- Convierte los montos en dólares a soles con el tipo de cambio que indiques.
- Muestra ingresos, gastos, balance y tasa de ahorro del mes.
- Presupuesto por categoría con alertas al 80 % ("Cerca del límite") y al 100 % ("Excedido").
- **Descargar Excel**: genera un `.xlsx` con fórmulas reales en 4 hojas:
  - **Movimientos**: todo el registro, con la conversión a S/ por fórmula.
  - **Resumen mensual**: ingresos, gastos, balance y % de ahorro por mes.
  - **Presupuesto**: cambia el mes en la celda `B1` y todo se recalcula.
  - **Gastos por categoría**: categorías por mes.
- **Importar Excel**: carga un Excel descargado desde la app (por ejemplo, para pasar tus datos de la PC al celular). Omite los movimientos repetidos.

## Dónde se guardan los datos

En el `localStorage` del navegador. No se envían a ningún servidor. Cada navegador y dispositivo tiene sus propios datos, así que conviene descargar el Excel de vez en cuando como respaldo. En modo incógnito no se guardan.

## Uso

Es un solo archivo, `index.html`, sin dependencias que instalar (SheetJS se carga desde cdnjs). Ábrelo en el navegador o publícalo con GitHub Pages.
