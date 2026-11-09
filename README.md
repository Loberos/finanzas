# Mis Finanzas

Formulario web para controlar tus finanzas personales en soles (S/) y dólares (US$). **La base de datos es una hoja de Google Sheets**: cada movimiento se escribe ahí y el formulario lee desde ahí. Puedes editar la hoja a mano y el formulario lo refleja.

## Qué hace

- Registra gastos e ingresos con fecha, monto, moneda, tipo de cambio, categoría, medio de pago y descripción.
- Muestra ingresos, gastos, balance y tasa de ahorro del mes.
- Presupuesto por categoría con alertas al 80 % ("Cerca del límite") y al 100 % ("Excedido").
- Botones para abrir la hoja y para descargarla como Excel (.xlsx).

## Pestañas que se crean en la hoja

| Pestaña | Contenido |
|---|---|
| Movimientos | La base de datos. `Mes` y `Monto S/` se calculan solos, también en las filas que escribas a mano. |
| Presupuesto | Límite por categoría, gastado, restante y estado. Cambia el mes en `B1`. |
| Resumen mensual | Ingresos, gastos, balance y % de ahorro de cada mes. |
| Gastos por categoría | Tabla dinámica de categorías por mes. |
| Config | Tipo de cambio por defecto. |

## Instalación (una sola vez)

1. Crea una hoja vacía en Google Sheets (https://sheets.new) y llámala **Finanzas**.
2. En la hoja: **Extensiones → Apps Script**. Borra lo que haya y pega `apps-script/Code.gs`.
3. Cambia `CLAVE` por una clave secreta tuya y guarda.
4. Arriba, elige la función **configurar** y pulsa **Ejecutar**. Acepta los permisos (si aparece "Google no verificó esta app": *Configuración avanzada → Ir a…*). Se crean las pestañas.
5. **Implementar → Nueva implementación → Tipo: Aplicación web**.
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
6. Copia la URL que termina en `/exec`.
7. Abre `index.html`, pega la URL y la clave, y pulsa **Conectar**. Se recuerdan en ese navegador.

Si cambias el código de Apps Script, publica una nueva versión en **Implementar → Administrar implementaciones → Editar → Nueva versión** para que la URL use el código nuevo.

## Seguridad

La URL de la aplicación web acepta pedidos de cualquiera, pero el script rechaza todo pedido que no traiga la `CLAVE`. No subas tu clave real al repositorio.
