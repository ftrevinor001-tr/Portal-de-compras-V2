# Portal de Compras V2

Portal completo de Compras: los dashboards de operación de siempre más el módulo de planeación, con indicadores calculados por el portal (antecedente y pronóstico por clave, propuesta de compra, alarmas, Target y OTIF).

Liga: https://ftrevinor001-tr.github.io/Portal-de-compras-V2/ (misma contraseña del portal actual)

## Archivos

| Archivo | Pestaña | Quién la ve |
|---|---|---|
| `index.html` | Portal: acceso, Mi Resumen, Admin y pestañas | Todos |
| `dash_faltantes.html` | Histórico Faltantes | Todos |
| `dash_monitor.html` | Monitor de Inventario | Todos |
| `dash_backorders.html` | Estatus Back Orders | Todos |
| `dash_planeacion.html` | Planeación e indicadores (Tablero, Detalle por clave, Pedidos propuestos, Target, OTIF) | Todos |
| `plan_oc.html` | Compromiso OCs | Todos |
| `dash_otif.html` | Indicadores (Excel), solo durante el paralelo de noviembre | Todos |
| `plan_admin.html` | Cálculo diario (cargas y cálculo) | Solo administrador |
| `plan_logica.js` | Reglas de cálculo compartidas por el cálculo diario, el tablero y Mi Resumen | — |
| `resumen_v2.js` | Mi Resumen con el cálculo del portal | — |

## Qué sale del cálculo del portal (no del Excel)

- Mi Resumen: Claves en riesgo, Sobre inventario, Pedidos propuestos, Target y OTIF.
- Planeación e indicadores: todo.

Faltantes, faltantes nuevos, prolongados y OCs vencidas siguen saliendo del Excel maestro (Admin).

## Base de datos

Proyecto Supabase SANVER_Forte. El módulo de planeación solo escribe en tablas `plan_` (scripts `plan_tablas.sql` a `plan_tablas_6.sql`). Los dashboards de operación usan las mismas tablas que el portal actual (Portal-de-operaciones-Compas): **sube el Excel maestro en uno solo de los dos**.

## Rutina diaria

1. Admin → Cargar Excel maestro (actualiza los dashboards de operación).
2. Cálculo diario → 6. Cálculo diario con el mismo respaldo del día (actualiza Planeación e indicadores y Mi Resumen).
3. Compradores → Compromiso OCs: capturan o suben las fechas reales del proveedor.
