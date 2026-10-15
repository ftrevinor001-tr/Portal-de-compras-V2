# Portal de Compras V2

Portal completo de Compras: los dashboards de operación de siempre más el módulo de planeación (antecedente/pronóstico, alarmas, Target, OTIF y fechas compromiso).

Liga: https://ftrevinor001-tr.github.io/Portal-de-compras-V2/ (misma contraseña del portal actual)

## Archivos

| Archivo | Pestaña | Quién la ve |
|---|---|---|
| `index.html` | Portal: acceso, Mi Resumen, Admin y pestañas | Todos |
| `dash_faltantes.html` | Histórico Faltantes | Todos |
| `dash_monitor.html` | Monitor de Inventario | Todos |
| `dash_backorders.html` | Estatus Back Orders | Todos |
| `dash_otif.html` | Indicadores | Todos |
| `plan_oc.html` | Compromiso OCs | Todos |
| `plan_admin.html` | Planeación (cargas y cálculo diario) | Solo administrador |

`plan_oc.html` y `plan_admin.html` también se pueden abrir solas por su liga.

## Base de datos

Proyecto Supabase SANVER_Forte. Los dashboards usan las mismas tablas que el portal actual (Portal-de-operaciones-Compas), así que los dos portales muestran lo mismo: **sube el Excel maestro en uno solo de los dos**. El módulo de planeación solo escribe en tablas `plan_`.

## Rutina diaria

1. Admin → Cargar Excel maestro (actualiza los dashboards).
2. Planeación → 6. Cálculo diario con el mismo respaldo del día.
3. Compradores → Compromiso OCs: capturan o suben las fechas reales del proveedor.
