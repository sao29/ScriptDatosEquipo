### Inventario de Equipo (CMDB)

Aplicación de consola en C# que obtiene información de un equipo Windows, la convierte a JSON y la envía por HTTP POST a una CMDB.

#### Datos enviados
- Equipo: `nombre_equipo`, `modelo`, `serial`, `uuid`
- SO: `sistema_operativo`
- CPU: `cpu`, `cores`, `threads`
- RAM: `ram` (en bytes)
- Red: `ip`, `mac`
- Motherboard: `tarjeta_madre`, `serial_motherboard`
- Office: `office_licencia_ultimos5`

#### Flujo
1. Recolecta datos (WMI/red/Office)
2. Genera JSON
3. Envía POST
4. Muestra resultado en consola

#### Endpoint
