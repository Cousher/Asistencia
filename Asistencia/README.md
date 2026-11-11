# Asistencia en puesto

Sistema para que los supervisores de Frigorífico Bermejo S.A. tomen asistencia en el puesto de trabajo desde el celular, con una planilla consolidada y un resumen diario para RR. HH.

Funciona sobre Google Sheets y Google Apps Script: no requiere servidor propio ni que los supervisores tengan cuenta de Google. Cada supervisor ingresa con su nombre y un PIN.

## Qué hace

**Página para supervisores (celular)**
- Lista de pendientes de *Mi gente* (según la columna *Reporta a* del padrón), con búsqueda por apellido o legajo sobre toda la dotación.
- Filtros encadenados por área, sub área y tarea; turnos TM y TT.
- Registro de presente o ausente con horas extra y observación en un toque.
- Funciona sin señal: guarda lo marcado en el celular y lo envía al volver la conexión.
- Coberturas por franco o vacaciones con la columna *También ve a*.

**Hoja de cálculo (RR. HH.)**
- `Asistencia`: una fila por registro (fecha, turno, legajo, estado, HHEE, observación, supervisor, hora, tarea).
- `Resumen`: totales, por área y sub área, por supervisor, ausentes, horas extra, sin registrar y dobles registros. Se actualiza cada 10 minutos.
- Reporte por mail a demanda o automático a una hora fija.

## Estructura

```
src/
  Code.gs          Lógica del lado de Google: hojas, PIN, lectura, guardado, resumen y mail
  Index.html       Página que usan los supervisores
  appsscript.json  Manifiesto del proyecto (zona horaria y configuración de la app web)
```

## Instalación

1. Crear una hoja de cálculo de Google y, en **Archivo > Configuración**, poner configuración regional Argentina y zona horaria Buenos Aires.
2. Abrir **Extensiones > Apps Script**. Reemplazar el contenido de `Código.gs` por `src/Code.gs` y crear un archivo HTML llamado exactamente `Index` con el contenido de `src/Index.html`.
3. En el desplegable de funciones elegir **configurar** (por defecto aparece `doGet`) y ejecutarla. Aceptar los permisos.
4. Cargar las pestañas `Padrón` y `Supervisores`.
5. **Implementar > Nueva implementación > Aplicación web**, con *Ejecutar como: Yo* y *Quién tiene acceso: Cualquier usuario*. Repartir la URL que termina en `/exec`.

### Pestañas que completa RR. HH.

**Padrón** (las columnas se reconocen por el título, el orden puede cambiar)

| Legajo | Apellido y nombre | Área | Sub área | Tarea | Turno | Reporta a |
| --- | --- | --- | --- | --- | --- | --- |

**Supervisores**

| Nombre | PIN | Activo | También ve a |
| --- | --- | --- | --- |

**Resumen**, fila 1: `B1` fecha, `E1` turno, `K1` correos para el reporte, `N1` hora del envío automático.

## Actualizar el código

1. Pegar los archivos nuevos en el editor y guardar.
2. Si cambió `Code.gs`, ejecutar **configurar** (no borra datos).
3. Si cambió algo que usa la página de los supervisores, ir a **Implementar > Gestionar implementaciones**, editar, elegir **Nueva versión** e implementar. La URL no cambia.

### Con clasp (opcional)

```bash
npm install -g @google/clasp
clasp login
clasp clone <ID_DEL_SCRIPT> --rootDir src   # o clasp create, la primera vez
clasp push
```

## Seguridad y datos

- Este repositorio contiene solo código. El padrón, los PIN y los registros quedan en la hoja de cálculo, que no debe compartirse fuera de RR. HH.
- No subir al repositorio la URL de la aplicación web, exportaciones del padrón ni el archivo `.clasp.json`.
- El celular guarda un token de sesión, no el PIN. La sesión vence a los 60 días o cuando RR. HH. cambia el PIN o desactiva al supervisor.
- Tras 5 intentos de PIN incorrectos, ese supervisor no puede volver a ingresar durante 10 minutos (las sesiones ya abiertas siguen funcionando).
- Los textos que empiezan con = + - @ se guardan como texto, nunca como fórmula.
