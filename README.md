# VB.NET 4.8 CRUD Generator

Este repositorio incluye un generador que produce, a partir de metadatos de tabla:

- Una clase VB estilo “cls_<tabla>” con campos privados, propiedades y métodos CRUD completos (Insert, Update, Delete, List) usando ADO.NET con consultas parametrizadas, similar a los patrones de VB tradicionales.
- Una página ASPX con un formulario y `GridView` para alta/edición/borrado.
- El code-behind de la página, que usa la clase generada para persistencia, conversión de tipos y mensajes de estado.

Todo el código es compatible con .NET Framework 4.8 o versiones anteriores.

## Archivos clave
- `src/CrudGenerator.vb`: lógica de generación de clases y páginas.
- `src/Program.vb`: ejemplo de uso que genera artefactos de demostración a partir de una definición de tabla.

## Uso rápido
1. Ajusta la definición de tabla en `src/Program.vb` (nombre, columnas, tipos, clave primaria).
2. Compila con el compilador de VB de .NET Framework 4.8:
   ```powershell
   vbc /target:exe /out:CrudGenerator.exe src/Program.vb src/CrudGenerator.vb
   ```
3. Ejecuta el generador:
   ```powershell
   .\CrudGenerator.exe
   ```
4. Revisa la carpeta `output/<NombreTabla>/` para encontrar:
   - `<NombreTabla>.vb` (clase VB con CRUD ADO.NET)
   - `<NombreTabla>.aspx` (markup)
   - `<NombreTabla>.aspx.vb` (code-behind)

## Personalización
- Edita `ConnectionString` en `src/CrudGenerator.vb` para apuntar a tu base de datos o deja que el code-behind use la misma cadena embebida.
- Ajusta `NamespaceName` en `src/Program.vb` para que coincida con tu solución.
- Si necesitas tipos adicionales, agrega al diccionario `SqlTypeMap` en `CrudGenerator.vb`.

## Consideraciones
- El CRUD generado es básico y no incluye validaciones avanzadas ni control de transacciones; úsalo como punto de partida y adáptalo.
- Todas las consultas usan parámetros para minimizar el riesgo de inyección SQL.
- Los métodos devuelven mensajes de error en un `ByRef` para poder mostrarlos en la UI o logs.
