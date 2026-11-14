---
name: iis-panax-development
description: Desarrollar integraciones con IIS Panax y el controlador genérico request.asp. Usar para definir consultas con FROM, proyecciones, filtros y argumentos SQL, o integrar un proyecto con local-server.
---

# Desarrollo con IIS Panax

Reutiliza `request.asp` para consultar los objetos autorizados. `FROM` identifica el objeto lógico. El servidor resuelve su tipo y firma. Los parámetros definen qué datos recuperar. Conserva la autenticación y el formato de respuesta del contrato desplegado.

## Cambios en local-server

Mantén request.asp como punto de entrada para objetos autorizados de múltiples aplicaciones. Antes de modificarlo, reproduce el caso usando el contrato de la rama desplegada y ubica dónde se introduce el comportamiento específico del proyecto. Declara las consultas de cada reporte en el cliente y deja las reglas del conjunto de datos en el objeto autorizado del backend.

Si falta una capacidad que comparten varios consumidores, aplica el cambio más pequeño al flujo genérico existente y conserva el contrato de los consumidores actuales. Verifica el caso nuevo y al menos una consulta existente de otro proyecto en la misma versión del servidor. En XOver, el manifest selecciona la fuente y la XSLT.

## Contrato y compatibilidad

- `FROM` selecciona el objeto. La copia local también admite `command`.
- `WHERE`, `AND` y `OR` son controles de predicado. En la copia local, `WHERE` contribuye al predicado base y `AND` y `OR` al predicado adicional. Verifica dónde aplica cada uno antes de intercambiarlos.
- `@param` expresa un argumento de rutina. Verifica los nombres, tipos y valores admitidos en la firma publicada.
- `SELECT` expresa la proyección en el diseño propuesto. La copia local revisada usa `fields`. Confirma el alias admitido antes de usar `SELECT` en un cliente.
- Taxan usa `FROM=fin.GetReportAccountBalances&view=income-mxn`. `view` identifica el conjunto de saldos solicitado al objeto. Aunque la clave se llame view, la selección de la plantilla y el tratamiento de los datos corresponden al cliente. El nombre tampoco demuestra una regla general para parámetros sin prefijo.

## Propuesta para parámetros sin prefijo

Clasifica primero las claves reservadas de la versión del servidor. Para cada clave restante, resuelve el significado con la metadata del objeto autorizado. Un nombre como `view` puede vincularse a un argumento declarado por una rutina o convertirse en un filtro de igualdad sobre una columna habilitada en un objeto tabular. Usa `@view` cuando el contrato exija un argumento explícito y `WHERE` cuando el contrato admita un predicado explícito.

Define qué ocurre si se envían `view` y `@view` juntos. Rechaza nombres desconocidos o ambiguos en lugar de descartarlos o aplicar un filtro aproximado. Valida objetos y columnas permitidos, vincula los valores con parámetros SQL y verifica los permisos en el servidor. Trata las cláusulas explícitas como parte de la interfaz autorizada del servidor. Comprueba el comportamiento para una rutina y un objeto tabular antes de declarar la convención compatible.
