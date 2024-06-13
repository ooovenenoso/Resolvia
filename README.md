# Gestión de Tickets de Mantenimiento: RESOLVIA

Este script de Google Apps Script automatiza la gestión de tickets de mantenimiento para una organización educativa. Captura respuestas del formulario, asigna números de ticket únicos y envía notificaciones por correo electrónico a los reportantes y responsables. También actualiza el estado de los tickets en tiempo real, asegurando una gestión eficiente y organizada de las solicitudes de mantenimiento.

## Funciones

### onFormSubmit(e)
- Captura las respuestas del formulario.
- Asigna un número de ticket único.
- Envía una notificación de confirmación por correo electrónico al reportante.

### onEdit(e)
- Monitorea las ediciones en la hoja de respuestas del formulario.
- Envía notificaciones por correo electrónico cuando se asignan tareas o se actualiza el estado de un ticket.

## Uso

1. Copia el script en el editor de scripts de Google Sheets.
2. Configura los activadores para `onFormSubmit` y `onEdit`:
   - `onFormSubmit` se activa al enviar el formulario.
   - `onEdit` se activa al editar la hoja de respuestas del formulario.
3. Personaliza el mapa de correos electrónicos (`emailMap`) con las direcciones de correo de los responsables.

## Ejemplo de Configuración

```javascript
var emailMap = {
  "Persona 1": "persona1@example.com",
  "Persona 2": "persona2@example.com",
  "Persona 3": "persona3@example.com",
  "Persona 4": "persona4@example.com"
  // Agregar más entradas según sea necesario
};
```


## Licencia

Este proyecto está bajo la Licencia MIT. Consulta el archivo `LICENSE` para más detalles.
