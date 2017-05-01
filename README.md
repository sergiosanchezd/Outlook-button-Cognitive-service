# Outlook-button-Cognitive-service
Creación de una acción cuando se selecciona un correo que incluye en la columna "Analyze" el porcentaje del sentimiento entre 0% (muy negativo) y 100% (muy positivo) sacado de "Azure cognitive service text".

REQUISTOS:
- Tener creado en Microsoft Azure un servicio Cognitive Service del tipo comprensor de Texto.
- Crear en outlook una columna que se llame "Analyze" y de tipo texto. Añadir a la vista o vistas que quieras verlo.

PROYECTO:
Es un proyecto de tipo "Add-in para Outlook". Este proyecto crea una acción cuando seleccionas un correo que lo analiza y mete el resultado en la columna "Analyze".

Por tanto, lo mejor es que incluyas en todas las vistas primero una columna que se llame "Analyze" de tipo texto y que puedas ver los resultados.

Además, debes cambiar en Addin.cs el API Key de tu servicio de azure que analiza el texto.
