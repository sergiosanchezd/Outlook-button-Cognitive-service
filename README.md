# Outlook-button-Cognitive-service
Creación de un botón para outlook que manda el cuerpo de un correo electrónico a "Azure cognitive service text" y analiza el sentimiento 
del correo entre 0% (muy negativo) y 100% (muy positivo).

REQUISTOS:
- Tener creado en Microsoft Azure un servicio Cognitive Service del tipo comprensor de Texto.

PROYECTO
Es un proyecto de tipo "Add-in para Outlook". Este proyecto crea un botón en tu cliente outlook para analizar el cuerpo del correo 
electrónico que se tiene seleccionado. Por tanto, el proyecto tiene 2 acciones:
1. Creación de un botón en la ribbon
2. La acción del botón:
    1. Recoger el cuerpo del correo seleccionado
    2. Enviar el cuerpo al servicio de Azure
    3. Mostrar en porcentaje el sentimiento del correo electrónico, siendo 0 % un correo muy malo y 100 % un correo genial.
