IMPORTANTE: Para que funcione la API, se deben copiar los archivos ubicados en SuperFacturaAPI\libs\ en el directorio c:\windows\system32 o c:\windows\system.
Si VB6 no logra encontrar estos archivos, se mostrará el mensaje "zlib error".

El ejemplo se puede descargar desde:
https://github.com/kripper/superfactura-api-vb6/archive/master.zip

Por favor, probar el ejemplo tal cual antes de modificarlo.
Para abrir el ejemplo, debe abrir el archivo SuperFacturaAPI.vbg desde el explorador de Windows.

Si el ejemplo funciona correctamente, puede copiar y pegar el código de ejemplo a su Software.
Recuerde agregar a su proyecto VB6 la referencia a la clase de la API SuperFactura de la misma forma como se realiza en el proyecto VB6 de ejemplo.

Si al ejecutar el ejemplo aparece un mensaje de error indicando que el JSON contiene errores, puede ser porque en algunas versiones de VB6 se ha observado que el texto JSON que aparece en el formulario del ejemplo viene con caracteres extraños al comienzo y al final. Para solucionar, reemplace estos caracteres por " {" y "} " respectivamente, insertando espacios antes y después para que no vuelvan a aparecer los caracteres extraños.
