Automatización de la generación de la plantilla de carga para Serfinanza

Este proyecto automatiza la generación de la plantilla de carga para Serfinanza. El proceso se realiza de la siguiente manera:

Se conecta a la base de datos de Serfinanza y se consultan los registros del día actual.
Se crea un DataFrame con los datos obtenidos de la consulta.
Se exporta el DataFrame a un archivo Excel.
Se envía un correo electrónico con el archivo Excel adjunto a los destinatarios especificados.

Requisitos

Python 3.8 o superior
Las bibliotecas pandas, mysql.connector, openpyxl y win32com.client

Instalación

Instala las dependencias:
pip install -r requirements.txt
Uso

Modifica los parámetros de conexión a la base de datos en el archivo config.py.
Ejecuta el script principal:
python main.py
Explicación del código

El código se divide en las siguientes partes:

Importación de librerías
Conexión a la base de datos
Consulta de los registros
Creación del DataFrame
Exportación del DataFrame
Envío del correo electrónico
Conexión a la base de datos

En esta sección se conecta a la base de datos de Serfinanza utilizando la biblioteca mysql.connector.

Consulta de los registros

En esta sección se consultan los registros del día actual utilizando una consulta SQL.

Creación del DataFrame

En esta sección se crea un DataFrame con los datos obtenidos de la consulta.

Exportación del DataFrame

En esta sección se exporta el DataFrame a un archivo Excel utilizando la biblioteca openpyxl.

Envío del correo electrónico

En esta sección se envía un correo electrónico con el archivo Excel adjunto utilizando la biblioteca win32com.client.