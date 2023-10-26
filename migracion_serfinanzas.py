#------------------------------------- Importar las libreria nesesarias para el proceso -------------------------------------#
import pandas as pd
import mysql.connector
import openpyxl
import win32com.client as win32
from pathlib import Path
from datetime import datetime

# ------------------------------------- Conexiones a los Servidores para la Migracion y Obtencion de los Datos -------------------------------------#
# Conexión a la base de datos del servidor1
cnx1 = mysql.connector.connect(user='wfm_rpa', password='wfm_rpa2022',  host='172.70.7.96', port='3306',
                              database='dbp_what_serfinanza',              auth_plugin='mysql_native_password')
cursor1 = cnx1.cursor()                              


# Obtener los registros del servidor 1
query="SET GROUP_CONCAT_MAX_LEN = 8048;"
cursor1.execute(query)
cnx1.commit()

query = ("""
         WITH temporal AS (SELECT
                                    sal.PKSA_CODIGO,
                                    sal.SA_CUENTA,
                                    sal.SA_DOCUMENTO,
                                    sal.SA_NOMBRES,
                                    sal.SA_PAGO_MINIMO_MES,
                                    sal.SA_PAGO_TOTAL,
                                    sal.SA_SALDO_CAPITAL,
                                    sal.SA_FEC_ULT_PAGO,
                                    sal.SA_VAL_ULT_PAGO,
                                    sal.SA_TIPO_DE_PRODUCTO,
                                    sal.SA_FECHA_DE_PROMESA,
                                    sal.SA_FECHA_REGISTRO,
                                    sal.SA_FECHA_MODIFICACION,
                                    sal.SA_ESTADO,
                                    cht.GES_NUMERO_COMUNICA,
                                    cht.GES_CDETALLE_ADICIONAL AS Canal,
                                    cht.PKGES_CODIGO AS PKGES_CODIGO,
                                    ty.TYP_OBSERVACIONES,
                                    usu.USU_CNOMBRE

                                FROM
                                    dbp_what_serfinanza.tbl_saldos AS sal
                                        left join dbp_what_serfinanza.tbl_chats_management AS cht ON FKGES_SA_CODIGO = sal.PKSA_CODIGO
                                        left join dbp_what_serfinanza.tbl_typifications AS ty ON ty.FKTYP_NGES_CODIGO = cht.PKGES_CODIGO
                                        left join dbp_what_serfinanza.tbl_usuarios as usu on cht.FKGES_NUSU_CODIGO = PKUSU_NCODIGO
                                where date(sal.SA_FECHA_REGISTRO) =  curdate() 
                                group by PKGES_CODIGO,sal.PKSA_CODIGO,ty.TYP_OBSERVACIONES
                                
                            )
                            SELECT

                                temporal.PKSA_CODIGO,
                                temporal.SA_CUENTA,
                                temporal.SA_DOCUMENTO,
                                temporal.SA_NOMBRES,
                                temporal.SA_PAGO_MINIMO_MES,
                                temporal.SA_PAGO_TOTAL,
                                temporal.SA_SALDO_CAPITAL,
                                temporal.SA_FEC_ULT_PAGO,
                                temporal.SA_VAL_ULT_PAGO,
                                temporal.SA_TIPO_DE_PRODUCTO,
                                temporal.SA_FECHA_DE_PROMESA,
                                temporal.SA_FECHA_REGISTRO,
                                temporal.SA_FECHA_MODIFICACION,
                                temporal.SA_ESTADO,
                                temporal.GES_NUMERO_COMUNICA,
                                temporal.Canal,
                                temporal.PKGES_CODIGO,
                                temporal.TYP_OBSERVACIONES,
                                REPLACE(GROUP_CONCAT(mgs.MES_BODY ORDER BY mgs.PK_MES_NCODE ASC SEPARATOR '\n'), '\r', '') AS historial,
                                MAX(mgs.MES_CREATION_DATE) AS Ultima_interacion,
                                temporal.USU_CNOMBRE

                            FROM
                                temporal
                                    LEFT JOIN dbp_what_serfinanza.tbl_messages AS mgs ON mgs.FK_GES_CODIGO = temporal.PKGES_CODIGO


                            GROUP BY
                                temporal.PKSA_CODIGO,
                                temporal.SA_CUENTA,
                                temporal.SA_DOCUMENTO,
                                temporal.SA_NOMBRES,
                                temporal.SA_PAGO_MINIMO_MES,
                                temporal.SA_PAGO_TOTAL,
                                temporal.SA_SALDO_CAPITAL,
                                temporal.SA_FEC_ULT_PAGO,
                                temporal.SA_VAL_ULT_PAGO,
                                temporal.SA_TIPO_DE_PRODUCTO,
                                temporal.SA_ESTADO,
                                temporal.GES_NUMERO_COMUNICA,
                                temporal.Canal,
                                temporal.PKGES_CODIGO,
                                temporal.TYP_OBSERVACIONES;""")
cursor1.execute(query)
result = cursor1.fetchall() # Obtener todos los registros devueltos por la consulta
fecha_actual= datetime.now().strftime("%Y/%m/%d")
if not result:
    # Envio del Correo al destinatarios
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    # Correo destinatarios
    mail.To = "william.Cuellar@contactosolutions.com; sandra.cardenas@contactosolutions.com; ricardo.arevalo@contactosolutions.com ; andres.cruz@contactosolutions.com"

    # Copia (CC)
    # mail.CC = "leon.gomez@groupcos.com.co"  # Agrega la dirección de correo para copia

    # Asunto del correo
    mail.Subject = 'Plantilla Cargue Serfinanza'

    # Cuerpo del correo
    mail.HTMLBody = f'''<html><body><p>Cordial Saludo,</p>
                    <p>Lainformación Plantilla Cargue Serfinanza al día actual {fecha_actual}, no hay informacion.</p>
                    <p>Quedo atento, cualquier inquietud o solicitud.</p>
                    </body></html>'''
    # Enviar correo
    mail.Send()
else:
    columns = [col[0] for col in cursor1.description]  # Obtener los nombres de las columnas
    df = pd.DataFrame(result, columns=columns)
    # Exportar el DataFrame a un archivo Excel
    ruta_archivo = 'C:\\Users\\Juan.Aya\\Documents\\Proyectos\\Ser Finanzas\\Resultado\\Plantilla_Serfinanzas.xlsx'
    df.to_excel(ruta_archivo,sheet_name='Datos',index=False)

    # Cargar el archivo Excel utilizando openpyxl
    wb = openpyxl.load_workbook(ruta_archivo)
    hoja = wb.active

    # Establecer el formato de tabla
    hoja.auto_filter.ref = hoja.dimensions
    hoja.title = 'Datos'  # Cambiar el nombre de la hoja si es necesario

    # Ajustar el ancho de las columnas
    for columna in hoja.columns:
        max_length = 0
        columna = [cell for cell in columna]
        for cell in columna:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        hoja.column_dimensions[columna[0].column_letter].width = adjusted_width

    # Guardar el archivo Excel con los cambios realizados
    wb.save(ruta_archivo)

    # Envio del Correo al destinatarios
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    # Correo destinatarios
    mail.To = "william.Cuellar@contactosolutions.com; sandra.cardenas@contactosolutions.com; ricardo.arevalo@contactosolutions.com ; andres.cruz@contactosolutions.com"

    # Copia (CC)
    # mail.CC = "leon.gomez@groupcos.com.co"  # Agrega la dirección de correo para copia

    # Asunto del correo
    mail.Subject = 'Plantilla Cargue Serfinanza'

    # Cuerpo del correo
    mail.HTMLBody = f'''<html><body><p>Cordial Saludo,</p>
                    <p>Hago entrega de la Plantilla Cargue Serfinanza al día actual {fecha_actual}.</p>
                    <p>Quedo atento, cualquier inquietud o solicitud.</p>
                    </body></html>'''

    # Ruta Archivo Excel adjuntar al correo
    ruta_archivo = r'C:\\Users\\Juan.Aya\\Documents\\Proyectos\\Ser Finanzas\\Resultado\\Plantilla_Serfinanzas.xlsx'

    # Obtener el nombre del archivo y la extensión
    nombre_archivo = Path(ruta_archivo).name

    # Adjuntar el archivo
    attachment = mail.Attachments.Add(ruta_archivo)
    attachment.DisplayName = nombre_archivo

    # Enviar correo
    mail.Send()

    print('finalizo')