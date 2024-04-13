import os
import win32com.client as win32
import openpyxl
import datetime

# Rutas de archivo
ruta_lista_correos = r'C:\Python Project\Email\Listado_correos\Listado_jefes.xlsx'
ruta_archivos_usuarios = r'C:\Python Project\Email\Archivos_por_usuario'
ruta_archivo_apoyo = r'C:\Python Project\Email\Archivo_apoyo\DESCRIPCION DE ROLES.xlsx'
ruta_imagen_banner = r'C:\Python Project\Email\Banner.png'

# Leer la lista de distribución desde el archivo Excel
workbook = openpyxl.load_workbook(ruta_lista_correos)
sheet = workbook.active

# Crear una instancia de Outlook
outlook = win32.Dispatch('Outlook.Application')

# Iterar sobre cada fila en la hoja de cálculo
for row in sheet.iter_rows(min_row=2, values_only=True):
    correo = row[0]
    nombre_usuario = row[1]
    
    # Crear el objeto del mensaje
    mail = outlook.CreateItem(0)
    mail.To = correo
    mail.Subject = 'UAM.8: Certificación de accesos usuarios - aplicaciones SOX'

    # Cuerpo del correo electrónico con imagen al inicio
    body = f"""
    <html>
    <body>
    <p><img src="cid:imagen_inicio" alt="Banner"></p>
    <p>Cordial saludo,</p>
    <p>Con el propósito de asegurar el uso adecuado, protección de la información de la compañía y el cumplimiento SOX, se solicita certificar los roles asignados en cada aplicación de los miembros de su área o equipo de trabajo (Vinculados y/o Terceros a su cargo). Se requiere certificar los roles que tienen las personas y confirmar si deben o no continuar con estos. El archivo DESCRIPCION DE ROLES es un archivo de ayuda que contiene la definición de cada rol de las aplicaciones a certificar. El archivo con su NOMBRE contiene el listado de usuarios que usted debe recertificar, para esto se debe responder en la columna H (DEBE CONTINUAR):</p>
    <ul>
    <li>SI: el usuario continua con los roles que tiene.</li>
    <li>NO: el usuario no continua con el rol que tiene. En este caso, desde la Dirección de Seguridad de la información, procederemos a gestionar el retiro de estos roles.</li>
    <li>NO PERTENECE: el usuario no hace parte de su equipo de trabajo. En caso de conocer al usuario, en la columna (OBSERVACIONES) puede indicarnos a quien puede ser redireccionado.</li>
    </ul>
    <p>Por favor responder este correo adjuntando el mismo archivo enviado el cual tiene su NOMBRE. El plazo máximo para enviar su respuesta son 5 días hábiles, a partir de la fecha de envío de este correo (06/03/2024).</p>
    <p>Recuerde:</p>
    <ol>
    <li>El correo contiene 2 archivos de Excel, un primer archivo de ayuda con la descripción de cada rol de las aplicaciones y un segundo archivo con el listado de los usuarios a su cargo que deben ser recertificados.</li>
    <li>Si tiene dudas sobre el significado de un rol de sus colaboradores, puede revisar el archivo de ayuda DESCRIPCION ROLES que se adjunta al correo.</li>
    <li>El listado de usuarios a recertificar contiene usuarios vinculados a Tigo y usuarios terceros.</li>
    <li>Si dentro del listado no identifica a un usuario, confirme con su equipo de trabajo, ya que puede obedecer a un tercero a cargo de alguno de sus colaboradores.</li>
    <li>En caso de que las personas en el listado de usuarios a recertificar no hagan parte de su equipo de trabajo o área, por favor responda en el mismo listado de usuarios en la columna H (NO PERTENECE). No deben quedar usuarios sin respuesta.</li>
    <li>En caso de que en el listado de accesos y usuarios que debes certificar, se identifica usted mismo, NO se autocertifique, devuelva el correo inmediatamente informando esta situación.</li>
    <li>El formato que contiene el listado de usuarios a recertificar no debe ser alterado, no deben eliminarse, ni agregarse filas y/o columnas. Todas las respuestas deben estar en el formato de Excel en la columna H. Información por fuera de esta columna no será tenida en cuenta.</li>
    <li>Si no se obtiene respuesta a este correo en el plazo indicado, los roles de las personas serán eliminados y se deberá hacer una nueva solicitud de accesos a través del Proceso Gestión de Accesos estándar de la compañía. Los perfiles que se retiren no se pueden restaurar sin una nueva solicitud.</li>
    </ol>
    <p>Si tiene dudas o inquietudes y necesita apoyo, estamos atentos a atenderle en el correo sox.itgc@tigo.com.co.</p>
    <p>NOTA: se aclara que los correos serán enviados en el trascurso del día considerándose día hábil en su horario de 7:00 AM – 6:00 PM de lunes a viernes. En caso que les llegue esté correo por fuera de estos horarios se considerará valido en los horarios indicados anteriormente.</p>
    <p>¡Contamos con Tigo!</p>
    </body>
    </html>
    """
    mail.HTMLBody = body

    # Adjuntar el archivo Excel del usuario
    archivo_excel = os.path.join(ruta_archivos_usuarios, f'{nombre_usuario}.xlsx')
    mail.Attachments.Add(archivo_excel)

    # Adjuntar el archivo 'Explicación'
    mail.Attachments.Add(ruta_archivo_apoyo)

    # Enviar el correo electrónico de inmediato
    mail.Send()

    print("Correo enviado a", correo)
