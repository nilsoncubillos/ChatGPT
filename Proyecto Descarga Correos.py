import os
import win32com.client

# Crear una instancia de Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Buzón específico
nombre_buzon = "SOX.ITGC@tigo.com"

# Ruta donde se guardarán los correos descargados
carpeta_destino = r'C:\Python Project\Email\Correos Descargados'

# Crear la carpeta de destino si no existe
if not os.path.exists(carpeta_destino):
    os.makedirs(carpeta_destino)

# Obtener el buzón específico
buzon = outlook.Folders.Item(nombre_buzon)

# Obtener la carpeta de la bandeja de entrada del buzón específico
bandeja_entrada = buzon.GetDefaultFolder(6)  # 6 representa la carpeta de la bandeja de entrada

# Iterar a través de todos los correos en la bandeja de entrada del buzón específico
for correo in bandeja_entrada.Items:
    # Verificar si el correo comienza con "UAM.8"
    if correo.Subject.startswith("UAM.8"):
        # Obtener la dirección de correo electrónico del destinatario
        destinatario = correo.To if correo.To else correo.CC if correo.CC else ""
        if destinatario:
            # Limpiar la dirección de correo electrónico para usarla como nombre de archivo
            nombre_archivo = destinatario.split("@")[0].replace(".", "_") + ".msg"
            # Guardar el correo como archivo .msg en la carpeta de destino
            ruta_archivo = os.path.join(carpeta_destino, nombre_archivo)
            correo.SaveAs(ruta_archivo)
            print(f"Correo guardado como '{nombre_archivo}'")

print("Proceso completado.")
