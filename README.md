📧 Sistema de Envío Masivo de Correos Personalizados
📖 Descripción
Sistema profesional para envío masivo de correos electrónicos personalizados con interfaz gráfica moderna. Desarrollado en Python con TKinter, permite enviar correos masivos utilizando plantillas personalizables y una base de datos en Excel.
Optimizado para correos @gmx.com

Autor: ERICK ANDRE MARTINEZ DE ANDA
Versión: 1.6

🚀 Características Principales
✨ Funcionalidades
Interfaz gráfica moderna con 4 pestañas organizadas
Personalización dinámica de mensajes con variables
Modo pruebas con pausas reducidas (2 segundos)
Gestión de archivos adjuntos opcional
Pausas anti-spam inteligentes y configurables
Validación completa de configuración antes del envío
Log de actividad en tiempo real con timestamp
Vista previa de datos del Excel
Barra de progreso durante el envío
Control de envío (iniciar/detener)


🛡️ Características de Seguridad
Conexión SMTP SSL segura con GMX
Pausas estratégicas entre envíos
Validación de credenciales y archivos
Manejo robusto de errores


📋 Requisitos del Sistema
🔧 Dependencias
python
pandas >= 1.3.0
tkinter (incluido en Python)
smtplib (incluido en Python)
email (incluido en Python)
💻 Compatibilidad
Sistema Operativo: Windows

Python: Versión 3.7 o superior

Archivos: Soporte para Excel (.xlsx, .xls) y cualquier tipo de archivo adjunto

📊 FORMATO DEL ARCHIVO EXCEL
🎯 Estructura Requerida
El sistema es FLEXIBLE y detecta automáticamente las columnas sin importar el orden.

📝 Columnas Reconocidas para Correos
El sistema busca automáticamente estas columnas (en cualquier orden):

email, correo, e-mail, mail, Email, Correo

🔤 Variables Personalizables
Puedes usar CUALQUIER columna como variable en tus mensajes:


python
# En el asunto:
"Contacto para {empresa} - {nombre}"

# En el cuerpo:
"Estimado {nombre} de {empresa}, nos comunicamos sobre {asunto}..."

📋 Ejemplos de Estructuras Válidas
Ejemplo 1: Formato Básico
csv
email
cliente1@empresa.com
cliente2@empresa.com
cliente3@empresa.com

Ejemplo 2: Formato Completo (Recomendado)
csv
email,empresa,nombre,telefono,puesto,ciudad
cliente1@empresa.com,Tech Solutions,Ana Martínez,555-1234,Gerente,Madrid
cliente2@empresa.com,Innovation Corp,Carlos López,555-5678,Director,Barcelona
cliente3@empresa.com,Services SL,María García,555-9012,Coordinador,Valencia
Ejemplo 3: Columnas Mezcladas

csv
telefono|nombre|empresa|correo|departamento
555-1111|Juan Pérez|Empresa SA|juan@empresa.com|Ventas
555-2222|Laura Gómez|Comercial SL|laura@comercial.com|Marketing

Ejemplo 4: Columnas en Inglés
csv
company,name,phone,email,position,industry
Company A|John Doe|333-1111|john@companya.com|Manager|Technology
Company B|Jane Smith|333-2222|jane@companyb.com|Director|Finance


🎨 Casos de Uso con Variables
Para Envío de CVs:
csv
email,empresa,contacto,puesto,vacante,ciudad
rh@tech.com,Tech Solutions,Ana Martínez,Reclutador,Desarrollador,Madrid
empleos@empresa.com,Empresa XYZ,Carlos López,Gerente RH,Diseñador,Barcelona

Plantilla:
text
Asunto: Aplicación para {vacante} en {empresa}

Estimado(a) {contacto},
Me interesa la posición de {vacante} en {empresa}...

Para Comunicación Comercial:
csv
correo,empresa,contacto,industria,producto_interes
ventas@cliente.com,Distribuidora SA,Luis García,Retail,Software ERP
gerente@otra.com,Otra Empresa,María Rodríguez,Manufactura,Consultoría

Plantilla:
text
Asunto: Propuesta de {producto_interes} para {empresa}

Estimado {contacto},
Tenemos una solución ideal para {industria}...


⚠️ Consideraciones Importantes
Orden flexible: Las columnas pueden estar en cualquier orden
Case-insensitive: empresa = EMPRESA = Empresa
Columnas opcionales: No es necesario que todas las filas tengan todos los datos
Detección automática: El sistema encuentra la columna de email automáticamente
Caracteres especiales: Usar encoding UTF-8 para caracteres especiales


❌ Estructuras NO Válidas
csv
# SIN encabezados de columna
cliente1@empresa.com
cliente2@empresa.com


# Columnas completamente vacías
email,empresa,nombre
cliente1@empresa.com,,Juan
cliente2@empresa.com,,

🖥️ Instalación y Configuración
1. 📥 Instalación de Dependencias
bash
pip install pandas openpyxl

3. ⚙️ Configuración de Cuenta GMX
Necesitas una cuenta en GMX.com

Habilitar el acceso de aplicaciones menos seguras (si es necesario)

Usar el servidor SMTP: mail.gmx.com puerto 465

3. 🗂️ Preparación de Archivos
Colocar el archivo Excel en la misma carpeta o especificar ruta

Preparar archivo adjunto (si se requiere)

Verificar que el Excel tenga al menos una columna con emails

🎮 Guía de Uso
1. 🏁 Inicio
Ejecutar el script:

bash
python enviar_correos.py
2. 📧 Configuración de Correo (Pestaña 1)
Ingresar correo remitente GMX

Ingresar contraseña

Seleccionar si adjuntar archivo (opcional)

3. ✉️ Personalización de Mensaje (Pestaña 2)
Asunto: Usar variables como {empresa}, {nombre}

Cuerpo: Plantilla personalizable con variables

Ejemplo:

text
Asunto: Contacto de {nombre} para {empresa}

Cuerpo:
Estimados de {empresa},
Mi nombre es {nombre} y me interesa...
4. 📊 Base de Datos (Pestaña 3)
Seleccionar archivo Excel

Ver vista previa de datos

Validar que se detecten las columnas correctamente

5. 🚀 Envío Masivo (Pestaña 4)
Generar Resumen: Verificar configuración

Activar Modo Pruebas (recomendado para pruebas)

Iniciar Envío: Comenzar proceso masivo

Monitorizar progreso en barra y log


🔧 Modo Pruebas
✅ Características del Modo Pruebas
Pausas reducidas: 2 segundos vs 60-180 segundos normales

Feedback inmediato: Log detallado del proceso

Ideal para testing: Verificar configuración con pocos correos


🎯 Cómo Usar Modo Pruebas
Activar checkbox "🔧 MODO PRUEBAS"

Usar Excel con 1-5 correos de prueba

Iniciar envío y verificar resultados

Revisar log para detectar errores

🐛 Solución de Problemas
❌ Error: "No se encontró correo destino"
Verificar que el Excel tenga columnas de email reconocidas

Revisar nombres de columnas en la vista previa

Asegurar que no haya filas vacías

❌ Error: "Error al enviar correo"
Verificar credenciales GMX

Confirmar conexión a internet

Revisar firewall/antivirus

❌ Error: "Archivo Excel no encontrado"
Verificar ruta del archivo

Confirmar que el archivo no esté abierto en Excel

Revisar permisos de lectura

❌ Error: "Timeout SMTP"
Verificar configuración de GMX

Revisar límites de envío de GMX

Esperar y reintentar




📈 Mejores Prácticas
✅ Para Excel
Usar encabezados de columna descriptivos

Mantener datos limpios y consistentes

Usar formato UTF-8 para caracteres especiales

Verificar que los emails sean válidos

✅ Para Mensajes
Personalizar asunto para cada destinatario

Usar variables relevantes en el cuerpo

Mantener formato profesional

Probar siempre en modo pruebas primero



✅ Para Envíos Masivos
Comenzar con modo pruebas activado

Verificar log después de cada envío

Respetar pausas anti-spam en producción

Monitorear límites de envío del proveedor




🔄 Flujo de Trabajo Recomendado
Preparación: Configurar Excel y plantillas

Prueba: Enviar a 1-2 correos con modo pruebas

Validación: Revisar log y correos recibidos

Ajuste: Corregir problemas identificados

Producción: Desactivar modo pruebas y enviar masivo

Seguimiento: Monitorear log durante el envío

