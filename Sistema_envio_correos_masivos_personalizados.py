"""
SISTEMA DE ENVÍO MASIVO DE CORREOS PERSONALIZADOS
Sistema modular para envío masivo de correos con personalización y gestión de archivos adjuntos.
Autor: ERICK ANDRE MARTINEZ DE ANDA
Version 1.6 - CON INTERFAZ GRÁFICA TKINTER Y MODO PRUEBAS

Este sistema permite el envío masivo de correos electrónicos personalizados utilizando
una base de datos en Excel, con interfaz gráfica
como pausas anti-spam y variables dinámicas.
"""

import pandas as pd
import smtplib
import os
from email.message import EmailMessage
import time
import random
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from datetime import datetime


# =============================================================================
# CLASE: ConfiguradorPausas
# =============================================================================
class ConfiguradorPausas:
    """
    Gestiona la configuración de pausas con modos normal y pruebas.
    """
    
    def __init__(self):
        """Inicializa el configurador con modos de pausa."""
        self.modo_pruebas = False
        self.min_segundos = 60
        self.max_segundos = 180
        self.pausa_pruebas = 2  # Solo 2 segundos en modo pruebas
    
    def set_modo_pruebas(self, activar):
        """
        Activa o desactiva el modo pruebas.
        
        Args:
            activar (bool): True para activar modo pruebas, False para modo normal
        """
        self.modo_pruebas = activar
        
    def obtener_tiempo_espera(self):
        """
        Retorna el tiempo de espera según el modo actual.
        
        Returns:
            int: Tiempo de espera en segundos
        """
        if self.modo_pruebas:
            return self.pausa_pruebas
        else:
            return random.randint(self.min_segundos, self.max_segundos)


# =============================================================================
# CLASE: ManejadorCorreo
# =============================================================================
class ManejadorCorreo:
    """
    Gestiona el envío de correos electrónicos a través del servidor SMTP de GMX.
    """
    
    def __init__(self, remitente, clave, archivo_adjunto="", adjuntar_archivo=False):
        """
        Inicializa el manejador de correo con las credenciales y configuración.
        """
        self.remitente = remitente
        self.clave = clave
        self.archivo_adjunto = archivo_adjunto
        self.adjuntar_archivo = adjuntar_archivo

    def enviar_correo(self, destinatario, asunto, cuerpo, variables, interfaz):
        """
        Envía un correo electrónico individual a través del servidor SMTP de GMX.
        """
        try:
            # Crear objeto de mensaje de email
            mensaje = EmailMessage()
            mensaje["From"] = self.remitente
            mensaje["To"] = destinatario
            mensaje["Subject"] = asunto
            mensaje.set_content(cuerpo)

            # Adjuntar archivo si está configurado y el archivo existe
            if self.adjuntar_archivo and self.archivo_adjunto and os.path.isfile(self.archivo_adjunto):
                with open(self.archivo_adjunto, "rb") as f:
                    data = f.read()
                    nombre = os.path.basename(self.archivo_adjunto)
                    mensaje.add_attachment(
                        data,
                        maintype="application",
                        subtype="octet-stream",
                        filename=nombre
                    )

            # Envío a través de GMX
            with smtplib.SMTP_SSL("mail.gmx.com", 465) as smtp:
                smtp.login(self.remitente, self.clave)
                smtp.send_message(mensaje)
                
                # Mostrar información del envío
                empresa = variables.get('empresa', 'N/A')
                nombre = variables.get('nombre', 'N/A')
                interfaz.log(f"✅ Correo enviado a {destinatario} | Empresa: {empresa} | Nombre: {nombre}")

        except Exception as e:
            interfaz.log(f"❌ Error al enviar correo a {destinatario}: {e}")


# =============================================================================
# CLASE: PersonalizadorMensaje
# =============================================================================
class PersonalizadorMensaje:
    """
    Gestiona la personalización dinámica de asuntos y cuerpos de correo.
    """
    
    def __init__(self):
        """Inicializa el personalizador con formatos vacíos."""
        self.formato_asunto = ""
        self.formato_cuerpo = ""
    
    def generar_mensaje(self, **variables):
        """
        Genera el asunto y cuerpo del mensaje aplicando las variables.
        """
        asunto = self.formato_asunto
        cuerpo = self.formato_cuerpo
        
        # Reemplazar cada variable en el asunto y cuerpo
        for key, value in variables.items():
            placeholder = f"{{{key}}}"
            asunto = asunto.replace(placeholder, str(value))
            cuerpo = cuerpo.replace(placeholder, str(value))
        
        return asunto, cuerpo


# =============================================================================
# CLASE: ManejadorPausas
# =============================================================================
class ManejadorPausas:
    """
    Gestiona pausas estratégicas entre envíos para evitar detección como spam.
    """
    
    def __init__(self, configurador_pausas):
        """
        Inicializa el manejador de pausas con el configurador.
        """
        self.configurador = configurador_pausas
    
    def pausa_estrategica(self, correo_actual, total_correos, interfaz):
        """
        Ejecuta una pausa aleatoria entre envíos con posibilidad de cancelación.
        """
        if interfaz.enviando and correo_actual < total_correos:
            # Obtener tiempo de espera según el modo
            espera = self.configurador.obtener_tiempo_espera()
            
            modo = "PRUEBAS" if self.configurador.modo_pruebas else "PRODUCCIÓN"
            interfaz.log(f"⏰ Pausa {modo}: {espera} segundos | Progreso: {correo_actual}/{total_correos}")
            
            # Pausa con verificación periódica para permitir cancelación
            for i in range(espera):
                if not interfaz.enviando:
                    break
                time.sleep(1)
                # Actualizar contador de pausa cada segundo
                if i % 5 == 0:  # Actualizar cada 5 segundos para no saturar
                    segundos_restantes = espera - i
                    interfaz.actualizar_estado_pausa(segundos_restantes)


# =============================================================================
# CLASE: ProcesadorExcel
# =============================================================================
class ProcesadorExcel:
    """
    Procesa archivos Excel y extrae información de contactos.
    """
    
    def __init__(self, ruta_excel):
        """
        Inicializa el procesador con la ruta del archivo Excel.
        """
        self.ruta_excel = ruta_excel
        self.dataframe = None
        self.columnas = []
        
    def cargar_datos(self):
        """
        Carga y valida los datos del archivo Excel.
        """
        try:
            self.dataframe = pd.read_excel(self.ruta_excel)
            self.columnas = self.dataframe.columns.tolist()
            return True
        except FileNotFoundError:
            raise FileNotFoundError(f"No se encontró el archivo Excel: {self.ruta_excel}")
        except Exception as e:
            raise Exception(f"Error al cargar el Excel: {str(e)}")
    
    def obtener_correo_destino(self, fila):
        """
        Busca y retorna el correo electrónico en una fila de datos.
        """
        columnas_posibles = ['email', 'correo', 'e-mail', 'mail', 'Email', 'Correo']
        
        for columna in columnas_posibles:
            if columna in fila and pd.notna(fila[columna]):
                return fila[columna]
        
        return None
    
    def obtener_total_filas(self):
        """
        Retorna el número total de filas (contactos) en el Excel.
        """
        return len(self.dataframe) if self.dataframe is not None else 0
    
    def iterar_filas(self):
        """
        Generador para iterar sobre todas las filas del DataFrame.
        """
        if self.dataframe is None:
            raise Exception("No hay datos cargados. Ejecute cargar_datos() primero.")
        
        for index, row in self.dataframe.iterrows():
            yield index, row


# =============================================================================
# CLASE: ManejadorBaseDatos
# =============================================================================
class ManejadorBaseDatos:
    """
    Coordina el proceso de envío masivo utilizando todos los componentes.
    """
    
    def __init__(self, ruta_excel, correo_obj, personalizador, manejador_pausas):
        """
        Inicializa el manejador de base de datos con todos los componentes necesarios.
        """
        self.ruta_excel = ruta_excel
        self.correo_obj = correo_obj
        self.personalizador = personalizador
        self.manejador_pausas = manejador_pausas
        self.contador = 0
        self.procesador_excel = ProcesadorExcel(ruta_excel)
    
    def enviar_todos(self, interfaz):
        """
        Ejecuta el proceso completo de envío masivo de correos.
        """
        try:
            # Cargar datos del Excel
            self.procesador_excel.cargar_datos()
            total_correos = self.procesador_excel.obtener_total_filas()
            interfaz.total_correos = total_correos
            
            interfaz.log(f"📤 INICIANDO ENVÍO DE {total_correos} CORREOS")
            interfaz.log("🔄 Procesando...")
            
            # Iterar sobre cada fila/contacto del Excel
            for index, fila in self.procesador_excel.iterar_filas():
                # Verificar si el usuario canceló el envío
                if not interfaz.enviando:
                    break
                    
                self.contador += 1
                
                # Convertir la fila a diccionario para las variables
                variables = fila.to_dict()
                
                # Generar mensaje personalizado usando las variables
                asunto, cuerpo = self.personalizador.generar_mensaje(**variables)
                
                # Obtener el correo del destinatario
                correo_destino = self.procesador_excel.obtener_correo_destino(fila)
                
                if correo_destino:
                    # Mostrar preparación de envío
                    interfaz.log(f"📝 Preparando correo {self.contador}/{total_correos} para {correo_destino}")
                    
                    # Enviar correo individual
                    self.correo_obj.enviar_correo(
                        destinatario=correo_destino,
                        asunto=asunto,
                        cuerpo=cuerpo,
                        variables=variables,
                        interfaz=interfaz
                    )
                    
                    # Actualizar barra de progreso en la interfaz
                    interfaz.actualizar_progreso(self.contador, total_correos)
                    
                    # Ejecutar pausa estratégica entre envíos
                    if self.contador < total_correos:
                        self.manejador_pausas.pausa_estrategica(self.contador, total_correos, interfaz)
                else:
                    # Log de advertencia si no se encuentra correo
                    interfaz.log(f"❌ No se encontró correo destino en la fila {self.contador}")
            
            # Mensaje final según el estado del envío
            if interfaz.enviando:
                interfaz.log(f"✅ ENVÍO COMPLETADO: {self.contador}/{total_correos} correos enviados")
                messagebox.showinfo("Éxito", f"Envio completado: {self.contador}/{total_correos} correos enviados")
            else:
                interfaz.log(f"⏹️ ENVÍO INTERRUMPIDO: {self.contador}/{total_correos} correos enviados")
                
        except Exception as e:
            # Manejo de errores generales
            interfaz.log(f"❌ Error inesperado: {e}")
            messagebox.showerror("Error", f"Error en el proceso: {str(e)}")


# =============================================================================
# CLASE: ValidadorConfiguracion
# =============================================================================
class ValidadorConfiguracion:
    """
    Valida la configuración del sistema antes del envío masivo.
    """
    
    def __init__(self, remitente, clave, ruta_excel, asunto, cuerpo):
        """
        Inicializa el validador con los datos de configuración.
        """
        self.remitente = remitente
        self.clave = clave
        self.ruta_excel = ruta_excel
        self.asunto = asunto
        self.cuerpo = cuerpo
    
    def validar_completo(self):
        """
        Ejecuta todas las validaciones de configuración.
        """
        validaciones = [
            self.validar_remitente(),
            self.validar_clave(),
            self.validar_excel(),
            self.validar_asunto(),
            self.validar_cuerpo()
        ]
        
        for valido, mensaje in validaciones:
            if not valido:
                return False, mensaje
                
        return True, "Configuración válida"
    
    def validar_remitente(self):
        """Valida que el remitente esté presente y tenga formato de email básico."""
        if not self.remitente.strip():
            return False, "Por favor ingrese el correo remitente"
        if '@' not in self.remitente:
            return False, "El correo remitente no tiene formato válido"
        return True, ""
    
    def validar_clave(self):
        """Valida que la contraseña esté presente."""
        if not self.clave.strip():
            return False, "Por favor ingrese la contraseña"
        return True, ""
    
    def validar_excel(self):
        """Valida que el archivo Excel exista y sea accesible."""
        if not self.ruta_excel.strip():
            return False, "Por favor seleccione un archivo Excel"
        if not os.path.exists(self.ruta_excel):
            return False, f"El archivo Excel no existe: {self.ruta_excel}"
        return True, ""
    
    def validar_asunto(self):
        """Valida que el asunto esté presente."""
        if not self.asunto.strip():
            return False, "Por favor ingrese el asunto del correo"
        return True, ""
    
    def validar_cuerpo(self):
        """Valida que el cuerpo del mensaje esté presente."""
        if not self.cuerpo.strip():
            return False, "Por favor ingrese el cuerpo del mensaje"
        return True, ""


# =============================================================================
# CLASE: GestorInterfaz
# =============================================================================
class GestorInterfaz:
    """
    Gestiona la interacción entre la lógica de negocio y la interfaz gráfica.
    """
    
    def __init__(self, interfaz_principal):
        """
        Inicializa el gestor con referencia a la interfaz principal.
        """
        self.interfaz = interfaz_principal
        self.enviando = False
        self.proceso_envio = None
        self.configurador_pausas = ConfiguradorPausas()
    
    def iniciar_envio(self):
        """Inicia el proceso de envío masivo en un hilo separado."""
        if self.enviando:
            self.interfaz.log("⚠️ El envío ya está en progreso")
            return
            
        # Validar configuración antes de iniciar
        validador = ValidadorConfiguracion(
            remitente=self.interfaz.entry_remitente.get(),
            clave=self.interfaz.entry_clave.get(),
            ruta_excel=self.interfaz.entry_excel.get(),
            asunto=self.interfaz.entry_asunto.get(),
            cuerpo=self.interfaz.text_cuerpo.get('1.0', tk.END).strip()
        )
        
        valido, mensaje = validador.validar_completo()
        if not valido:
            messagebox.showerror("Error de Validación", mensaje)
            return
        
        # Configurar estado de envío
        self.enviando = True
        self.interfaz.enviando = True
        self.interfaz.progress_var.set(0)
        
        # Configurar modo pruebas si está activado
        modo_pruebas = self.interfaz.modo_pruebas_var.get()
        self.configurador_pausas.set_modo_pruebas(modo_pruebas)
        
        if modo_pruebas:
            self.interfaz.log("🔧 MODO PRUEBAS ACTIVADO - Pausas reducidas a 2 segundos")
        
        # Actualizar interfaz
        self.interfaz.actualizar_estado_botones(envio_activo=True)
        
        # Ejecutar en hilo separado para no bloquear la interfaz
        self.proceso_envio = threading.Thread(target=self._ejecutar_envio)
        self.proceso_envio.daemon = True
        self.proceso_envio.start()
        
        self.interfaz.log("🚀 Iniciando proceso de envío masivo...")
    
    def detener_envio(self):
        """Detiene el proceso de envío masivo."""
        if self.enviando:
            self.enviando = False
            self.interfaz.enviando = False
            self.interfaz.actualizar_estado_botones(envio_activo=False)
            self.interfaz.log("⏹️ Solicitando detención del envío...")
        else:
            self.interfaz.log("ℹ️ No hay envío en progreso")
    
    def _ejecutar_envio(self):
        """Método interno que ejecuta el envío masivo."""
        try:
            # Configurar todos los componentes del sistema
            correo = ManejadorCorreo(
                remitente=self.interfaz.entry_remitente.get(),
                clave=self.interfaz.entry_clave.get(),
                archivo_adjunto=self.interfaz.entry_archivo.get() if self.interfaz.adjuntar_var.get() else "",
                adjuntar_archivo=self.interfaz.adjuntar_var.get()
            )
            
            personalizador = PersonalizadorMensaje()
            personalizador.formato_asunto = self.interfaz.entry_asunto.get()
            personalizador.formato_cuerpo = self.interfaz.text_cuerpo.get('1.0', tk.END).strip()
            
            manejador_pausas = ManejadorPausas(self.configurador_pausas)
            
            base_datos = ManejadorBaseDatos(
                ruta_excel=self.interfaz.entry_excel.get(),
                correo_obj=correo,
                personalizador=personalizador,
                manejador_pausas=manejador_pausas
            )
            
            # Ejecutar envío pasando referencia al gestor para control
            base_datos.enviar_todos(self.interfaz)
            
        except Exception as e:
            self.interfaz.log(f"❌ Error en el envío masivo: {str(e)}")
            messagebox.showerror("Error", f"Error en el envío masivo: {str(e)}")
        finally:
            # Restablecer estado al finalizar
            self.enviando = False
            self.interfaz.enviando = False
            self.interfaz.actualizar_estado_botones(envio_activo=False)


# =============================================================================
# CLASE: InterfazGrafica (MAIN UI)
# =============================================================================
class InterfazGrafica:
    """
    Interfaz gráfica principal del sistema de envío masivo de correos.
    """
    
    def __init__(self):
        """Inicializa la interfaz gráfica principal y todos sus componentes."""
        self.root = tk.Tk()
        self.root.title("Sistema de Envío Masivo de Correos - v1.6")
        self.root.geometry("900x750")
        self.root.configure(bg='#f0f0f0')
        
        # Variables de estado del sistema
        self.enviando = False
        self.progreso = 0
        self.total_correos = 0
        
        # Inicializar gestor de interfaz
        self.gestor = GestorInterfaz(self)
        
        # Configurar y crear la interfaz
        self.configurar_estilo()
        self.crear_interfaz()
        
    def configurar_estilo(self):
        """Configura los estilos visuales de la interfaz."""
        style = ttk.Style()
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
        style.configure('TButton', font=('Arial', 10))
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'))
        style.configure('Section.TLabel', font=('Arial', 12, 'bold'))
        
    def crear_interfaz(self):
        """Crea todos los elementos de la interfaz gráfica."""
        # Título principal
        titulo = ttk.Label(self.root, text="🚀 SISTEMA DE ENVÍO MASIVO DE CORREOS v1.6", style='Title.TLabel')
        titulo.pack(pady=10)
        
        # Frame de controles rápidos
        self.crear_controles_rapidos()
        
        # Notebook (pestañas) principal
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Crear todas las pestañas de configuración
        self.crear_pestana_configuracion(notebook)
        self.crear_pestana_mensaje(notebook)
        self.crear_pestana_base_datos(notebook)
        self.crear_pestana_envio(notebook)
        
        # Área de log de actividad
        self.crear_area_log()
        
    def crear_controles_rapidos(self):
        """Crea la barra de controles rápidos en la parte superior."""
        controles_frame = ttk.Frame(self.root)
        controles_frame.pack(fill='x', padx=20, pady=5)
        
        # Modo pruebas
        self.modo_pruebas_var = tk.BooleanVar()
        ttk.Checkbutton(controles_frame, text="🔧 MODO PRUEBAS (Pausas de 2 segundos)", 
                       variable=self.modo_pruebas_var).pack(side='left', padx=10)
        
        # Estado del sistema
        self.estado_label = ttk.Label(controles_frame, text="🔴 Listo", foreground="red")
        self.estado_label.pack(side='right', padx=10)
        
    def crear_pestana_configuracion(self, notebook):
        """Crea la pestaña de configuración de correo."""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="📧 Configuración Correo")
        
        # Remitente
        ttk.Label(frame, text="Correo Remitente:", style='Section.TLabel').grid(row=0, column=0, sticky='w', padx=10, pady=10)
        self.entry_remitente = ttk.Entry(frame, width=40, font=('Arial', 10))
        self.entry_remitente.grid(row=0, column=1, padx=10, pady=10, sticky='ew')
        
        # Contraseña
        ttk.Label(frame, text="Contraseña:", style='Section.TLabel').grid(row=1, column=0, sticky='w', padx=10, pady=10)
        self.entry_clave = ttk.Entry(frame, width=40, font=('Arial', 10), show='*')
        self.entry_clave.grid(row=1, column=1, padx=10, pady=10, sticky='ew')
        
        # Adjuntar archivo
        self.adjuntar_var = tk.BooleanVar()
        ttk.Checkbutton(frame, text="Adjuntar archivo", variable=self.adjuntar_var, 
                       command=self.toggle_adjuntar).grid(row=2, column=0, sticky='w', padx=10, pady=10)
        
        # Selección de archivo
        self.frame_archivo = ttk.Frame(frame)
        self.frame_archivo.grid(row=3, column=0, columnspan=2, sticky='ew', padx=10, pady=5)
        
        ttk.Label(self.frame_archivo, text="Archivo a adjuntar:").grid(row=0, column=0, sticky='w')
        self.entry_archivo = ttk.Entry(self.frame_archivo, width=30, font=('Arial', 10))
        self.entry_archivo.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(self.frame_archivo, text="Buscar", command=self.buscar_archivo).grid(row=0, column=2, padx=5, pady=5)
        
        # Configurar grid weights
        frame.columnconfigure(1, weight=1)
        self.frame_archivo.columnconfigure(1, weight=1)
        
        # Inicialmente deshabilitado
        self.frame_archivo.grid_remove()
        
    def crear_pestana_mensaje(self, notebook):
        """Crea la pestaña de personalización del mensaje."""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="✉️ Personalización Mensaje")
        
        # Información sobre variables
        info_frame = ttk.LabelFrame(frame, text="Variables Disponibles")
        info_frame.grid(row=0, column=0, columnspan=2, sticky='ew', padx=10, pady=10)
        
        info_text = "Puede usar variables como: {nombre}, {empresa}, {fecha}, {telefono}, etc.\n"
        info_text += "Estas variables se reemplazarán automáticamente con los datos del Excel."
        ttk.Label(info_frame, text=info_text, justify='left').grid(row=0, column=0, sticky='w', padx=10, pady=10)
        
        # Asunto
        ttk.Label(frame, text="Asunto del Correo:", style='Section.TLabel').grid(row=1, column=0, sticky='w', padx=10, pady=10)
        self.entry_asunto = ttk.Entry(frame, width=60, font=('Arial', 10))
        self.entry_asunto.grid(row=1, column=1, padx=10, pady=10, sticky='ew')
        self.entry_asunto.insert(0, "Solicitud para {empresa} - {nombre}")
        
        # Cuerpo del mensaje
        ttk.Label(frame, text="Cuerpo del Mensaje:", style='Section.TLabel').grid(row=2, column=0, sticky='nw', padx=10, pady=10)
        self.text_cuerpo = scrolledtext.ScrolledText(frame, width=60, height=15, font=('Arial', 10))
        self.text_cuerpo.grid(row=2, column=1, padx=10, pady=10, sticky='nsew')
        
        # Texto de ejemplo en el cuerpo
        cuerpo_ejemplo = """Estimados señores de {empresa},

Me dirijo a ustedes para expresar mi interés en formar parte de su equipo de trabajo.

Mi nombre es {nombre} y estoy interesado en las oportunidades que su empresa ofrece.

Quedo a disposición para cualquier consulta.

Atentamente,
{nombre}"""
        self.text_cuerpo.insert('1.0', cuerpo_ejemplo)
        
        # Configurar grid weights
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(2, weight=1)
        
    def crear_pestana_base_datos(self, notebook):
        """Crea la pestaña de configuración de base de datos."""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="📊 Base de Datos")
        
        # Selección de archivo Excel
        ttk.Label(frame, text="Archivo Excel:", style='Section.TLabel').grid(row=0, column=0, sticky='w', padx=10, pady=10)
        
        file_frame = ttk.Frame(frame)
        file_frame.grid(row=0, column=1, columnspan=2, sticky='ew', padx=10, pady=10)
        
        self.entry_excel = ttk.Entry(file_frame, width=50, font=('Arial', 10))
        self.entry_excel.grid(row=0, column=0, padx=5, pady=5, sticky='ew')
        ttk.Button(file_frame, text="Buscar Excel", command=self.buscar_excel).grid(row=0, column=1, padx=5, pady=5)
        
        # Vista previa de datos
        ttk.Label(frame, text="Vista Previa de Datos:", style='Section.TLabel').grid(row=1, column=0, sticky='nw', padx=10, pady=10)
        
        # Frame para la tabla de vista previa
        table_frame = ttk.Frame(frame)
        table_frame.grid(row=1, column=1, columnspan=2, sticky='nsew', padx=10, pady=10)
        
        # Treeview para mostrar datos
        columns = ('#1', '#2', '#3', '#4')
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=8)
        
        # Scrollbars
        v_scroll = ttk.Scrollbar(table_frame, orient='vertical', command=self.tree.yview)
        h_scroll = ttk.Scrollbar(table_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')
        
        # Botón para cargar vista previa
        ttk.Button(frame, text="Cargar Vista Previa", command=self.cargar_vista_previa).grid(row=2, column=1, pady=10)
        
        # Configurar grid weights
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(1, weight=1)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        file_frame.columnconfigure(0, weight=1)
        
    def crear_pestana_envio(self, notebook):
        """Crea la pestaña de resumen y envío."""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="🚀 Envío Masivo")
        
        # Resumen de configuración
        ttk.Label(frame, text="Resumen de Configuración", style='Title.TLabel').grid(row=0, column=0, columnspan=2, pady=20)
        
        self.text_resumen = scrolledtext.ScrolledText(frame, width=80, height=15, font=('Arial', 9))
        self.text_resumen.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky='nsew')
        
        # Barra de progreso
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=2, column=0, columnspan=2, padx=20, pady=10, sticky='ew')
        
        # Etiqueta de progreso
        self.label_progreso = ttk.Label(frame, text="Listo para comenzar")
        self.label_progreso.grid(row=3, column=0, columnspan=2, pady=5)
        
        # Etiqueta de estado de pausa
        self.label_pausa = ttk.Label(frame, text="", foreground="blue")
        self.label_pausa.grid(row=4, column=0, columnspan=2, pady=2)
        
        # Botones de control
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=20)
        
        self.btn_generar = ttk.Button(btn_frame, text="Generar Resumen", command=self.generar_resumen)
        self.btn_generar.pack(side='left', padx=10)
        
        self.btn_iniciar = ttk.Button(btn_frame, text="Iniciar Envío", command=self.gestor.iniciar_envio)
        self.btn_iniciar.pack(side='left', padx=10)
        
        self.btn_detener = ttk.Button(btn_frame, text="Detener Envío", command=self.gestor.detener_envio, state='disabled')
        self.btn_detener.pack(side='left', padx=10)
        
        # Configurar grid weights
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        
    def crear_area_log(self):
        """Crea el área de log en la parte inferior."""
        log_frame = ttk.LabelFrame(self.root, text="📝 Log de Actividad")
        log_frame.pack(fill='x', padx=20, pady=10)
        
        self.text_log = scrolledtext.ScrolledText(log_frame, height=8, font=('Consolas', 9))
        self.text_log.pack(fill='both', expand=True, padx=10, pady=10)
        
    def toggle_adjuntar(self):
        """Muestra u oculta la sección de archivo adjunto."""
        if self.adjuntar_var.get():
            self.frame_archivo.grid()
        else:
            self.frame_archivo.grid_remove()
            
    def buscar_archivo(self):
        """Abre diálogo para buscar archivo a adjuntar."""
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo a adjuntar",
            filetypes=[("Todos los archivos", "*.*")]
        )
        if archivo:
            self.entry_archivo.delete(0, tk.END)
            self.entry_archivo.insert(0, archivo)
            
    def buscar_excel(self):
        """Abre diálogo para buscar archivo Excel."""
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if archivo:
            self.entry_excel.delete(0, tk.END)
            self.entry_excel.insert(0, archivo)
            self.cargar_vista_previa()
            
    def cargar_vista_previa(self):
        """Carga una vista previa de los datos del Excel."""
        excel_path = self.entry_excel.get()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Error", "Por favor seleccione un archivo Excel válido")
            return
            
        try:
            df = pd.read_excel(excel_path)
            
            # Limpiar treeview
            for item in self.tree.get_children():
                self.tree.delete(item)
                
            # Configurar columnas
            columnas = df.columns.tolist()[:4]  # Mostrar máximo 4 columnas
            self.tree['columns'] = columnas
            
            for col in columnas:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=100)
                
            # Agregar datos (mostrar máximo 10 filas)
            for index, row in df.head(10).iterrows():
                self.tree.insert('', 'end', values=row.tolist()[:4])
                
            self.log(f"✓ Vista previa cargada: {len(df)} registros encontrados")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar el Excel: {str(e)}")
            
    def generar_resumen(self):
        """Genera un resumen de la configuración actual."""
        try:
            resumen = "📋 RESUMEN DE CONFIGURACIÓN\n"
            resumen += "=" * 50 + "\n\n"
            
            # Configuración de correo
            resumen += "📧 CONFIGURACIÓN DE CORREO:\n"
            resumen += f"   Remitente: {self.entry_remitente.get()}\n"
            resumen += f"   Adjuntar archivo: {'Sí' if self.adjuntar_var.get() else 'No'}\n"
            if self.adjuntar_var.get():
                resumen += f"   Archivo adjunto: {self.entry_archivo.get()}\n"
            resumen += "\n"
            
            # Base de datos
            resumen += "📊 BASE DE DATOS:\n"
            resumen += f"   Archivo Excel: {self.entry_excel.get()}\n"
            resumen += "\n"
            
            # Mensaje
            resumen += "✉️ CONFIGURACIÓN DEL MENSAJE:\n"
            resumen += f"   Asunto: {self.entry_asunto.get()}\n"
            resumen += f"   Cuerpo:\n"
            
            cuerpo = self.text_cuerpo.get('1.0', tk.END)
            for linea in cuerpo.split('\n'):
                resumen += f"   {linea}\n"
                
            self.text_resumen.delete('1.0', tk.END)
            self.text_resumen.insert('1.0', resumen)
            self.log("✓ Resumen generado correctamente")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar resumen: {str(e)}")
            
    def log(self, mensaje):
        """Agrega un mensaje al log con timestamp."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_line = f"[{timestamp}] {mensaje}\n"
        
        self.text_log.insert(tk.END, log_line)
        self.text_log.see(tk.END)
        self.root.update()
        
    def actualizar_progreso(self, actual, total):
        """Actualiza la barra de progreso."""
        if total > 0:
            porcentaje = (actual / total) * 100
            self.progress_var.set(porcentaje)
            self.label_progreso.config(text=f"Progreso: {actual}/{total} ({porcentaje:.1f}%)")
        self.root.update()
        
    def actualizar_estado_pausa(self, segundos_restantes):
        """Actualiza el estado de la pausa."""
        if segundos_restantes > 0:
            self.label_pausa.config(text=f"⏳ Pausa: {segundos_restantes} segundos restantes")
        else:
            self.label_pausa.config(text="")
        self.root.update()
        
    def actualizar_estado_botones(self, envio_activo):
        """Actualiza el estado de los botones según el estado del envío."""
        if envio_activo:
            self.btn_iniciar.config(state='disabled')
            self.btn_detener.config(state='normal')
            self.estado_label.config(text="🟢 ENVIANDO", foreground="green")
        else:
            self.btn_iniciar.config(state='normal')
            self.btn_detener.config(state='disabled')
            self.estado_label.config(text="🔴 Listo", foreground="red")
        self.root.update()
        
    def run(self):
        """Inicia la aplicación."""
        self.root.mainloop()


# =============================================================================
# FUNCIÓN PRINCIPAL
# =============================================================================
def main():
    """
    Función principal que inicia la aplicación.
    """
    app = InterfazGrafica()
    app.run()

if __name__ == "__main__":
    main()
