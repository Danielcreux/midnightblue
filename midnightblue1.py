import os
import json
import smtplib
import time
import openpyxl
import threading
import ttkbootstrap as ttk
from tkinter import filedialog, StringVar, messagebox, scrolledtext
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from PIL import Image, ImageTk  # Importar Pillow para trabajar con imágenes

# Archivo de configuración
CONFIG_FILE = "config.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {}

def save_config():
    config = {
        "smtp_server": smtp_var.get(),
        "username": user_var.get(),
        "password": pass_var.get()
    }
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f)

def select_html():
    file_path = filedialog.askopenfilename(filetypes=[["HTML files", "*.html"]])
    html_var.set(file_path)

def select_excel():
    file_path = filedialog.askopenfilename(filetypes=[["Excel files", "*.xlsx;*.ods"]])
    excel_var.set(file_path)

def log_message(message):
    console_text.config(state='normal')
    console_text.insert('end', message + "\n")
    console_text.config(state='disabled')
    console_text.yview('end')

def send_emails():
    def task():
        smtp_server = smtp_var.get()
        username = user_var.get()
        password = pass_var.get()
        html_file = html_var.get()
        excel_file = excel_var.get()
        
        if not all([smtp_server, username, password, html_file, excel_file]):
            messagebox.showerror("Error", "Todos los campos son obligatorios")
            return
        
        try:
            with open(html_file, "r", encoding="utf-8") as file:
                html_content = file.read()
        except FileNotFoundError:
            messagebox.showerror("Error", "Archivo HTML no encontrado")
            return
        
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active
        total_emails = sheet.max_row - 1
        root.after(0, lambda: progress_var.set(0))
        
        try:
            server = smtplib.SMTP(smtp_server, 587)
            server.starttls()
            server.login(username, password)
            log_message("Conexión exitosa al servidor SMTP")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo conectar al servidor SMTP: {e}")
            return
        
        sent_count = 0
        for row in sheet.iter_rows(min_row=2, values_only=True):
            receiver_email = row[0]
            if receiver_email:
                message = MIMEMultipart("alternative")
                message["Subject"] = "Your email subject"
                message["From"] = username
                message["To"] = receiver_email
                message["Date"] = formatdate(localtime=True)
                message.attach(MIMEText(html_content, "html"))
                
                try:
                    server.sendmail(username, receiver_email, message.as_string())
                    sent_count += 1
                    root.after(0, lambda count=sent_count: progress_var.set((count / total_emails) * 100))
                    log_message(f"✅ Enviado a {receiver_email}")
                    time.sleep(10)
                except Exception as e:
                    log_message(f"❌ Error con {receiver_email}: {e}")
        
        server.quit()
        root.after(0, lambda: progress_var.set(100))  # Asegurar que la barra llegue al 100%
        messagebox.showinfo("Éxito", "Todos los correos han sido enviados")
    
    thread = threading.Thread(target=task)
    thread.start()

# Cargar configuración previa
config = load_config()

# Crear ventana
root = ttk.Window(themename="superhero")
root.title("Email Sender")
root.geometry("500x550")

# Agregar el logo a la ventana
logo_path = "midnightblue.png"  # Ruta del logo PNG
if os.path.exists(logo_path):
    logo_image = Image.open(logo_path)
    logo_image = logo_image.resize((100, 100), Image.Resampling.LANCZOS)  # Ajustar el tamaño
    logo_photo = ImageTk.PhotoImage(logo_image)
    ttk.Label(root, image=logo_photo).pack(pady=10)  # Mostrar el logo en la ventana

# Variables
smtp_var = StringVar(value=config.get("smtp_server", ""))
user_var = StringVar(value=config.get("username", ""))
pass_var = StringVar(value=config.get("password", ""))
html_var = StringVar()
excel_var = StringVar()
progress_var = ttk.DoubleVar()

# Widgets
ttk.Label(root, text="SMTP Server:").pack(pady=5)
ttk.Entry(root, textvariable=smtp_var).pack(fill='x', padx=20)

ttk.Label(root, text="Usuario:").pack(pady=5)
ttk.Entry(root, textvariable=user_var).pack(fill='x', padx=20)

ttk.Label(root, text="Contraseña:").pack(pady=5)
ttk.Entry(root, textvariable=pass_var, show='*').pack(fill='x', padx=20)

ttk.Button(root, text="Seleccionar HTML", command=select_html).pack(pady=5)
ttk.Entry(root, textvariable=html_var, state='readonly').pack(fill='x', padx=20)

ttk.Button(root, text="Seleccionar Excel", command=select_excel).pack(pady=5)
ttk.Entry(root, textvariable=excel_var, state='readonly').pack(fill='x', padx=20)

ttk.Button(root, text="Enviar Correos", command=send_emails, bootstyle='success').pack(pady=20)
ttk.Progressbar(root, variable=progress_var, maximum=100).pack(fill='x', padx=20, pady=5)

ttk.Label(root, text="Consola de mensajes:").pack(pady=5)
console_text = scrolledtext.ScrolledText(root, height=8, state='disabled')
console_text.pack(fill='both', padx=20, pady=5)

ttk.Button(root, text="Guardar Configuración", command=save_config, bootstyle='info').pack(pady=5)

# Ejecutar ventana
root.mainloop()


