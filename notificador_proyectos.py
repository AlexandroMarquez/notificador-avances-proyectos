
import pandas as pd
#Para leer el excel con los datos de envío
import tkinter as tk
from tkinter import filedialog, messagebox
#librería con la que construí la interfaz gráfica
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
#Librería para enviar correos
import ssl
#Seguridad para el envío de correos

APP_TITLE = "Notificador de Avances de Proyectos - Aktivgroup"

DEFAULT_SUBJECT = "Estado semanal | {Cliente} · {Proyecto} · Avance {pct}%"
DEFAULT_BODY = (
    "Hola {pm_cliente_nombre},<br><br>"
    "<b>Avance del proyecto:</b> {Proyecto} ({Cliente})<br>"
    "<ul>"
    "<li><b>Avance:</b> {pct}%</li>"
    "<li><b>Hitos cumplidos (última semana):</b> {hitos}</li>"
    "<li><b>Bloqueos / Riesgos:</b> {riesgos}</li>"
    "<li><b>Próximos pasos:</b> {proximos}</li>"
    "<li><b>Próxima entrega:</b> {fecha_entrega}</li>"
    "</ul>"
    "Saludos,<br>"
    "{pm_aktiv_nombre} · Aktivgroup"
)

def build_email(sender, to_addr, subject, html_body):
    msg = MIMEMultipart('alternative')
    msg['From'] = sender
    msg['To'] = to_addr
    msg['Subject'] = subject
    msg.attach(MIMEText(html_body, 'html', 'utf-8'))
    return msg

class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.excel_path = tk.StringVar()
        self.sender_email = tk.StringVar()
        self.sender_pass = tk.StringVar()
        self.subject_tpl = tk.StringVar(value=DEFAULT_SUBJECT)
        self.body_tpl = tk.StringVar(value=DEFAULT_BODY)
        self.smtp_host = tk.StringVar(value="smtp.gmail.com")
        self.smtp_port = tk.IntVar(value=587)

        # Layout
        row = 0
        tk.Label(root, text="Ruta del Documento:").grid(row=row, column=0, sticky="w", padx=6, pady=4)
        tk.Entry(root, textvariable=self.excel_path, width=60).grid(row=row, column=1, padx=6, pady=4)
        tk.Button(root, text="Seleccionar...", command=self.pick_excel).grid(row=row, column=2, padx=6, pady=4)
        row += 1

        tk.Label(root, text="Correo remitente:").grid(row=row, column=0, sticky="w", padx=6, pady=4)
        tk.Entry(root, textvariable=self.sender_email, width=40).grid(row=row, column=1, padx=6, pady=4, sticky="w")
        row += 1

        tk.Label(root, text="Contraseña de aplicación:").grid(row=row, column=0, sticky="w", padx=6, pady=4)
        tk.Entry(root, textvariable=self.sender_pass, show="*", width=40).grid(row=row, column=1, padx=6, pady=4, sticky="w")
        row += 1

        tk.Label(root, text="SMTP host:").grid(row=row, column=0, sticky="w", padx=6, pady=4)
        tk.Entry(root, textvariable=self.smtp_host, width=40).grid(row=row, column=1, padx=6, pady=4, sticky="w")
        tk.Label(root, text="Puerto:").grid(row=row, column=2, sticky="e", padx=6, pady=4)
        tk.Entry(root, textvariable=self.smtp_port, width=6).grid(row=row, column=3, padx=6, pady=4, sticky="w")
        row += 1

        tk.Label(root, text="Asunto (plantilla):").grid(row=row, column=0, sticky="nw", padx=6, pady=4)
        tk.Entry(root, textvariable=self.subject_tpl, width=80).grid(row=row, column=1, columnspan=2, padx=6, pady=4, sticky="w")
        row += 1

        tk.Label(root, text="Cuerpo (HTML plantilla):").grid(row=row, column=0, sticky="nw", padx=6, pady=4)
        self.body_text = tk.Text(root, width=80, height=10)
        self.body_text.grid(row=row, column=1, columnspan=2, padx=6, pady=4, sticky="w")
        self.body_text.insert("1.0", DEFAULT_BODY)
        row += 1

        tk.Button(root, text="Enviar correos", command=self.send_all).grid(row=row, column=1, pady=8, sticky="w")
        tk.Button(root, text="Vista previa (primero)", command=self.preview_first).grid(row=row, column=1, pady=8, sticky="e")
        row += 1

        tk.Label(root, text="Resultados:").grid(row=row, column=0, sticky="nw", padx=6, pady=4)
        self.result_text = tk.Text(root, width=80, height=12)
        self.result_text.grid(row=row, column=1, columnspan=2, padx=6, pady=4, sticky="w")

    def pick_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.excel_path.set(path)

    def load_rows(self):
        if not self.excel_path.get():
            raise ValueError("Selecciona un archivo Excel.")
        df = pd.read_excel(self.excel_path.get())
        required = [
            "Cliente","Proyecto","% Avance",
            "Hitos cumplidos (última semana)","Bloqueos / Riesgos","Próximos pasos",
            "Fecha próxima entrega","PM Cliente (Nombre)","Correo PM Cliente",
            "PM Aktivgroup (Nombre)","Correo PM Aktivgroup"
        ]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"Faltan columnas en el Excel: {missing}")

        for _, r in df.iterrows():
            yield dict(
                Cliente=r["Cliente"],
                Proyecto=r["Proyecto"],
                pct=int(r["% Avance"]),
                hitos=r["Hitos cumplidos (última semana)"],
                riesgos=r["Bloqueos / Riesgos"],
                proximos=r["Próximos pasos"],
                fecha_entrega=str(r["Fecha próxima entrega"]),
                pm_cliente_nombre=r["PM Cliente (Nombre)"],
                pm_cliente_correo=r["Correo PM Cliente"],
                pm_aktiv_nombre=r["PM Aktivgroup (Nombre)"],
                pm_aktiv_correo=r["Correo PM Aktivgroup"],
            )

    def preview_first(self):
        try:
            rows = list(self.load_rows())
            if not rows:
                messagebox.showinfo("Vista previa", "No hay filas en el Excel.")
                return
            row = rows[0]
            subject = self.subject_tpl.get().format(Cliente=row["Cliente"], Proyecto=row["Proyecto"], pct=row["pct"])
            body_html = self.body_text.get("1.0","end").format(**row)
            self.result_text.delete("1.0","end")
            self.result_text.insert("1.0", f"Para: {row['pm_cliente_correo']}\nAsunto: {subject}\n\nCuerpo HTML:\n{body_html}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def send_all(self):
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP(self.smtp_host.get(), int(self.smtp_port.get())) as server:
                server.starttls(context=context)
                server.login(self.sender_email.get(), self.sender_pass.get())
                sent = 0
                for row in self.load_rows():
                    subject = self.subject_tpl.get().format(Cliente=row["Cliente"], Proyecto=row["Proyecto"], pct=row["pct"])
                    body_html = self.body_text.get("1.0","end").format(**row)
                    msg = build_email(self.sender_email.get(), row["pm_cliente_correo"], subject, body_html)
                    server.sendmail(self.sender_email.get(), [row["pm_cliente_correo"]], msg.as_string())
                    sent += 1
                    self.result_text.insert("end", f"✔ Enviado a {row['pm_cliente_correo']} | {row['Cliente']} - {row['Proyecto']}\n")
            messagebox.showinfo("Listo", f"Correos enviados: {sent}")
        except Exception as e:
            messagebox.showerror("Error al enviar", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
