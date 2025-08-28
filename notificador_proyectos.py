import pandas as pd
import smtplib, ssl
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ---------- Estilo opcional: ttkbootstrap ----------
BOOTSTRAP = False
try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *
    BOOTSTRAP = True
except Exception:
    BOOTSTRAP = False

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

# ---------- Email ----------
def build_email(sender, to_addr, subject, html_body):
    msg = MIMEMultipart("alternative")
    msg["From"] = sender
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    return msg


class App:
    REQUIRED_COLUMNS = [
        "Cliente","Proyecto","% Avance",
        "Hitos cumplidos (última semana)","Bloqueos / Riesgos","Próximos pasos",
        "Fecha próxima entrega","PM Cliente (Nombre)","Correo PM Cliente",
        "PM Aktivgroup (Nombre)","Correo PM Aktivgroup"
    ]

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("980x720")
        self.root.minsize(920, 660)
        self.root.columnconfigure(0, weight=1)

        # ---------- Paleta y estilos ----------
        self.ACCENT = "#2563EB"     # azul primario
        self.ACCENT_H = "#1E40AF"   # hover
        self.SECOND  = "#E5E7EB"    # gris claro
        self.TEXT_M  = "#475569"    # slate-600
        self.OK      = "#16A34A"
        self.ERROR   = "#DC2626"

        if BOOTSTRAP:
            self.root.style = tb.Style(theme="cosmo")
            self.BTN_PRIMARY   = {"bootstyle": "primary"}
            self.BTN_SECONDARY = {"bootstyle": "secondary"}
            self.BTN_SUCCESS   = {"bootstyle": "success"}
        else:
            style = ttk.Style()
            try:
                style.theme_use("vista")
            except:
                style.theme_use("clam")

            style.configure("Header.TLabel", font=("Segoe UI Semibold", 16))
            style.configure("Subheader.TLabel", foreground=self.TEXT_M)
            style.configure("TLabel", font=("Segoe UI", 10))
            style.configure("TButton", font=("Segoe UI", 10), padding=6)
            style.configure("Card.TLabelframe", padding=12)
            style.configure("Card.TLabelframe.Label", font=("Segoe UI Semibold", 11))

            # estilos de botón coloreados (fallback sin bootstrap)
            style.configure("Primary.TButton", background=self.ACCENT, foreground="white")
            style.map("Primary.TButton",
                      background=[("active", self.ACCENT_H), ("pressed", self.ACCENT_H)])
            style.configure("Secondary.TButton", background=self.SECOND)
            style.map("Secondary.TButton",
                      background=[("active", "#CDD2D8"), ("pressed", "#CDD2D8")])

            self.BTN_PRIMARY   = {"style": "Primary.TButton"}
            self.BTN_SECONDARY = {"style": "Secondary.TButton"}
            self.BTN_SUCCESS   = {"style": "Primary.TButton"}

        PADX, PADY = 12, 8

        # ---------- Vars ----------
        self.excel_path   = tk.StringVar()
        self.sender_email = tk.StringVar()
        self.sender_pass  = tk.StringVar()
        self.subject_tpl  = tk.StringVar(value=DEFAULT_SUBJECT)
        self.smtp_host    = tk.StringVar(value="smtp.gmail.com")
        self.smtp_port    = tk.IntVar(value=587)
        self.show_pass    = tk.BooleanVar(value=False)

        # ---------- Header ----------
        header = (tb.Frame(self.root, padding=PADX) if BOOTSTRAP
                  else ttk.Frame(self.root, padding=PADX))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        (tb.Label(header, text="Notificador de Avances de Proyectos")
            if BOOTSTRAP else ttk.Label(header, text="Notificador de Avances de Proyectos",
                                        style="Header.TLabel")
        ).grid(row=0, column=0, sticky="w")
        (tb.Label(header, text="Genera correos masivos con contexto y avance",
                  bootstyle="secondary") if BOOTSTRAP
         else ttk.Label(header, text="Genera correos masivos con contexto y avance",
                        style="Subheader.TLabel")
        ).grid(row=1, column=0, sticky="w", pady=(2,0))

        # ---------- Fuente de datos ----------
        src = (tb.Labelframe(self.root, text="Fuente de datos", padding=PADX) if BOOTSTRAP
               else ttk.Labelframe(self.root, text="Fuente de datos", style="Card.TLabelframe"))
        src.grid(row=1, column=0, sticky="ew", padx=PADX, pady=(0,6))
        src.columnconfigure(1, weight=1)

        ttk.Label(src, text="Documento:").grid(row=0, column=0, sticky="w", padx=(0,8), pady=PADY)
        ttk.Entry(src, textvariable=self.excel_path).grid(row=0, column=1, sticky="ew", pady=PADY)
        (tb.Button(src, text="Seleccionar…", command=self.pick_excel, **self.BTN_SECONDARY) if BOOTSTRAP
         else ttk.Button(src, text="Seleccionar…", command=self.pick_excel, **self.BTN_SECONDARY)
        ).grid(row=0, column=2, padx=(8,0), pady=PADY)

        # ---------- SMTP ----------
        smtp = (tb.Labelframe(self.root, text="SMTP y cuenta", padding=PADX) if BOOTSTRAP
               else ttk.Labelframe(self.root, text="SMTP y cuenta", style="Card.TLabelframe"))
        smtp.grid(row=2, column=0, sticky="ew", padx=PADX, pady=6)
        for c in range(5):
            smtp.columnconfigure(c, weight=1)

        ttk.Label(smtp, text="Correo remitente:").grid(row=0, column=0, sticky="w", padx=(0,8), pady=PADY)
        ttk.Entry(smtp, textvariable=self.sender_email).grid(row=0, column=1, sticky="ew", pady=PADY)

        ttk.Label(smtp, text="Contraseña de aplicación:").grid(row=0, column=2, sticky="w", padx=(12,8), pady=PADY)
        self.pass_entry = ttk.Entry(smtp, textvariable=self.sender_pass, show="•")
        self.pass_entry.grid(row=0, column=3, sticky="ew", pady=PADY)
        ttk.Checkbutton(smtp, text="Ver", variable=self.show_pass,
                        command=self.toggle_pass).grid(row=0, column=4, sticky="w", padx=(8,0), pady=PADY)

        ttk.Label(smtp, text="SMTP host:").grid(row=1, column=0, sticky="w", padx=(0,8), pady=PADY)
        self.host_combo = ttk.Combobox(
            smtp, textvariable=self.smtp_host,
            values=["smtp.gmail.com", "smtp.office365.com", "smtp.mail.yahoo.com"]
        )
        self.host_combo.grid(row=1, column=1, sticky="ew", pady=PADY)

        ttk.Label(smtp, text="Puerto:").grid(row=1, column=2, sticky="w", padx=(12,8), pady=PADY)
        ttk.Entry(smtp, textvariable=self.smtp_port, width=8).grid(row=1, column=3, sticky="w", pady=PADY)

        # ---------- Plantillas ----------
        tpl = (tb.Labelframe(self.root, text="Plantillas", padding=PADX) if BOOTSTRAP
               else ttk.Labelframe(self.root, text="Plantillas", style="Card.TLabelframe"))
        tpl.grid(row=3, column=0, sticky="nsew", padx=PADX, pady=6)
        tpl.columnconfigure(1, weight=1)
        self.root.rowconfigure(3, weight=1)

        ttk.Label(tpl, text="Asunto:").grid(row=0, column=0, sticky="nw", padx=(0,8), pady=(0,4))
        ttk.Entry(tpl, textvariable=self.subject_tpl).grid(row=0, column=1, sticky="ew", pady=(0,4))

        ttk.Label(tpl, text="Cuerpo (HTML):").grid(row=1, column=0, sticky="nw", padx=(0,8), pady=(4,0))
        self.body_text = ScrolledText(tpl, height=12, font=("Consolas", 10), wrap="word")
        self.body_text.grid(row=1, column=1, sticky="nsew", pady=(4,0))
        self.body_text.insert("1.0", DEFAULT_BODY)

        # ---------- Acciones ----------
        actions = (tb.Frame(self.root, padding=(PADX, 0)) if BOOTSTRAP
                   else ttk.Frame(self.root, padding=(PADX, 0)))
        actions.grid(row=4, column=0, sticky="ew")
        actions.columnconfigure(0, weight=1)

        bar = (tb.Frame(actions) if BOOTSTRAP else ttk.Frame(actions))
        bar.grid(row=0, column=0, sticky="e", pady=(0, 2))

        self.btn_preview = (tb.Button(bar, text="👁  Vista previa", command=self.preview_first, **self.BTN_SECONDARY)
                            if BOOTSTRAP else ttk.Button(bar, text="Vista previa",
                                                          command=self.preview_first, **self.BTN_SECONDARY))
        self.btn_preview.grid(row=0, column=0, padx=(0,8))

        self.btn_send = (tb.Button(bar, text="✈  Enviar correos", command=self.send_all, **self.BTN_PRIMARY)
                         if BOOTSTRAP else ttk.Button(bar, text="Enviar correos",
                                                       command=self.send_all, **self.BTN_PRIMARY))
        self.btn_send.grid(row=0, column=1)

        # Barra de progreso (visible durante el envío)
        self.progress = ttk.Progressbar(actions, mode="indeterminate")
        self.progress.grid(row=1, column=0, sticky="ew", pady=(8,0))
        self.progress.grid_remove()  # oculta por defecto

        # ---------- Resultados ----------
        res = (tb.Labelframe(self.root, text="Resultados", padding=PADX) if BOOTSTRAP
               else ttk.Labelframe(self.root, text="Resultados", style="Card.TLabelframe"))
        res.grid(row=5, column=0, sticky="nsew", padx=PADX, pady=6)
        self.root.rowconfigure(5, weight=1)

        self.result_text = ScrolledText(res, height=9, font=("Consolas", 10), wrap="word")
        self.result_text.grid(row=0, column=0, sticky="nsew")

        # ---------- Status bar ----------
        self.status = tk.StringVar(value="Listo")
        self.status_dot = tk.StringVar(value="●")  # usaremos color por estilo/label separado

        status_frame = (tb.Frame(self.root, padding=(12,6)) if BOOTSTRAP
                        else ttk.Frame(self.root, padding=(12,6)))
        status_frame.grid(row=6, column=0, sticky="ew")

        self.dot_label = ttk.Label(status_frame, textvariable=self.status_dot, foreground=self.OK)
        self.dot_label.grid(row=0, column=0, sticky="w")
        ttk.Label(status_frame, textvariable=self.status).grid(row=0, column=1, sticky="w", padx=(6,0))

    # ---------- Helpers UI ----------
    def toggle_pass(self):
        self.pass_entry.configure(show="" if self.show_pass.get() else "•")

    def pick_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.excel_path.set(path)

    def _busy(self, on=True):
        """Activa/desactiva progreso y botones (UX de envío)."""
        if on:
            self.btn_preview.configure(state="disabled")
            self.btn_send.configure(state="disabled")
            self.progress.grid()
            self.progress.start(12)
            self.dot_label.configure(foreground="#F59E0B")  # amarillo
            self.status.set("Procesando…")
        else:
            self.btn_preview.configure(state="normal")
            self.btn_send.configure(state="normal")
            self.progress.stop()
            self.progress.grid_remove()
            self.dot_label.configure(foreground=self.OK)   # verde

    # ---------- Data ----------
    def load_rows(self):
        if not self.excel_path.get():
            raise ValueError("Selecciona un archivo Excel.")
        df = pd.read_excel(self.excel_path.get())
        missing = [c for c in self.REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            raise ValueError(f"Faltan columnas en el Excel: {', '.join(missing)}")

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

    # ---------- Actions ----------
    def preview_first(self):
        try:
            rows = list(self.load_rows())
            if not rows:
                messagebox.showinfo("Vista previa", "No hay filas en el Excel.")
                return
            row = rows[0]
            subject = self.subject_tpl.get().format(
                Cliente=row["Cliente"], Proyecto=row["Proyecto"], pct=row["pct"]
            )
            body_html = self.body_text.get("1.0", "end").format(**row)
            self.result_text.delete("1.0", "end")
            self.result_text.insert("1.0", f"Para: {row['pm_cliente_correo']}\nAsunto: {subject}\n\nCuerpo HTML:\n{body_html}")
            self.status.set("Vista previa generada.")
            self.dot_label.configure(foreground=self.OK)
        except Exception as e:
            self.dot_label.configure(foreground=self.ERROR)
            self.status.set("Error en vista previa")
            messagebox.showerror("Error", str(e))

    def send_all(self):
        try:
            self._busy(True)
            context = ssl.create_default_context()
            with smtplib.SMTP(self.smtp_host.get(), int(self.smtp_port.get())) as server:
                server.starttls(context=context)
                server.login(self.sender_email.get(), self.sender_pass.get())
                sent = 0
                for row in self.load_rows():
                    subject = self.subject_tpl.get().format(
                        Cliente=row["Cliente"], Proyecto=row["Proyecto"], pct=row["pct"]
                    )
                    body_html = self.body_text.get("1.0", "end").format(**row)
                    msg = build_email(self.sender_email.get(), row["pm_cliente_correo"], subject, body_html)
                    server.sendmail(self.sender_email.get(), [row["pm_cliente_correo"]], msg.as_string())
                    sent += 1
                    self.result_text.insert(
                        "end", f"✔ Enviado a {row['pm_cliente_correo']} | {row['Cliente']} - {row['Proyecto']}\n"
                    )
            self.status.set(f"Listo. Correos enviados: {sent}")
            self.dot_label.configure(foreground=self.OK)
            messagebox.showinfo("Listo", f"Correos enviados: {sent}")
        except Exception as e:
            self.dot_label.configure(foreground=self.ERROR)
            self.status.set("Error al enviar")
            messagebox.showerror("Error al enviar", str(e))
        finally:
            self._busy(False)


if __name__ == "__main__":
    if BOOTSTRAP:
        root = tb.Window(title=APP_TITLE, themename="cosmo")
    else:
        root = tk.Tk()
    App(root)
    root.mainloop()
