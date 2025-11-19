import tkinter as tk
from tkinter import ttk
import subprocess
import sys
import os
import time


def abrir_quick_alert():
    ruta = os.path.join(os.path.dirname(__file__), "app_quick_alert.py")
    subprocess.Popen([sys.executable, ruta])


def abrir_oberon_news():
    ruta = os.path.join(os.path.dirname(__file__), "app_oberon_news.py")
    subprocess.Popen([sys.executable, ruta])


def abrir_historial_quick_alert():
    ruta = os.path.join(os.path.dirname(__file__), "Historial Quick Alert.pbix")
    os.startfile(ruta)


class Launcher(tk.Tk):
    def __init__(self):
        super().__init__()

        # Fade in inicial
        self.attributes("-alpha", 0.0)

        self.title("Oberon Suite – Launcher")
        self.geometry("1150x520")

        # COLOR CORPORATIVO
        self.config(bg="#003057")

        self.resizable(True, True)
        self.minsize(950, 460)

        self._estilos()
        self._ui()
        self.fade_in()

    # -------------------------------
    # ANIMACIÓN DE APERTURA
    # -------------------------------
    def fade_in(self):
        for i in range(0, 11):
            alpha = i / 10
            self.attributes("-alpha", alpha)
            self.update()
            time.sleep(0.03)

    # -------------------------------
    # ANIMACIÓN DE SALIDA
    # -------------------------------
    def fade_out(self):
        for i in range(10, -1, -1):
            alpha = i / 10
            self.attributes("-alpha", alpha)
            self.update()
            time.sleep(0.03)
        self.destroy()

    # -------------------------------
    # ESTILOS PREMIUM (Glass + Glow)
    # -------------------------------
    def _estilos(self):
        style = ttk.Style()
        style.theme_use("clam")

        # Fondo Glass
        style.configure(
            "Glass.TFrame",
            background="#0F263A",
            borderwidth=0
        )

        # CARD base estilo vidrio
        style.configure(
            "Card.TFrame",
            background="#12314D",
            relief="flat",
            borderwidth=1
        )

        # Glow en hover
        style.configure(
            "Glow.TFrame",
            background="#146B99",
            relief="solid",
            borderwidth=1
        )

        style.configure(
            "CardTitle.TLabel",
            background="#12314D",
            foreground="white",
            font=("Century Gothic", 15, "bold")
        )

        style.configure(
            "CardText.TLabel",
            background="#12314D",
            foreground="white",
            font=("Century Gothic", 12)
        )

        # Título general
        style.configure(
            "Title.TLabel",
            background="#003057",
            foreground="white",
            font=("Century Gothic", 30, "bold")
        )

        # Botón premium
        style.configure(
            "TButton",
            font=("Century Gothic", 11, "bold"),
            foreground="white",
            background="#00A3AF",
            padding=8,
            borderwidth=0
        )

        style.map(
            "TButton",
            background=[("active", "#008C97")]
        )

    # -------------------------------
    # HOVER (Glow + Sombra dinámica)
    # -------------------------------
    def apply_hover(self, frame, shadow):
        frame.configure(style="Glow.TFrame")
        shadow.config(bg="#1C3F5C")  # sombra clara en hover

    def remove_hover(self, frame, shadow):
        frame.configure(style="Card.TFrame")
        shadow.config(bg="#00213F")  # sombra normal sin negro

    # -------------------------------
    # INTERFAZ
    # -------------------------------
    def _ui(self):

        # -------------------------------------------------------------
        # LOGO DENTRO DE LA VENTANA (PARTE SUPERIOR IZQUIERDA)
        # -------------------------------------------------------------
        ruta_logo = os.path.join(os.path.dirname(__file__), "Logo.png")
        self.logo_img = tk.PhotoImage(file=ruta_logo)
        logo_label = tk.Label(self, image=self.logo_img, bg="#003057")
        logo_label.place(x=25, y=15)     # posición exacta arriba a la izquierda

        ttk.Label(self, text="Oberon Suite", style="Title.TLabel").pack(pady=20)

        cont = ttk.Frame(self, style="Glass.TFrame")
        cont.pack(fill="both", expand=True, padx=30, pady=10)

        cont.columnconfigure((0, 1, 2), weight=1)
        cont.rowconfigure(0, weight=1)

        def crear_card(parent, titulo, texto, comando, col):

            # SOMBRA PREMIUM AZUL (NO NEGRO)
            shadow = tk.Frame(parent, bg="#00213F")
            shadow.grid(row=0, column=col, padx=28, pady=20, sticky="nsew")

            # Card Glass
            card = ttk.Frame(shadow, style="Card.TFrame")
            card.pack(padx=6, pady=6, fill="both", expand=True)

            inner = ttk.Frame(card, style="Card.TFrame")
            inner.pack(expand=True)

            ttk.Label(inner, text=titulo, style="CardTitle.TLabel").pack(pady=5)
            ttk.Label(inner, text=texto, style="CardText.TLabel").pack(pady=5)

            ttk.Button(inner, text="ABRIR", width=16, command=comando).pack(pady=12)

            # Hover
            card.bind("<Enter>", lambda e: self.apply_hover(card, shadow))
            card.bind("<Leave>", lambda e: self.remove_hover(card, shadow))

        crear_card(cont, "QUICK ALERT", "Generador profesional", abrir_quick_alert, 0)
        crear_card(cont, "OBERON NEWS", "Generador de boletín", abrir_oberon_news, 1)
        crear_card(cont, "HISTORIAL QUICK ALERT", "Dashboard en Power BI", abrir_historial_quick_alert, 2)

        ttk.Button(self, text="SALIR", width=20, command=self.fade_out).pack(pady=25)


if __name__ == "__main__":
    Launcher().mainloop()
