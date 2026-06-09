import customtkinter as ctk
from tkinter import messagebox
import threading, os, time, re, json
from PIL import Image
from datetime import datetime
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from agent import brain, vision
from office import word_controller, excel_controller, ppt_controller
from storage import db

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# ── PALETA macOS ──────────────────────────────────────────────────────────────
BG       = "#0d0d0f"        # fondo principal
CARD     = "#1c1c1e"        # tarjetas
CARD2    = "#2c2c2e"        # tarjetas secundarias
BORDER   = "#3a3a3c"        # bordes
BLUE     = "#2d7bf4"        # azul principal
BLUE2    = "#3a85f5"        # azul hover
GREEN    = "#30d158"        # verde
AMBER    = "#ff9f0a"        # ámbar/naranja
PURPLE   = "#7c3aed"        # morado enterprise
PURPLE2  = "#8b5cf6"
TEAL     = "#0f766e"
RED      = "#ff453a"
WHITE    = "#f5f5f7"
GRAY1    = "#8e8e93"        # texto secundario
GRAY2    = "#48484a"        # texto terciario
GRAY3    = "#3a3a3c"        # líneas/bordes suaves

# v14.6.2 — Colores de marca alineados con el icon oficial de WEP AI.
# W = Word (azul tipo Microsoft), E = Excel (verde), P = PowerPoint (naranja)
WORD_C   = "#2196F3"        # azul Word — vibrante, brillante (alineado al icon)
EXCEL_C  = "#34A853"        # verde Excel — Google green, oscuro suficiente para fondo dark
PPT_C    = "#FB8C00"        # naranja PowerPoint — Material orange 600

PLAN_COLORS   = {"free": GRAY1,  "pro": BLUE,   "biz": PURPLE2}
PLAN_NAMES_ES = {"free": "Personal Free", "pro": "Personal Pro", "biz": "Enterprise"}
PLAN_NAMES_EN = {"free": "Personal Free", "pro": "Personal Pro", "biz": "Enterprise"}
PLAN_PRICES   = {"free": "Gratis",        "pro": "$4.99/mes",    "biz": "$24.99/mes"}
PLAN_PRICES_EN= {"free": "Free",          "pro": "$4.99/mo",     "biz": "$24.99/mo"}

CONTROLLERS = {
    "word":       (word_controller.generate_word,      word_controller.open_in_word,      word_controller.apply_correction),
    "excel":      (excel_controller.generate_excel,    excel_controller.open_in_excel,    excel_controller.apply_correction),
    "powerpoint": (ppt_controller.generate_powerpoint, ppt_controller.open_in_powerpoint, ppt_controller.apply_correction),
}
APP_NAMES = {"word": "Microsoft Word", "excel": "Microsoft Excel", "powerpoint": "Microsoft PowerPoint"}
# ── EJEMPLOS DE PROMPTS POR TIPO ──────────────────────────────────────────
EXAMPLES_ES = {
    "auto": [
        "Contrato de trabajo indefinido para México con salario de 25,000 MXN",
        "Cotización para cliente de servicios de diseño con IVA y descuento",
        "Pitch deck para startup de tecnología, 10 slides para ronda seed",
        "Nómina mensual de empleados con prestaciones de ley y descuentos",
    ],
    "word": [
        "Contrato de arrendamiento para bodega en Medellín por 2 años con opción de renovación",
        "Carta de presentación profesional para gerente de ventas con 5 años de experiencia",
        "Acuerdo de confidencialidad (NDA) entre empresa y freelancer de software",
        "Demanda formal ante Superintendencia de Industria y Comercio por publicidad engañosa",
        "Política de teletrabajo y vacaciones para empresa de 30 empleados",
    ],
    "excel": [
        "Inventario de cremas y cosméticos con alertas de stock y valor total por categoría",
        "Presupuesto mensual familiar con gráfica de gastos fijos vs variables",
        "Nómina de empleados con salario, descuentos de salud, pensión y retención en la fuente",
        "Dashboard de ventas con KPIs, comparativo mes a mes y proyecciones Q4",
        "Cotización para cliente de servicios de diseño con IVA, descuento y términos de pago",
    ],
    "powerpoint": [
        "Pitch deck para startup de tecnología educativa, 10 slides para ronda seed",
        "Presentación de resultados Q3 para junta directiva con gráficas de crecimiento",
        "Propuesta comercial de servicios de consultoría financiera para PyME",
        "Capacitación de inducción para empleados nuevos, 12 slides con valores y cultura",
        "Informe de sostenibilidad ESG para empresa manufacturera, 15 slides",
    ],
}
EXAMPLES_EN = {
    "auto": [
        "Permanent employment contract for Mexico with a 25,000 MXN salary",
        "Client quotation for design services with tax and discount",
        "Pitch deck for a tech startup, 10 slides for a seed round",
        "Monthly employee payroll with statutory benefits and deductions",
    ],
    "word": [
        "Commercial warehouse lease agreement in Miami for 2 years with renewal option",
        "Professional cover letter for VP of Sales position at a fintech startup",
        "Non-disclosure agreement (NDA) between company and software freelancer",
        "Formal complaint to FTC about deceptive advertising practices",
        "Remote work and PTO policy for a 30-person technology company",
    ],
    "excel": [
        "Skincare product inventory with stock alerts and total value by category",
        "Monthly family budget with fixed vs variable expense chart",
        "Employee payroll with salary, health, pension deductions and net pay",
        "Sales dashboard with KPIs, month-over-month comparison and Q4 projections",
        "Client quotation for design services with tax, discount and payment terms",
    ],
    "powerpoint": [
        "Pitch deck for edtech startup, 10 slides for seed funding round",
        "Q3 results presentation for board of directors with growth charts",
        "Commercial proposal for financial consulting services for SMB",
        "Employee onboarding training, 12 slides covering values and culture",
        "ESG sustainability report for manufacturing company, 15 slides",
    ],
}



LANG = {
    "es": {
        "app_tag":     "Word · Excel · PowerPoint con IA",
        "login_title": "Iniciar sesión",
        "login_sub":   "Ingresa tus credenciales para continuar",
        "email_lbl":   "Correo electrónico",
        "pass_lbl":    "Contraseña",
        "email_ph":    "usuario@empresa.com",
        "pass_ph":     "Contraseña",
        "forgot":      "¿Olvidaste tu contraseña?",
        "login_btn":   "Entrar a WEP AI",
        "or":          "o",
        "create_btn":  "Crear cuenta nueva",
        "footer":      "WEP AI · Word · Excel · PowerPoint con IA  |  v1.0",
        # Registro
        "reg_title":   "Crear cuenta nueva",
        "reg_sub":     "Completa tu información y elige tu plan en un solo paso",
        "sec1":        "1  Información personal",
        "sec2":        "2  Elige tu plan",
        "name_lbl":    "Nombre",
        "last_lbl":    "Apellido",
        "name_ph":     "Juan",
        "last_ph":     "Vásquez",
        "pass2_ph":    "Mínimo 8 caracteres",
        "conf_lbl":    "Confirmar contraseña",
        "conf_ph":     "Repite tu contraseña",
        "continue_btn":"Crear mi cuenta",
        "back_btn":    "← Volver al login",
        "have_acc":    "¿Ya tienes cuenta? Iniciar sesión",
        "secure":      "🔒  Pago seguro · Sin compromiso · Cancela cuando quieras",
        # v14.6 — T&C y Privacy
        "terms_check": "He leído y acepto los",
        "terms_link":  "Términos",
        "terms_and":   "y la",
        "privacy_link":"Política de Privacidad",
        "terms_err":   "⚠ Debes aceptar los Términos y la Política de Privacidad para continuar",
        # Planes
        "free_name":   "Personal Free",
        "pro_name":    "Personal Pro",
        "biz_name":    "Enterprise",
        "free_price":  "$0  —  Gratis para siempre",
        "pro_price":   "$4.99 / mes",
        "biz_price":   "$24.99 / mes · hasta 10 usuarios",
        "free_feats":  ["⚡ 5 documentos por mes", "✓ Word, Excel, PowerPoint", "⚡ Historial 7 días", "✗ Sin correcciones IA", "✗ Sin soporte"],
        "pro_feats":   ["✓ Documentos ilimitados", "✓ Word, Excel, PowerPoint", "✓ Historial 90 días", "✓ Correcciones IA",    "✓ Soporte por email"],
        "biz_feats":   ["✓ Documentos ilimitados", "✓ Word, Excel, PowerPoint", "✓ Historial 1 año",   "✓ Correcciones IA",    "✓ Soporte 24/7"],
        "btn_free":    "Comenzar gratis",
        "btn_pro":     "Crear cuenta Pro",
        "btn_biz":     "Crear cuenta Business",
        "recommended": "⭐ Popular",
        "enterprise":  "🏢 Empresa",
        # Confirmación
        "ok_title":    "¡Bienvenido/a a WEP AI!",
        "ok_sub":      "Tu cuenta está lista. Entrando a la plataforma...",
        # App
        "new_conv":    "+ Nueva conversación",
        "convs":       "CONVERSACIONES",
        "active_model":"Modelo activo",
        "preview":     "Vista previa",
        "preview_sub": "El documento abierto\naparecerá aquí",
        "open_in":     "Abrir en Office",
        "profile":     "Mi perfil",
        "settings":    "Configuración",
        "my_docs":     "Mis documentos",
        "cancel_sub":  "Cancelar suscripción",
        "logout":      "Cerrar sesión",
        "welcome_t":   "¿En qué te puedo ayudar hoy?",
        "welcome_s":   "Descríbeme lo que necesitas con todos los detalles.\nPuedo crear documentos Word, hojas Excel y presentaciones PowerPoint.",
        "inp_ph":      "Escríbeme lo que necesitas con todos los detalles...",
        "cmd_hint":    "Cmd+Enter para enviar",
        "ready":       "Listo",
        "thinking":    "Analizando...",
        "generating":  "Generando documento...",
        "opening":     "Abriendo en Office...",
        "capturing":   "Verificando resultado...",
        "presenting":  "Esperando tu feedback",
        "correcting":  "Aplicando correcciones...",
        "feedback_q":  "¿Cómo lo ves? ¿Deseas algún cambio?",
        "doc_ready":   "listo y abierto en tu Mac",
        "no_doc":      "No hay documento generado aún.",
        "limit_title": "⚠ Límite del plan gratuito",
        "limit_msg":   "Has usado {} de 5 documentos este mes.\nActualiza a Pro para acceso ilimitado.",
        "limit_reached":"Has alcanzado el límite de 5 documentos del plan gratuito.\nActualiza a Pro para continuar.",
        "upgrade_btn": "Actualizar a Pro — $4.99/mes",
        "confirm_logout":"¿Deseas cerrar sesión?",
        "err_fields":  "Por favor completa todos los campos.",
        "err_email":   "Ingresa un correo electrónico válido.",
        "err_pass":    "La contraseña debe tener al menos 8 caracteres.",
        "err_match":   "Las contraseñas no coinciden.",
        "err_exists":  "Este correo ya está registrado.",
        "err_login":   "Correo o contraseña incorrectos.",
        "err_api":     "Error al conectar con la IA",
        "err_gen":     "Error generando el documento",
        "pw_weak":     "Débil", "pw_med": "Media", "pw_str": "Fuerte", "pw_vstr": "Muy fuerte",
    },
    "en": {
        "app_tag":     "Word · Excel · PowerPoint with AI",
        "login_title": "Sign in",
        "login_sub":   "Enter your credentials to continue",
        "email_lbl":   "Email",
        "pass_lbl":    "Password",
        "email_ph":    "user@company.com",
        "pass_ph":     "Password",
        "forgot":      "Forgot your password?",
        "login_btn":   "Sign in to WEP AI",
        "or":          "or",
        "create_btn":  "Create new account",
        "footer":      "WEP AI · Word · Excel · PowerPoint with AI  |  v1.0",
        "reg_title":   "Create new account",
        "reg_sub":     "Complete your information and choose your plan in one step",
        "sec1":        "1  Personal information",
        "sec2":        "2  Choose your plan",
        "name_lbl":    "First name",
        "last_lbl":    "Last name",
        "name_ph":     "John",
        "last_ph":     "Smith",
        "pass2_ph":    "Minimum 8 characters",
        "conf_lbl":    "Confirm password",
        "conf_ph":     "Repeat your password",
        "continue_btn":"Create my account",
        "back_btn":    "← Back to login",
        "have_acc":    "Already have an account? Sign in",
        "secure":      "🔒  Secure payment · No commitment · Cancel anytime",
        # v14.6 — T&C y Privacy
        "terms_check": "I have read and agree to the",
        "terms_link":  "Terms",
        "terms_and":   "and",
        "privacy_link":"Privacy Policy",
        "terms_err":   "⚠ You must accept the Terms and Privacy Policy to continue",
        "free_name":   "Personal Free",
        "pro_name":    "Personal Pro",
        "biz_name":    "Enterprise",
        "free_price":  "$0  —  Free forever",
        "pro_price":   "$4.99 / month",
        "biz_price":   "$24.99 / month · up to 10 users",
        "free_feats":  ["⚡ 5 documents/month", "✓ Word, Excel, PowerPoint", "⚡ 7-day history", "✗ No AI corrections", "✗ No support"],
        "pro_feats":   ["✓ Unlimited documents", "✓ Word, Excel, PowerPoint", "✓ 90-day history", "✓ AI corrections",    "✓ Email support"],
        "biz_feats":   ["✓ Unlimited documents", "✓ Word, Excel, PowerPoint", "✓ 1-year history",  "✓ AI corrections",    "✓ Priority support 24/7"],
        "btn_free":    "Start for free",
        "btn_pro":     "Create Pro account",
        "btn_biz":     "Create Business account",
        "recommended": "⭐ Popular",
        "enterprise":  "🏢 Business",
        "ok_title":    "Welcome to WEP AI!",
        "ok_sub":      "Your account is ready. Entering the platform...",
        "new_conv":    "+ New conversation",
        "convs":       "CONVERSATIONS",
        "active_model":"Active model",
        "preview":     "Preview",
        "preview_sub": "The open document\nwill appear here",
        "open_in":     "Open in Office",
        "profile":     "My profile",
        "settings":    "Settings",
        "my_docs":     "My documents",
        "cancel_sub":  "Cancel subscription",
        "logout":      "Sign out",
        "welcome_t":   "How can I help you today?",
        "welcome_s":   "Describe what you need in as much detail as you want.\nI can create Word documents, Excel sheets and PowerPoint presentations.",
        "inp_ph":      "Tell me what you need with as much detail as you want...",
        "cmd_hint":    "Cmd+Enter to send",
        "ready":       "Ready",
        "thinking":    "Analyzing...",
        "generating":  "Generating document...",
        "opening":     "Opening in Office...",
        "capturing":   "Verifying result...",
        "presenting":  "Waiting for feedback",
        "correcting":  "Applying corrections...",
        "feedback_q":  "How does it look? Any changes you'd like?",
        "doc_ready":   "ready and open on your Mac",
        "no_doc":      "No document generated yet.",
        "limit_title": "⚠ Free plan limit",
        "limit_msg":   "You've used {} of 5 documents this month.\nUpgrade to Pro for unlimited access.",
        "limit_reached":"You've reached the 5-document limit.\nUpgrade to Pro to continue.",
        "upgrade_btn": "Upgrade to Pro — $4.99/month",
        "confirm_logout":"Do you want to sign out?",
        "err_fields":  "Please fill in all fields.",
        "err_email":   "Please enter a valid email address.",
        "err_pass":    "Password must be at least 8 characters.",
        "err_match":   "Passwords do not match.",
        "err_exists":  "This email is already registered.",
        "err_login":   "Incorrect email or password.",
        "err_api":     "Error connecting to AI",
        "err_gen":     "Error generating document",
        "pw_weak":     "Weak", "pw_med": "Medium", "pw_str": "Strong", "pw_vstr": "Very strong",
    }
}

def T(lang, key): return LANG[lang].get(key, key)


# v14.8 — Helper para enviar emails en background sin bloquear el thread UI.
# Antes (v14.7), send_welcome se llamaba síncrono en _do_register y
# _open_payment_stripe; con EMAIL_PROVIDER configurado y red lenta podía
# congelar la app hasta 10s (timeout HTTP del cliente del provider).
def _send_welcome_async(email_to: str, name: str, lang: str):
    """Lanza send_welcome en un thread daemon. No bloquea, no propaga errores."""
    def _run():
        try:
            from storage import email as _email
            _email.send_welcome(email_to, name, lang=lang)
        except Exception as e:
            print(f"[WEP AI] send_welcome failed (non-blocking): {e}")
    threading.Thread(target=_run, daemon=True).start()


# ══════════════════════════════════════════════════════════════════════════════
#  VENTANA 1 — LOGIN
# ══════════════════════════════════════════════════════════════════════════════
class LoginWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("WEP AI")
        self.geometry("500x660")
        self.resizable(False, False)
        self.configure(fg_color=BG)
        self.lang  = ["es"]
        self.result = None
        self._build()

    def _t(self, k): return T(self.lang[0], k)

    def _build(self):
        # LOGO
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(pady=(38, 4))
        row = ctk.CTkFrame(top, fg_color="transparent"); row.pack()
        for ch, col in [("W", WORD_C), ("E", EXCEL_C), ("P", PPT_C)]:
            ctk.CTkLabel(row, text=ch, font=ctk.CTkFont("Arial", 52, "bold"),
                         text_color=col).pack(side="left")
        ctk.CTkLabel(row, text=" AI", font=ctk.CTkFont("Arial", 32, "normal"),
                     text_color=WHITE).pack(side="left")

        self.apps_row = ctk.CTkFrame(top, fg_color="transparent"); self.apps_row.pack(pady=(2,0))
        for nm, col in [("Word", WORD_C), ("  ·  ", GRAY2), ("Excel", EXCEL_C),
                         ("  ·  ", GRAY2), ("PowerPoint", PPT_C)]:
            ctk.CTkLabel(self.apps_row, text=nm,
                         font=ctk.CTkFont("Arial", 12, "bold" if nm.strip() else "normal"),
                         text_color=col).pack(side="left")
        self.tag_lbl = ctk.CTkLabel(top, text=self._t("app_tag"),
                                    font=ctk.CTkFont("Arial", 11), text_color=GRAY2)
        self.tag_lbl.pack(pady=(2, 0))

        # LANG TOGGLE
        lt = ctk.CTkFrame(self, fg_color=CARD, corner_radius=22,
                          border_width=1, border_color=BORDER)
        lt.pack(pady=(14, 18))
        self.b_es = ctk.CTkButton(lt, text="🇪🇸  Español", width=120, height=34,
                                   corner_radius=19, font=ctk.CTkFont("Arial", 12, "bold"),
                                   fg_color=BLUE, text_color=WHITE, hover_color=BLUE2,
                                   command=lambda: self._lang("es"))
        self.b_es.pack(side="left", padx=3, pady=3)
        self.b_en = ctk.CTkButton(lt, text="🇺🇸  English", width=120, height=34,
                                   corner_radius=19, font=ctk.CTkFont("Arial", 12, "bold"),
                                   fg_color="transparent", text_color=GRAY2, hover_color=CARD2,
                                   command=lambda: self._lang("en"))
        self.b_en.pack(side="left", padx=3, pady=3)

        # CARD
        card = ctk.CTkFrame(self, fg_color=CARD, corner_radius=18,
                             border_width=1, border_color=BORDER)
        card.pack(padx=32, fill="x")
        inn = ctk.CTkFrame(card, fg_color="transparent")
        inn.pack(padx=28, pady=(24, 18), fill="x")

        self.lt_lbl = ctk.CTkLabel(inn, text=self._t("login_title"),
                                   font=ctk.CTkFont("Arial", 22, "bold"), text_color=WHITE)
        self.lt_lbl.pack(pady=(0, 3))
        self.ls_lbl = ctk.CTkLabel(inn, text=self._t("login_sub"),
                                   font=ctk.CTkFont("Arial", 13), text_color=GRAY1)
        self.ls_lbl.pack(pady=(0, 20))

        self.el_lbl = ctk.CTkLabel(inn, text=self._t("email_lbl"),
                                   font=ctk.CTkFont("Arial", 12, "bold"),
                                   text_color=GRAY1, anchor="w")
        self.el_lbl.pack(fill="x", pady=(0, 5))
        self.email_e = ctk.CTkEntry(inn, height=46, corner_radius=12,
                                    font=ctk.CTkFont("Arial", 15),
                                    placeholder_text=self._t("email_ph"),
                                    fg_color=CARD2, border_color=BORDER,
                                    border_width=1, text_color=WHITE)
        self.email_e.pack(fill="x", pady=(0, 12))

        self.pl_lbl = ctk.CTkLabel(inn, text=self._t("pass_lbl"),
                                   font=ctk.CTkFont("Arial", 12, "bold"),
                                   text_color=GRAY1, anchor="w")
        self.pl_lbl.pack(fill="x", pady=(0, 5))
        self.pass_e = ctk.CTkEntry(inn, height=46, corner_radius=12, show="●",
                                   font=ctk.CTkFont("Arial", 15),
                                   placeholder_text=self._t("pass_ph"),
                                   fg_color=CARD2, border_color=BORDER,
                                   border_width=1, text_color=WHITE)
        self.pass_e.pack(fill="x", pady=(0, 4))
        self.pass_e.bind("<Return>", lambda e: self._login())

        self.fg_lbl = ctk.CTkLabel(inn, text=self._t("forgot"),
                                   font=ctk.CTkFont("Arial", 11),
                                   text_color=BLUE, cursor="hand2", anchor="e")
        self.fg_lbl.pack(fill="x", pady=(0, 18))

        self.login_btn = ctk.CTkButton(inn, text=self._t("login_btn"),
                                       height=48, corner_radius=12,
                                       font=ctk.CTkFont("Arial", 15, "bold"),
                                       fg_color=BLUE, hover_color=BLUE2,
                                       command=self._login)
        self.login_btn.pack(fill="x", pady=(0, 12))

        sep = ctk.CTkFrame(inn, fg_color="transparent"); sep.pack(fill="x", pady=(0, 12))
        ctk.CTkFrame(sep, height=1, fg_color=BORDER).pack(
            side="left", fill="x", expand=True, padx=(0, 10))
        self.or_lbl = ctk.CTkLabel(sep, text=self._t("or"),
                                   font=ctk.CTkFont("Arial", 11), text_color=GRAY2)
        self.or_lbl.pack(side="left")
        ctk.CTkFrame(sep, height=1, fg_color=BORDER).pack(
            side="left", fill="x", expand=True, padx=(10, 0))

        self.create_btn = ctk.CTkButton(inn, text=self._t("create_btn"),
                                        height=46, corner_radius=12,
                                        font=ctk.CTkFont("Arial", 14, "bold"),
                                        fg_color="transparent",
                                        border_width=1, border_color=BORDER,
                                        text_color=GRAY1, hover_color=CARD2,
                                        command=self._open_register)
        self.create_btn.pack(fill="x")

        self.footer_lbl = ctk.CTkLabel(self, text=self._t("footer"),
                                       font=ctk.CTkFont("Arial", 9), text_color=GRAY3)
        self.footer_lbl.pack(pady=(14, 8))

    def _lang(self, l):
        self.lang[0] = l
        self.b_es.configure(fg_color=BLUE  if l == "es" else "transparent",
                            text_color=WHITE if l == "es" else GRAY2)
        self.b_en.configure(fg_color=BLUE  if l == "en" else "transparent",
                            text_color=WHITE if l == "en" else GRAY2)
        for attr, key in [("lt_lbl","login_title"),("ls_lbl","login_sub"),
                           ("el_lbl","email_lbl"),("pl_lbl","pass_lbl"),
                           ("fg_lbl","forgot"),("login_btn","login_btn"),
                           ("or_lbl","or"),("create_btn","create_btn"),
                           ("footer_lbl","footer"),("tag_lbl","app_tag")]:
            getattr(self, attr).configure(text=self._t(key))
        self.email_e.configure(placeholder_text=self._t("email_ph"))
        self.pass_e.configure(placeholder_text=self._t("pass_ph"))

    def _login(self):
        u = self.email_e.get().strip()
        p = self.pass_e.get().strip()

        # Campos vacíos
        if not u or not p:
            self._shake_error(self._t("err_fields"))
            return

        # Validar formato de email
        if not re.match(r"^[^\s@]+@[^\s@]+\.[^\s@]+$", u):
            self._shake_error(self._t("err_email"))
            self.email_e.configure(border_color=RED)
            return

        # Intentos fallidos
        if not hasattr(self, "_attempts"):
            self._attempts = 0

        user = db.login_user(u, p)

        if not user:
            self._attempts += 1
            self.email_e.configure(border_color=RED)
            self.pass_e.configure(border_color=RED)

            if self._attempts >= 3:
                # Bloquear 30 segundos después de 3 intentos fallidos
                msg = ("Demasiados intentos fallidos.\nEspera 30 segundos antes de intentar de nuevo."
                       if self.lang[0] == "es"
                       else "Too many failed attempts.\nWait 30 seconds before trying again.")
                messagebox.showerror("WEP AI — Acceso denegado", msg)
                self.login_btn.configure(state="disabled",
                    text="Espera 30s..." if self.lang[0]=="es" else "Wait 30s...")
                self.after(30000, lambda: (
                    self.login_btn.configure(state="normal", text=self._t("login_btn")),
                    setattr(self, "_attempts", 0)
                ))
            else:
                remaining = 3 - self._attempts
                hint = (f"Correo o contraseña incorrectos. ({remaining} intento{'s' if remaining!=1 else ''} restante{'s' if remaining!=1 else ''})"
                        if self.lang[0]=="es"
                        else f"Incorrect email or password. ({remaining} attempt{'s' if remaining!=1 else ''} remaining)")
                self._shake_error(hint)
            return

        # Login exitoso — resetear intentos
        self._attempts = 0
        self.email_e.configure(border_color=BORDER)
        self.pass_e.configure(border_color=BORDER)
        self.result = (user, self.lang[0])
        self.destroy()

    def _shake_error(self, msg):
        """Muestra mensaje de error con animación visual."""
        if hasattr(self, "_err_lbl") and self._err_lbl.winfo_exists():
            self._err_lbl.configure(text=msg)
        else:
            # Crear label de error si no existe
            self._err_lbl = ctk.CTkLabel(
                self, text=msg,
                font=ctk.CTkFont("Arial", 11),
                text_color=RED,
                fg_color="#1C0A0A",
                corner_radius=7,
                wraplength=330
            )
            self._err_lbl.pack(pady=(4, 0), padx=32, fill="x")
        # Efecto visual en campos
        self.email_e.configure(border_color=RED)
        self.pass_e.configure(border_color=RED)
        # Restaurar borde después de 3 segundos
        self.after(3000, lambda: (
            self.email_e.configure(border_color=BORDER) if self.email_e.winfo_exists() else None,
            self.pass_e.configure(border_color=BORDER) if self.pass_e.winfo_exists() else None,
        ))

    def _open_register(self):
        self.withdraw()
        reg = RegisterWindow(self.lang[0],
                             on_success=self._on_reg_ok,
                             on_cancel=self._on_reg_cancel)
        reg.mainloop()

    def _on_reg_ok(self, user, lang):
        self.result = (user, lang)
        self.destroy()

    def _on_reg_cancel(self, lang):
        self.lang[0] = lang
        self._lang(lang)
        self.deiconify()


# ══════════════════════════════════════════════════════════════════════════════
#  VENTANA 2 — REGISTRO (info + plan en UNA sola ventana)
# ══════════════════════════════════════════════════════════════════════════════
class RegisterWindow(ctk.CTk):
    def __init__(self, lang="es", on_success=None, on_cancel=None):
        super().__init__()
        self.title("WEP AI — Crear cuenta")
        self.geometry("560x820")
        self.resizable(False, True)
        self.configure(fg_color=BG)
        self.lang       = [lang]
        self.on_success = on_success
        self.on_cancel  = on_cancel
        self.sel_plan   = "pro"
        self.plan_frames = {}
        self._build()

    def _t(self, k): return T(self.lang[0], k)

    def _build(self):
        # Scrollable container
        scroll = ctk.CTkScrollableFrame(self, fg_color=BG, corner_radius=0)
        scroll.pack(fill="both", expand=True)
        self._scroll = scroll
        self._fill(scroll)

    def _fill(self, parent):
        # LOGO pequeño
        top = ctk.CTkFrame(parent, fg_color="transparent"); top.pack(pady=(24, 2))
        row = ctk.CTkFrame(top, fg_color="transparent"); row.pack()
        for ch, col in [("W", WORD_C), ("E", EXCEL_C), ("P", PPT_C)]:
            ctk.CTkLabel(row, text=ch, font=ctk.CTkFont("Arial", 34, "bold"),
                         text_color=col).pack(side="left")
        ctk.CTkLabel(row, text=" AI", font=ctk.CTkFont("Arial", 22, "normal"),
                     text_color=WHITE).pack(side="left")
        self.tag2 = ctk.CTkLabel(parent, text=self._t("app_tag"),
                                 font=ctk.CTkFont("Arial", 10), text_color=GRAY2)
        self.tag2.pack(pady=(0, 6))

        # LANG TOGGLE
        lt = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=20,
                          border_width=1, border_color=BORDER)
        lt.pack(pady=(0, 14))
        self.r_es = ctk.CTkButton(lt, text="🇪🇸  Español", width=110, height=30,
                                   corner_radius=17, font=ctk.CTkFont("Arial", 11, "bold"),
                                   fg_color=BLUE if self.lang[0]=="es" else "transparent",
                                   text_color=WHITE if self.lang[0]=="es" else GRAY2,
                                   hover_color=CARD2,
                                   command=lambda: self._lang("es"))
        self.r_es.pack(side="left", padx=3, pady=3)
        self.r_en = ctk.CTkButton(lt, text="🇺🇸  English", width=110, height=30,
                                   corner_radius=17, font=ctk.CTkFont("Arial", 11, "bold"),
                                   fg_color=BLUE if self.lang[0]=="en" else "transparent",
                                   text_color=WHITE if self.lang[0]=="en" else GRAY2,
                                   hover_color=CARD2,
                                   command=lambda: self._lang("en"))
        self.r_en.pack(side="left", padx=3, pady=3)

        # CARD PRINCIPAL
        card = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=18,
                             border_width=1, border_color=BORDER)
        card.pack(padx=24, fill="x", pady=(0, 8))
        inn = ctk.CTkFrame(card, fg_color="transparent")
        inn.pack(padx=26, pady=(22, 18), fill="x")

        self.reg_title = ctk.CTkLabel(inn, text=self._t("reg_title"),
                                      font=ctk.CTkFont("Arial", 20, "bold"), text_color=WHITE)
        self.reg_title.pack(pady=(0, 3))
        self.reg_sub = ctk.CTkLabel(inn, text=self._t("reg_sub"),
                                    font=ctk.CTkFont("Arial", 12), text_color=GRAY1)
        self.reg_sub.pack(pady=(0, 20))

        # ── SECCIÓN 1: INFO PERSONAL ──────────────────────────────────────────
        self.sec1_lbl = ctk.CTkLabel(inn, text=self._t("sec1"),
                                     font=ctk.CTkFont("Arial", 13, "bold"),
                                     text_color=BLUE, anchor="w")
        self.sec1_lbl.pack(fill="x", pady=(0, 10))

        self.err_lbl = ctk.CTkLabel(inn, text="",
                                    font=ctk.CTkFont("Arial", 11),
                                    text_color=RED, wraplength=480)
        self.err_lbl.pack(fill="x", pady=(0, 6))

        # Nombre + Apellido en dos columnas
        row2 = ctk.CTkFrame(inn, fg_color="transparent")
        row2.pack(fill="x", pady=(0, 10))
        row2.grid_columnconfigure((0, 1), weight=1)

        self.nl_lbl = ctk.CTkLabel(row2, text=self._t("name_lbl"),
                                   font=ctk.CTkFont("Arial", 11, "bold"),
                                   text_color=GRAY1, anchor="w")
        self.nl_lbl.grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.ll_lbl = ctk.CTkLabel(row2, text=self._t("last_lbl"),
                                   font=ctk.CTkFont("Arial", 11, "bold"),
                                   text_color=GRAY1, anchor="w")
        self.ll_lbl.grid(row=0, column=1, sticky="w")

        self.name_e = ctk.CTkEntry(row2, height=42, corner_radius=10,
                                   font=ctk.CTkFont("Arial", 13),
                                   placeholder_text=self._t("name_ph"),
                                   fg_color=CARD2, border_color=BORDER,
                                   border_width=1, text_color=WHITE)
        self.name_e.grid(row=1, column=0, sticky="ew", padx=(0, 8), pady=(4, 0))
        self.last_e = ctk.CTkEntry(row2, height=42, corner_radius=10,
                                   font=ctk.CTkFont("Arial", 13),
                                   placeholder_text=self._t("last_ph"),
                                   fg_color=CARD2, border_color=BORDER,
                                   border_width=1, text_color=WHITE)
        self.last_e.grid(row=1, column=1, sticky="ew", pady=(4, 0))

        # Email / Pass / Confirm
        for lbl_k, attr, ph_k, show in [
            ("email_lbl", "email2_e", "email_ph",  None),
            ("pass_lbl",  "pass2_e",  "pass2_ph",  "●"),
            ("conf_lbl",  "conf_e",   "conf_ph",   "●"),
        ]:
            l = ctk.CTkLabel(inn, text=self._t(lbl_k),
                             font=ctk.CTkFont("Arial", 11, "bold"),
                             text_color=GRAY1, anchor="w")
            l.pack(fill="x", pady=(8, 4))
            setattr(self, f"_lbl_{attr}", l)
            e = ctk.CTkEntry(inn, height=42, corner_radius=10,
                             font=ctk.CTkFont("Arial", 13),
                             placeholder_text=self._t(ph_k),
                             fg_color=CARD2, border_color=BORDER,
                             border_width=1, text_color=WHITE)
            if show: e.configure(show=show)
            if attr == "pass2_e":
                e.bind("<KeyRelease>", lambda ev: self._pw_strength())
            e.pack(fill="x")
            setattr(self, attr, e)

        # Barra de fuerza de contraseña
        self.str_bar_bg = ctk.CTkFrame(inn, height=4, corner_radius=2, fg_color=CARD2)
        self.str_bar_bg.pack(fill="x", pady=(6, 0))
        self.str_bar = ctk.CTkFrame(self.str_bar_bg, height=4, corner_radius=2,
                                    fg_color=BLUE, width=0)
        self.str_bar.place(x=0, y=0, relheight=1, relwidth=0)
        self.str_lbl = ctk.CTkLabel(inn, text="", font=ctk.CTkFont("Arial", 10),
                                    text_color=GRAY1, anchor="w")
        self.str_lbl.pack(fill="x", pady=(3, 10))

        # ── DIVISOR ──────────────────────────────────────────────────────────
        ctk.CTkFrame(inn, height=1, fg_color=BORDER).pack(fill="x", pady=(8, 18))

        # ── SECCIÓN 2: PLANES ─────────────────────────────────────────────────
        self.sec2_lbl = ctk.CTkLabel(inn, text=self._t("sec2"),
                                     font=ctk.CTkFont("Arial", 13, "bold"),
                                     text_color=AMBER, anchor="w")
        self.sec2_lbl.pack(fill="x", pady=(0, 12))

        plan_frame = ctk.CTkFrame(inn, fg_color="transparent")
        plan_frame.pack(fill="x", pady=(0, 16))
        plan_frame.grid_columnconfigure((0, 1, 2), weight=1)

        plans = [
            ("free", GRAY2,   BORDER, self._t("free_name"), self._t("free_price"), self._t("free_feats")),
            ("pro",  BLUE,    BLUE,   self._t("pro_name"),  self._t("pro_price"),  self._t("pro_feats")),
            ("biz",  PURPLE2, PURPLE, self._t("biz_name"),  self._t("biz_price"),  self._t("biz_feats")),
        ]
        self.plan_frames = {}
        for i, (pid, price_col, sel_col, name, price, feats) in enumerate(plans):
            is_sel = (pid == self.sel_plan)
            pf = ctk.CTkFrame(plan_frame, fg_color=CARD2 if is_sel else CARD,
                              corner_radius=10,
                              border_width=2 if is_sel else 1,
                              border_color=sel_col if is_sel else BORDER)
            pf.grid(row=0, column=i, padx=5, sticky="nsew")
            pf.bind("<Button-1>", lambda e, p=pid: self._sel_plan(p))

            # Badge
            badge_txt = self._t("recommended") if pid == "pro" else (
                        self._t("enterprise") if pid == "biz" else "")
            if badge_txt:
                b = ctk.CTkLabel(pf, text=badge_txt,
                                 font=ctk.CTkFont("Arial", 9, "bold"),
                                 text_color=WHITE,
                                 fg_color=BLUE if pid == "pro" else PURPLE,
                                 corner_radius=6)
                b.pack(padx=6, pady=(6, 2), anchor="e")
                b.bind("<Button-1>", lambda e, p=pid: self._sel_plan(p))

            # Plan name
            nm_lbl = ctk.CTkLabel(pf, text=name, font=ctk.CTkFont("Arial", 11, "bold"),
                                  text_color=WHITE)
            nm_lbl.pack(padx=8, pady=(6 if not badge_txt else 2, 2), anchor="w")
            nm_lbl.bind("<Button-1>", lambda e, p=pid: self._sel_plan(p))

            # Precio
            pc_lbl = ctk.CTkLabel(pf, text=price, font=ctk.CTkFont("Arial", 10),
                                  text_color=price_col, wraplength=130)
            pc_lbl.pack(padx=8, pady=(0, 6), anchor="w")
            pc_lbl.bind("<Button-1>", lambda e, p=pid: self._sel_plan(p))

            # Features
            for feat in feats:
                fc = GRAY2 if feat.startswith("✗") else (AMBER if feat.startswith("⚡") else GRAY1)
                fl = ctk.CTkLabel(pf, text=feat, font=ctk.CTkFont("Arial", 9),
                                  text_color=fc, anchor="w", wraplength=130)
                fl.pack(padx=8, pady=1, fill="x")
                fl.bind("<Button-1>", lambda e, p=pid: self._sel_plan(p))
            # Último padding
            ctk.CTkFrame(pf, height=8, fg_color="transparent").pack()
            self.plan_frames[pid] = pf

        # ── BOTÓN PRINCIPAL ───────────────────────────────────────────────────
        btn_k = {"free": "btn_free", "pro": "btn_pro", "biz": "btn_biz"}
        # ── v14.6 — Checkbox de aceptación de Términos y Privacidad ─────────
        # Se monta JUSTO antes del botón principal para que el usuario los
        # vea antes de continuar. Si está desmarcado, _create() bloquea.
        terms_frame = ctk.CTkFrame(inn, fg_color="transparent")
        terms_frame.pack(fill="x", pady=(8, 6))

        self.terms_var = ctk.BooleanVar(value=False)
        self.terms_check = ctk.CTkCheckBox(
            terms_frame, text="", variable=self.terms_var,
            checkbox_width=18, checkbox_height=18,
            border_color=BORDER, fg_color=BLUE, hover_color=BLUE2)
        self.terms_check.pack(side="left", padx=(0, 6))

        # Etiqueta + links inline
        terms_text_frame = ctk.CTkFrame(terms_frame, fg_color="transparent")
        terms_text_frame.pack(side="left", fill="x", expand=True)
        self.terms_lbl = ctk.CTkLabel(
            terms_text_frame, text=self._t("terms_check"),
            font=ctk.CTkFont("Arial", 11), text_color=GRAY1)
        self.terms_lbl.pack(side="left")
        self.terms_link_lbl = ctk.CTkLabel(
            terms_text_frame, text=" " + self._t("terms_link"),
            font=ctk.CTkFont("Arial", 11, underline=True),
            text_color=BLUE, cursor="hand2")
        self.terms_link_lbl.pack(side="left")
        self.terms_link_lbl.bind("<Button-1>",
                                  lambda e: self._open_terms("terms"))
        self.terms_and_lbl = ctk.CTkLabel(
            terms_text_frame, text=" " + self._t("terms_and"),
            font=ctk.CTkFont("Arial", 11), text_color=GRAY1)
        self.terms_and_lbl.pack(side="left")
        self.privacy_link_lbl = ctk.CTkLabel(
            terms_text_frame, text=" " + self._t("privacy_link"),
            font=ctk.CTkFont("Arial", 11, underline=True),
            text_color=BLUE, cursor="hand2")
        self.privacy_link_lbl.pack(side="left")
        self.privacy_link_lbl.bind("<Button-1>",
                                    lambda e: self._open_terms("privacy"))

        # Mensaje de error específico de T&C (oculto por default)
        self.terms_err_lbl = ctk.CTkLabel(
            inn, text="", font=ctk.CTkFont("Arial", 10),
            text_color=AMBER)
        self.terms_err_lbl.pack(pady=(0, 2))

        self.main_btn = ctk.CTkButton(
            inn, text=self._t(btn_k[self.sel_plan]),
            height=48, corner_radius=12,
            font=ctk.CTkFont("Arial", 14, "bold"),
            fg_color=PURPLE if self.sel_plan == "biz" else BLUE,
            hover_color=PURPLE2 if self.sel_plan == "biz" else BLUE2,
            command=self._create)
        self.main_btn.pack(fill="x", pady=(0, 8))

        self.secure_lbl = ctk.CTkLabel(inn, text=self._t("secure"),
                                       font=ctk.CTkFont("Arial", 10), text_color=GRAY2)
        self.secure_lbl.pack(pady=(0, 4))

        # Volver / ya tengo cuenta
        self.back_btn_w = ctk.CTkButton(
            inn, text=self._t("back_btn"),
            height=40, corner_radius=10,
            font=ctk.CTkFont("Arial", 12, "bold"),
            fg_color="transparent", border_width=1, border_color=BORDER,
            text_color=GRAY1, hover_color=CARD2,
            command=self._go_back)
        self.back_btn_w.pack(fill="x", pady=(0, 6))

        self.have_lbl = ctk.CTkLabel(inn, text=self._t("have_acc"),
                                     font=ctk.CTkFont("Arial", 11),
                                     text_color=BLUE, cursor="hand2")
        self.have_lbl.pack(pady=(0, 4))
        self.have_lbl.bind("<Button-1>", lambda e: self._go_back())

    def _lang(self, l):
        self.lang[0] = l
        self.r_es.configure(fg_color=BLUE if l=="es" else "transparent",
                            text_color=WHITE if l=="es" else GRAY2)
        self.r_en.configure(fg_color=BLUE if l=="en" else "transparent",
                            text_color=WHITE if l=="en" else GRAY2)
        for attr, key in [
            ("tag2","app_tag"),("reg_title","reg_title"),("reg_sub","reg_sub"),
            ("sec1_lbl","sec1"),("sec2_lbl","sec2"),
            ("nl_lbl","name_lbl"),("ll_lbl","last_lbl"),("back_btn_w","back_btn"),
            ("have_lbl","have_acc"),("secure_lbl","secure"),
            # v14.6 — labels de T&C
            ("terms_lbl","terms_check"),
        ]:
            getattr(self, attr).configure(text=self._t(key))
        # Labels de T&C con espacio de prefijo o sufijo
        self.terms_link_lbl.configure(text=" " + self._t("terms_link"))
        self.terms_and_lbl.configure(text=" " + self._t("terms_and"))
        self.privacy_link_lbl.configure(text=" " + self._t("privacy_link"))
        self.name_e.configure(placeholder_text=self._t("name_ph"))
        self.last_e.configure(placeholder_text=self._t("last_ph"))
        self.email2_e.configure(placeholder_text=self._t("email_ph"))
        self.pass2_e.configure(placeholder_text=self._t("pass2_ph"))
        self.conf_e.configure(placeholder_text=self._t("conf_ph"))
        self._update_plan_btn()

    def _pw_strength(self):
        v = self.pass2_e.get()
        score = sum([len(v) >= 8, bool(re.search(r"[A-Z]", v)),
                     bool(re.search(r"[0-9]", v)), bool(re.search(r"[^A-Za-z0-9]", v))])
        colors = [RED, AMBER, BLUE, GREEN]
        keys   = ["pw_weak", "pw_med", "pw_str", "pw_vstr"]
        if v:
            self.str_bar.place(relwidth=score / 4)
            self.str_bar.configure(fg_color=colors[score-1] if score else RED)
            self.str_lbl.configure(text=self._t(keys[score-1]) if score else "",
                                   text_color=colors[score-1] if score else RED)
        else:
            self.str_bar.place(relwidth=0)
            self.str_lbl.configure(text="")

    def _sel_plan(self, plan):
        self.sel_plan = plan
        colors = {"free": BORDER, "pro": BLUE, "biz": PURPLE}
        bgs    = {"free": CARD,   "pro": CARD2, "biz": "#1a0f35"}
        for pid, pf in self.plan_frames.items():
            pf.configure(fg_color=bgs[pid] if pid == plan else CARD,
                         border_width=2 if pid == plan else 1,
                         border_color=colors[pid] if pid == plan else BORDER)
        self._update_plan_btn()

    def _update_plan_btn(self):
        btn_k = {"free": "btn_free", "pro": "btn_pro", "biz": "btn_biz"}
        self.main_btn.configure(
            text=self._t(btn_k[self.sel_plan]),
            fg_color=PURPLE if self.sel_plan == "biz" else BLUE,
            hover_color=PURPLE2 if self.sel_plan == "biz" else BLUE2)

    def _create(self):
        nm = self.name_e.get().strip()
        ls = self.last_e.get().strip()
        em = self.email2_e.get().strip()
        ps = self.pass2_e.get()
        cf = self.conf_e.get()
        if not all([nm, ls, em, ps, cf]):
            self.err_lbl.configure(text=self._t("err_fields")); return
        if not re.match(r"^[^\s@]+@[^\s@]+\.[^\s@]+$", em):
            self.err_lbl.configure(text=self._t("err_email")); return
        if len(ps) < 8:
            self.err_lbl.configure(text=self._t("err_pass")); return
        if ps != cf:
            self.err_lbl.configure(text=self._t("err_match")); return
        # v14.6 — exigir aceptación de Términos y Privacidad
        if not self.terms_var.get():
            self.terms_err_lbl.configure(text=self._t("terms_err"))
            return
        else:
            self.terms_err_lbl.configure(text="")

        # ── Si el plan requiere pago, abrir formulario de pago ──
        if self.sel_plan in ("pro", "biz"):
            self._open_payment(nm, ls, em, ps)
        else:
            # Plan Free: crear cuenta directamente
            self._do_register(nm, ls, em, ps)

    def _open_terms(self, kind: str):
        """v14.6 — Abre TERMS o PRIVACY en el viewer del sistema.
        Busca el archivo en la raíz del paquete según el idioma actual.
        Si no encuentra el archivo, abre la URL pública (placeholder)."""
        import os, webbrowser
        suffix = "EN" if self.lang[0] == "en" else "ES"
        filename = f"{'TERMS' if kind == 'terms' else 'PRIVACY'}_{suffix}.md"
        # Buscar el archivo subiendo desde ui/ hasta la raíz del paquete wepai/
        here = os.path.dirname(os.path.abspath(__file__))
        candidate_paths = [
            os.path.join(here, "..", filename),                      # wepai/
            os.path.join(os.path.dirname(here), filename),            # wepai/
            os.path.join(os.getcwd(), filename),                      # cwd
            os.path.expanduser(f"~/wepai/{filename}"),                # ~/wepai/
        ]
        for path in candidate_paths:
            if os.path.exists(path):
                # Convertir a URL file:// para que el navegador lo abra
                abs_path = os.path.abspath(path)
                webbrowser.open(f"file://{abs_path}")
                return
        # Fallback: URL pública del sitio (placeholder)
        public_url = (f"https://wepai.app/{kind}-{'en' if suffix == 'EN' else 'es'}")
        webbrowser.open(public_url)

    def _open_payment(self, nm, ls, em, ps):
        """Pago: si Stripe está configurado abre Checkout en navegador,
        si no, cae al modo demo (formulario simulado)."""
        # v14.5 — Path A: Stripe Checkout real
        try:
            from storage import stripe_client
        except ImportError:
            stripe_client = None
        if stripe_client and stripe_client.is_configured():
            self._open_payment_stripe(nm, ls, em, ps)
        else:
            self._open_payment_demo(nm, ls, em, ps)

    def _open_payment_stripe(self, nm, ls, em, ps):
        """v14.5 — flujo real con Stripe Checkout.
        1. Registra el usuario inicialmente como 'free' (no le damos el plan
           pagado hasta que Stripe confirme via webhook).
        2. Crea Checkout Session y abre el navegador del usuario.
        3. Muestra modal de espera con instrucciones.
        El webhook server (storage/webhook_server.py desplegado en la nube)
        se encarga de actualizar el plan en la DB cuando Stripe confirma."""
        from storage import stripe_client
        import webbrowser
        lbl = self.lang[0]

        # Registro como 'free' (pendiente de upgrade)
        uid, err = db.register_user(nm, ls, em, ps, "free")
        if err == "email_exists":
            self.err_lbl.configure(text="⚠ " + self._t("err_exists")); return
        if err:
            self.err_lbl.configure(text=f"Error: {err}"); return

        # v14.8 — Email de bienvenida en background daemon (no bloquea UI).
        # El email de "pago confirmado" lo envía el webhook server cuando
        # Stripe notifica. Si no hay EMAIL_PROVIDER, es no-op (dry-run).
        _send_welcome_async(em, nm, lang=lbl)

        session = stripe_client.create_checkout_session(
            self.sel_plan, em, uid,
            # v14.6.1 — Sin trial period: el plan Personal Free ya permite
            # probar el software gratis (5 docs/mes). Quien elige Pro/Enterprise
            # ya conoce el producto y va con intención de pago directo. Si en
            # el futuro se decide reactivar trial, basta con setear la env var
            # STRIPE_TRIAL_DAYS=N — el cliente Stripe lo respeta automáticamente.
            trial_days=0,
            # v14.7 — pasamos el idioma actual al metadata de Stripe para que
            # el webhook server use ese idioma al enviar emails transaccionales.
            user_lang=lbl,
        )
        if not session:
            # Stripe falló — caer al modo demo en lugar de bloquear
            self._open_payment_demo(nm, ls, em, ps)
            return

        webbrowser.open(session["url"])

        # Modal de espera
        wait = ctk.CTkToplevel(self)
        wait.title("WEP AI — Stripe Checkout")
        wait.geometry("460x320")
        wait.configure(fg_color=BG)
        wait.grab_set(); wait.lift(); wait.focus_force()
        wait.attributes("-topmost", True)

        ctk.CTkLabel(wait, text="🔒  Stripe Checkout",
                     font=ctk.CTkFont("Arial", 18, "bold"),
                     text_color=BLUE).pack(pady=(28, 6))
        ctk.CTkLabel(wait,
            text=("Completá el pago en tu navegador."
                  if lbl=="es" else
                  "Complete payment in your browser."),
            font=ctk.CTkFont("Arial", 13), text_color=WHITE).pack(pady=(0, 18))
        ctk.CTkLabel(wait,
            text=("Una vez confirmado el pago, tu suscripción se activará "
                  "automáticamente. Podés cerrar esta ventana y volver a "
                  "iniciar sesión en unos minutos."
                  if lbl=="es" else
                  "Once payment is confirmed, your subscription will activate "
                  "automatically. You can close this window and log in again "
                  "in a few minutes."),
            font=ctk.CTkFont("Arial", 11), text_color=GRAY1,
            wraplength=400, justify="center").pack(pady=(0, 22))

        def _close():
            wait.destroy()
            user = {"id": uid, "name": nm, "last_name": ls, "email": em,
                    "plan": "free", "docs_used": 0}
            if self.on_success:
                self.on_success(user, self.lang[0])
            self.destroy()

        ctk.CTkButton(wait, text=("Cerrar" if lbl=="es" else "Close"),
                      fg_color=BLUE, hover_color=BLUE2,
                      command=_close, width=160).pack(pady=(4, 24))

    def _open_payment_demo(self, nm, ls, em, ps):
        """Ventana de pago en MODO DEMO (sin cobro real).
        Se usa cuando Stripe NO está configurado — útil en desarrollo."""
        prices = {"pro": "$4.99", "biz": "$24.99"}
        price_mo = {"pro": "/mes", "biz": "/mes · hasta 10 usuarios"}
        plan_names = {"pro": "Personal Pro", "biz": "Enterprise"}
        plan_color = {"pro": BLUE, "biz": PURPLE}
        lbl = self.lang[0]

        pay = ctk.CTkToplevel(self)
        pay.title("WEP AI — " + ("Modo demo · Sin cobro real" if lbl=="es"
                                  else "Demo mode · No real charge"))
        pay.geometry("480x660")
        pay.resizable(False, False)
        pay.configure(fg_color=BG)
        pay.grab_set(); pay.lift(); pay.focus_force()
        pay.attributes("-topmost", True)

        # ── HEADER ──
        hf = ctk.CTkFrame(pay, fg_color=CARD, corner_radius=0); hf.pack(fill="x")
        ctk.CTkLabel(hf, text="⚠  " + ("MODO DEMO" if lbl=="es" else "DEMO MODE"),
                     font=ctk.CTkFont("Arial", 16, "bold"), text_color=AMBER).pack(pady=(14,2))
        ctk.CTkLabel(hf, text=("Esta versión no procesa cobros reales. "
                               "No ingreses una tarjeta real."
                               if lbl=="es"
                               else "This build does not process real charges. "
                                    "Do not enter a real card."),
                     font=ctk.CTkFont("Arial", 11), text_color=GRAY1,
                     wraplength=420).pack(pady=(0,10))

        # ── RESUMEN DEL PLAN ──
        psum = ctk.CTkFrame(pay, fg_color=CARD2, corner_radius=12,
                            border_width=1, border_color=plan_color[self.sel_plan])
        psum.pack(padx=24, pady=(14,8), fill="x")
        ps_in = ctk.CTkFrame(psum, fg_color="transparent"); ps_in.pack(padx=16, pady=12, fill="x")
        ps_in.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(ps_in, text=plan_names[self.sel_plan],
                     font=ctk.CTkFont("Arial", 15, "bold"),
                     text_color=WHITE, anchor="w").grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(ps_in,
                     text=prices[self.sel_plan] + price_mo[self.sel_plan],
                     font=ctk.CTkFont("Arial", 14, "bold"),
                     text_color=plan_color[self.sel_plan]).grid(row=0, column=1, sticky="e")
        feats = ({"pro": ["✓ Documentos ilimitados","✓ Correcciones IA","✓ Historial 90 días","✓ Soporte por email"],
                  "biz": ["✓ Documentos ilimitados","✓ Hasta 10 usuarios","✓ Historial 1 año","✓ Soporte 24/7"]}
                 if lbl=="es" else
                 {"pro": ["✓ Unlimited documents","✓ AI corrections","✓ 90-day history","✓ Email support"],
                  "biz": ["✓ Unlimited documents","✓ Up to 10 users","✓ 1-year history","✓ 24/7 support"]}
                 )[self.sel_plan]
        ctk.CTkLabel(ps_in, text="  ·  ".join(feats),
                     font=ctk.CTkFont("Arial", 10), text_color=GRAY1,
                     anchor="w", wraplength=400).grid(row=1, column=0, columnspan=2,
                                                       sticky="w", pady=(4,0))

        # ── MÉTODO DE PAGO ──
        card_f = ctk.CTkFrame(pay, fg_color=CARD, corner_radius=12,
                              border_width=1, border_color=BORDER)
        card_f.pack(padx=24, pady=(0,8), fill="x")
        inn = ctk.CTkFrame(card_f, fg_color="transparent"); inn.pack(padx=20, pady=16, fill="x")

        ctk.CTkLabel(inn, text=("💳  Información de pago" if lbl=="es" else "💳  Payment information"),
                     font=ctk.CTkFont("Arial", 13, "bold"), text_color=WHITE,
                     anchor="w").pack(fill="x", pady=(0,12))

        # Iconos de tarjetas
        ic_row = ctk.CTkFrame(inn, fg_color="transparent"); ic_row.pack(fill="x", pady=(0,10))
        for card_name, col in [("VISA", "#1A1F71"), ("Mastercard", "#EB001B"), ("Amex", "#007BC1")]:
            ctk.CTkLabel(ic_row, text=card_name, font=ctk.CTkFont("Arial", 8, "bold"),
                         text_color=WHITE, fg_color=col, corner_radius=4,
                         width=48, height=22).pack(side="left", padx=(0,5))
        ctk.CTkLabel(ic_row, text=("y más..." if lbl=="es" else "and more..."),
                     font=ctk.CTkFont("Arial", 10), text_color=GRAY2).pack(side="left", padx=4)

        # Nombre del titular
        ctk.CTkLabel(inn, text=("Nombre del titular" if lbl=="es" else "Cardholder name"),
                     font=ctk.CTkFont("Arial", 10, "bold"), text_color=GRAY1,
                     anchor="w").pack(fill="x", pady=(0,4))
        f_name = ctk.CTkEntry(inn, height=40, corner_radius=9, font=ctk.CTkFont("Arial", 13),
                              placeholder_text=f"{nm} {ls}".upper(),
                              fg_color=CARD2, border_color=BORDER, border_width=1,
                              text_color=WHITE)
        f_name.pack(fill="x", pady=(0,10))
        f_name.insert(0, f"{nm} {ls}".upper())

        # Número de tarjeta
        ctk.CTkLabel(inn, text=("Número de tarjeta" if lbl=="es" else "Card number"),
                     font=ctk.CTkFont("Arial", 10, "bold"), text_color=GRAY1,
                     anchor="w").pack(fill="x", pady=(0,4))

        card_row = ctk.CTkFrame(inn, fg_color="transparent"); card_row.pack(fill="x", pady=(0,10))
        card_row.grid_columnconfigure(0, weight=1)
        f_card = ctk.CTkEntry(card_row, height=40, corner_radius=9, font=ctk.CTkFont("Arial", 14),
                              placeholder_text="1234  5678  9012  3456",
                              fg_color=CARD2, border_color=BORDER, border_width=1,
                              text_color=WHITE)
        f_card.grid(row=0, column=0, sticky="ew", padx=(0,8))
        ctk.CTkLabel(card_row, text="💳", font=ctk.CTkFont("Arial", 18),
                     fg_color="transparent").grid(row=0, column=1)

        def fmt_card(e):
            # v16.9.1 — bugfix: el cursor saltaba al final en cada tecla porque
            # 'pos' se calculaba pero nunca se restauraba. Ahora reposicionamos el
            # caret contando los dígitos a su izquierda y reinsertando los espacios.
            old_pos = f_card.index(ctk.INSERT)
            old_txt = f_card.get()
            digits_left = len(re.sub(r"\D", "", old_txt[:old_pos]))
            raw = re.sub(r"\D", "", old_txt)[:16]
            grp = "  ".join(raw[i:i+4] for i in range(0, len(raw), 4))
            f_card.delete(0, "end"); f_card.insert(0, grp)
            # mapear 'digits_left' dígitos a su posición real dentro de 'grp'
            new_pos, seen = len(grp), 0
            for idx, ch in enumerate(grp):
                if seen == digits_left:
                    new_pos = idx
                    break
                if ch.isdigit():
                    seen += 1
            f_card.icursor(new_pos)
        f_card.bind("<KeyRelease>", fmt_card)

        # Expiración + CVV
        exp_cvv = ctk.CTkFrame(inn, fg_color="transparent"); exp_cvv.pack(fill="x", pady=(0,10))
        exp_cvv.grid_columnconfigure((0,1), weight=1)

        ctk.CTkLabel(exp_cvv, text=("Fecha de expiración" if lbl=="es" else "Expiry date"),
                     font=ctk.CTkFont("Arial", 10, "bold"), text_color=GRAY1,
                     anchor="w").grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(exp_cvv, text="CVV / CVC",
                     font=ctk.CTkFont("Arial", 10, "bold"), text_color=GRAY1,
                     anchor="w").grid(row=0, column=1, sticky="w", padx=(8,0))

        f_exp = ctk.CTkEntry(exp_cvv, height=40, corner_radius=9, font=ctk.CTkFont("Arial", 13),
                             placeholder_text="MM / AA",
                             fg_color=CARD2, border_color=BORDER, border_width=1,
                             text_color=WHITE)
        f_exp.grid(row=1, column=0, sticky="ew", padx=(0,8), pady=(4,0))

        def fmt_exp(e):
            raw = re.sub(r"\D", "", f_exp.get())[:4]
            if len(raw) > 2:
                val = raw[:2] + " / " + raw[2:]
            else:
                val = raw
            f_exp.delete(0, "end"); f_exp.insert(0, val)
        f_exp.bind("<KeyRelease>", fmt_exp)

        f_cvv = ctk.CTkEntry(exp_cvv, height=40, corner_radius=9, font=ctk.CTkFont("Arial", 13),
                             placeholder_text="123", show="●",
                             fg_color=CARD2, border_color=BORDER, border_width=1,
                             text_color=WHITE)
        f_cvv.grid(row=1, column=1, sticky="ew", padx=(8,0), pady=(4,0))

        # Error label
        err_pay = ctk.CTkLabel(inn, text="", font=ctk.CTkFont("Arial", 11),
                               text_color=RED, wraplength=400)
        err_pay.pack(fill="x", pady=(6,0))

        # Texto legal
        legal_f = ctk.CTkFrame(pay, fg_color="transparent"); legal_f.pack(padx=24, fill="x")
        ctk.CTkLabel(legal_f,
                     text=("🔒 Tu pago es seguro. Se te cobrará mensualmente. Cancela cuando quieras."
                            if lbl=="es"
                            else "🔒 Your payment is secure. You'll be billed monthly. Cancel anytime."),
                     font=ctk.CTkFont("Arial", 10), text_color=GRAY2,
                     wraplength=432, justify="center").pack(pady=(0,4))

        # ── BOTONES ──
        btn_f = ctk.CTkFrame(pay, fg_color="transparent"); btn_f.pack(padx=24, pady=8, fill="x")
        btn_f.grid_columnconfigure((0,1), weight=1)

        ctk.CTkButton(btn_f, text=("Cancelar" if lbl=="es" else "Cancel"),
                      height=44, corner_radius=10,
                      font=ctk.CTkFont("Arial", 13, "bold"),
                      fg_color="transparent", border_width=1, border_color=BORDER,
                      text_color=GRAY1, hover_color=CARD2,
                      command=pay.destroy).grid(row=0, column=0, padx=(0,8), sticky="ew")

        pay_btn_txt = (f"Pagar {prices[self.sel_plan]}/mes" if lbl=="es"
                       else f"Pay {prices[self.sel_plan]}/month")
        pay_btn = ctk.CTkButton(btn_f, text=pay_btn_txt,
                                height=44, corner_radius=10,
                                font=ctk.CTkFont("Arial", 13, "bold"),
                                fg_color=plan_color[self.sel_plan],
                                hover_color=PURPLE2 if self.sel_plan=="biz" else BLUE2)
        pay_btn.grid(row=0, column=1, sticky="ew")

        def process_payment():
            card_raw = re.sub(r"\D", "", f_card.get())
            exp_raw  = re.sub(r"\D", "", f_exp.get())
            cvv_raw  = f_cvv.get().strip()
            name_raw = f_name.get().strip()

            # Validaciones
            if not name_raw:
                err_pay.configure(text="⚠ " + ("Ingresa el nombre del titular" if lbl=="es" else "Enter cardholder name")); return
            if len(card_raw) < 16:
                err_pay.configure(text="⚠ " + ("Número de tarjeta incompleto (16 dígitos)" if lbl=="es" else "Incomplete card number (16 digits)")); return
            if len(exp_raw) < 4:
                err_pay.configure(text="⚠ " + ("Ingresa la fecha de expiración" if lbl=="es" else "Enter expiry date")); return
            month = int(exp_raw[:2])
            if month < 1 or month > 12:
                err_pay.configure(text="⚠ " + ("Mes inválido (01-12)" if lbl=="es" else "Invalid month (01-12)")); return
            if len(cvv_raw) < 3:
                err_pay.configure(text="⚠ " + ("CVV inválido (mínimo 3 dígitos)" if lbl=="es" else "Invalid CVV (minimum 3 digits)")); return

            # Simular procesamiento
            err_pay.configure(text="")
            pay_btn.configure(state="disabled",
                              text=("Procesando..." if lbl=="es" else "Processing..."))
            pay.update()

            # Registrar usuario con plan pagado
            uid, err = db.register_user(nm, ls, em, ps, self.sel_plan)
            if err == "email_exists":
                pay_btn.configure(state="normal", text=pay_btn_txt)
                err_pay.configure(text="⚠ " + self._t("err_exists")); return
            if err:
                pay_btn.configure(state="normal", text=pay_btn_txt)
                err_pay.configure(text=f"Error: {err}"); return

            # Persistir sólo los últimos 4 dígitos (la columna ya existe vía
            # migración en init_db — no más ALTER TABLE en cada pago).
            try:
                db.save_card_last4(uid, card_raw[-4:])
            except Exception as e:
                print(f"[WEP AI] save_card_last4 failed: {e}")

            pay.destroy()
            # Confirmación honesta — modo demo
            messagebox.showinfo(
                "WEP AI — " + ("Cuenta demo activada" if lbl=="es" else "Demo account activated"),
                (f"✓ Plan {plan_names[self.sel_plan]} activado en modo demo.\n"
                 f"Sin cobro real. Tarjeta no procesada.\n\n"
                 "Para activar cobros reales, integra Stripe Checkout."
                 if lbl=="es" else
                 f"✓ {plan_names[self.sel_plan]} plan activated in demo mode.\n"
                 f"No real charge. Card not processed.\n\n"
                 "To enable real payments, integrate Stripe Checkout."))

            user = {"id": uid, "name": nm, "last_name": ls, "email": em,
                    "plan": self.sel_plan, "docs_used": 0}
            if self.on_success:
                self.on_success(user, self.lang[0])
            self.destroy()

        pay_btn.configure(command=process_payment)

    def _do_register(self, nm, ls, em, ps):
        """Registrar cuenta gratuita sin pago."""
        uid, err = db.register_user(nm, ls, em, ps, "free")
        if err == "email_exists":
            self.err_lbl.configure(text=self._t("err_exists")); return
        if err:
            self.err_lbl.configure(text=f"Error: {err}"); return

        # v14.8 — Email de bienvenida en background daemon (no bloquea UI).
        # Si no hay EMAIL_PROVIDER configurado, send_welcome es no-op (dry-run).
        _send_welcome_async(em, nm, lang=self.lang[0])

        messagebox.showinfo("WEP AI",
            f"{self._t('ok_title')}\n{self._t('ok_sub')}")
        user = {"id": uid, "name": nm, "last_name": ls, "email": em,
                "plan": "free", "docs_used": 0}
        if self.on_success:
            self.on_success(user, self.lang[0])
        self.destroy()

    def _go_back(self):
        if self.on_cancel:
            self.on_cancel(self.lang[0])
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
#  VENTANA 3 — APP PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
class WEPAIApp(ctk.CTk):
    def __init__(self, user: dict, lang: str = "es"):
        super().__init__()
        self.user  = user
        self.lang  = [lang]
        self.title("WEP AI")
        self.geometry("1160x760")
        self.minsize(920, 640)
        self.configure(fg_color=BG)
        db.init_db()
        self.conversation_id   = None
        self.chat_history      = []
        self.current_doc_path  = None
        self.current_gen_data  = None
        self.selected_doc_type = "auto"  # v15.16.3 — auto = la IA decide por el texto
        self.current_doc_type  = None
        # v16.9.1 — bugfix: 'state' colisionaba con Tk.state(); renombrado a app_state.
        self.app_state         = "idle"
        self.photo_ref         = None
        # v16.9.1 — bugfix: _prefs_path / _prefs se leían en __init__ pero nunca se
        # asignaban (crash AttributeError al arrancar). Inicializados aquí.
        self._prefs_path = os.path.expanduser("~/.wepai/prefs.json")
        self._prefs      = self._load_prefs()
        self._build()
        self._bind_shortcuts()
        self._load_conversations()
        # v14.7 — Banner de estado de suscripción (past_due / cancelled)
        try:
            self._refresh_billing_banner()
        except Exception as _e:
            print(f"[WEP AI] _refresh_billing_banner failed: {_e}")
        # Onboarding primera vez
        if self._prefs.get("first_launch", True):
            self.after(600, self._show_onboarding)
        else:
            self._new_chat()

    def _t(self, k): return T(self.lang[0], k)

    # ── BUILD ─────────────────────────────────────────────────────────────────

    def _load_prefs(self):
        try:
            os.makedirs(os.path.dirname(self._prefs_path), exist_ok=True)
            if os.path.exists(self._prefs_path):
                return json.load(open(self._prefs_path))
        except Exception as e:
            print(f"[WEP AI] _load_prefs failed: {e}")
        return {"first_launch": True, "api_key": "", "docs_folder": "~/Documents/WEP_AI"}

    def _save_prefs(self):
        try:
            json.dump(self._prefs, open(self._prefs_path, "w"), indent=2)
        except Exception as e:
            print(f"[WEP AI] _save_prefs failed: {e}")

    def _bind_shortcuts(self):
        self.bind("<Command-n>", lambda e: self._new_chat())
        self.bind("<Command-N>", lambda e: self._new_chat())
        self.bind("<Command-comma>", lambda e: self._open_settings())

    def _select_doc_type(self, dtype):
        self.selected_doc_type = dtype
        if hasattr(self, 'doc_btns'):
            colors = {"auto":BLUE,"word":WORD_C,"excel":EXCEL_C,"powerpoint":PPT_C}
            for d, btn in self.doc_btns.items():
                sel = (d == dtype)
                btn.configure(
                    fg_color=colors[d] if sel else "transparent",
                    text_color=WHITE if sel else GRAY1,
                    border_color=colors[d] if sel else BORDER)

    def _build(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── SIDEBAR ──────────────────────────────────────────────────────────
        sb = ctk.CTkFrame(self, width=210, corner_radius=0, fg_color=CARD)
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_rowconfigure(2, weight=1)
        sb.grid_propagate(False)

        self.new_btn = ctk.CTkButton(sb, text=self._t("new_conv"),
                                     command=self._new_chat, height=40, corner_radius=10,
                                     font=ctk.CTkFont("Arial", 13, "bold"),
                                     fg_color=BLUE, hover_color=BLUE2)
        self.new_btn.grid(row=0, column=0, padx=10, pady=(12, 5), sticky="ew")

        self.convs_lbl = ctk.CTkLabel(sb, text=self._t("convs"),
                                      font=ctk.CTkFont("Arial", 9, "bold"), text_color=GRAY2)
        self.convs_lbl.grid(row=1, column=0, padx=14, pady=(3, 3), sticky="w")

        self.conv_scroll = ctk.CTkScrollableFrame(sb, corner_radius=0, fg_color="transparent")
        self.conv_scroll.grid(row=2, column=0, sticky="nsew", padx=4, pady=2)
        self.conv_buttons = []

        # Plan label
        plan = self.user.get("plan", "free")
        pc   = PLAN_COLORS.get(plan, GRAY1)
        pr   = PLAN_PRICES if self.lang[0] == "es" else PLAN_PRICES_EN
        pn   = PLAN_NAMES_ES
        self.plan_lbl = ctk.CTkLabel(sb,
                                     text=f"{pn.get(plan,'')} · {pr.get(plan,'')}",
                                     font=ctk.CTkFont("Arial", 9, "bold"), text_color=pc)
        self.plan_lbl.grid(row=3, column=0, padx=12, pady=(4, 0), sticky="w")

        # Docs used (free only)
        if plan == "free":
            used = db.get_docs_used(self.user["id"])
            self.docs_lbl = ctk.CTkLabel(sb, text=f"Docs: {used}/5",
                                         font=ctk.CTkFont("Arial", 9), text_color=AMBER)
            self.docs_lbl.grid(row=4, column=0, padx=12, pady=(0, 4), sticky="w")
        else:
            # v14.6 — Botón "Manage billing" para usuarios pagos
            billing_lbl = ("Gestionar facturación" if self.lang[0] == "es"
                           else "Manage billing")
            self.billing_btn = ctk.CTkButton(
                sb, text=billing_lbl,
                font=ctk.CTkFont("Arial", 10),
                fg_color="transparent", text_color=BLUE,
                hover_color=CARD2, border_width=1, border_color=BORDER,
                height=24, command=self._open_billing_portal)
            self.billing_btn.grid(row=4, column=0, padx=12, pady=(2, 4), sticky="ew")

        foot = ctk.CTkFrame(sb, fg_color=CARD, corner_radius=0)
        foot.grid(row=5, column=0, sticky="ew")
        self.status_lbl = ctk.CTkLabel(foot, text="● " + self._t("ready"),
                                       font=ctk.CTkFont("Arial", 10),
                                       text_color=BLUE, anchor="w")
        self.status_lbl.pack(padx=12, pady=(6, 3), fill="x")

        mr = ctk.CTkFrame(foot, fg_color=CARD2, corner_radius=8,
                          border_width=1, border_color=BORDER)
        mr.pack(padx=10, pady=(0, 10), fill="x")
        # v15.16.3 — En vez de mostrar el modelo (dato técnico irrelevante para el
        # usuario final), mostramos su uso de documentos del mes, que sí le importa.
        mi = ctk.CTkFrame(mr, fg_color="transparent"); mi.pack(side="left", fill="x", expand=True, padx=10, pady=8)
        _docs_mes = 0
        try:
            _docs_mes = db.get_docs_used(self.user["id"])
        except Exception:
            pass
        _uso_lbl = ("Documentos este mes" if self.lang[0] == "es" else "Documents this month")
        ctk.CTkLabel(mi, text=_uso_lbl.upper(),
                     font=ctk.CTkFont("Arial", 8, "bold"), text_color=GRAY2).pack(anchor="w")
        ctk.CTkLabel(mi, text=f"{_docs_mes} " + ("generados" if self.lang[0]=="es" else "generated"),
                     font=ctk.CTkFont("Arial", 11, "bold"), text_color=WHITE).pack(anchor="w")
        _plan_txt = ("Plan " + PLAN_NAMES_ES.get(plan, "") if self.lang[0]=="es"
                     else PLAN_NAMES_EN.get(plan, "") + " plan")
        ctk.CTkLabel(mi, text=_plan_txt,
                     font=ctk.CTkFont("Arial", 9), text_color=(GREEN if plan != "free" else AMBER)).pack(anchor="w")

        # ── MAIN ─────────────────────────────────────────────────────────────
        self.main = ctk.CTkFrame(self, corner_radius=0, fg_color=CARD2)
        self.main.grid(row=0, column=1, sticky="nsew")
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(2, weight=1)

        # v14.7 — Status banner (oculto por default). Aparece si la suscripción
        # está past_due o cancelled. Lo populamos al final con _refresh_billing_banner.
        self.banner_frame = ctk.CTkFrame(self.main, fg_color=AMBER, corner_radius=0,
                                          height=0)
        self.banner_frame.grid(row=1, column=0, sticky="ew")
        self.banner_frame.grid_remove()  # oculto inicialmente
        self.banner_lbl = ctk.CTkLabel(self.banner_frame, text="",
                                        font=ctk.CTkFont("Arial", 11, "bold"),
                                        text_color="#1c1c1e")
        self.banner_lbl.pack(side="left", padx=14, pady=8)
        self.banner_btn = ctk.CTkButton(self.banner_frame,
                                         text="", width=140, height=24,
                                         font=ctk.CTkFont("Arial", 10, "bold"),
                                         fg_color="#1c1c1e",
                                         text_color=WHITE,
                                         hover_color="#000000",
                                         command=self._open_billing_portal)
        self.banner_btn.pack(side="right", padx=14, pady=6)

        # Header con logo centrado
        hdr = ctk.CTkFrame(self.main, height=60, corner_radius=0, fg_color=BG)
        hdr.grid(row=0, column=0, sticky="ew"); hdr.grid_propagate(False)

        lf = ctk.CTkFrame(hdr, fg_color="transparent")
        lf.place(relx=0.5, rely=0.35, anchor="center")
        lr = ctk.CTkFrame(lf, fg_color="transparent"); lr.pack()
        for ch, col in [("W", WORD_C), ("E", EXCEL_C), ("P", PPT_C)]:
            ctk.CTkLabel(lr, text=ch, font=ctk.CTkFont("Arial", 26, "bold"),
                         text_color=col).pack(side="left")
        ctk.CTkLabel(lr, text=" AI", font=ctk.CTkFont("Arial", 17, "normal"),
                     text_color=WHITE).pack(side="left")
        sr = ctk.CTkFrame(lf, fg_color="transparent"); sr.pack()
        for nm, col in [("Word", WORD_C), ("  ·  ", GRAY3), ("Excel", EXCEL_C),
                         ("  ·  ", GRAY3), ("PowerPoint", PPT_C)]:
            ctk.CTkLabel(sr, text=nm,
                         font=ctk.CTkFont("Arial", 10, "bold" if nm.strip() != "·" else "normal"),
                         text_color=col).pack(side="left")

        # Controls a la derecha
        rc = ctk.CTkFrame(hdr, fg_color="transparent")
        rc.place(relx=1.0, rely=0.5, anchor="e", x=-14)
        lf2 = ctk.CTkFrame(rc, fg_color=CARD, corner_radius=16,
                            border_width=1, border_color=BORDER)
        lf2.pack(side="left", padx=(0, 8))
        self.bes = ctk.CTkButton(lf2, text="ES", width=36, height=26, corner_radius=14,
                                 font=ctk.CTkFont("Arial", 10, "bold"),
                                 fg_color=BLUE if self.lang[0]=="es" else "transparent",
                                 text_color=WHITE if self.lang[0]=="es" else GRAY2,
                                 hover_color=CARD2, command=lambda: self._set_lang("es"))
        self.bes.pack(side="left", padx=2, pady=2)
        self.ben = ctk.CTkButton(lf2, text="EN", width=36, height=26, corner_radius=14,
                                 font=ctk.CTkFont("Arial", 10, "bold"),
                                 fg_color=BLUE if self.lang[0]=="en" else "transparent",
                                 text_color=WHITE if self.lang[0]=="en" else GRAY2,
                                 hover_color=CARD2, command=lambda: self._set_lang("en"))
        self.ben.pack(side="left", padx=2, pady=2)

        name = self.user.get("name", "Usuario")
        self.user_btn = ctk.CTkButton(rc, text=f"👤  {name}  ▾", height=32, corner_radius=16,
                                      font=ctk.CTkFont("Arial", 11, "bold"),
                                      fg_color=CARD, hover_color=CARD2,
                                      border_width=1, border_color=BORDER,
                                      text_color=GRAY1, command=self._user_menu)
        self.user_btn.pack(side="left")

        # Topbar conversación
        tb = ctk.CTkFrame(self.main, height=36, corner_radius=0, fg_color=CARD)
        tb.grid(row=1, column=0, sticky="ew")
        tb.grid_columnconfigure(0, weight=1); tb.grid_propagate(False)
        self.chat_title = ctk.CTkLabel(tb, text="", font=ctk.CTkFont("Arial", 12, "bold"),
                                       text_color=GRAY1, anchor="w")
        self.chat_title.grid(row=0, column=0, padx=14, pady=6, sticky="w")
        self.doc_badge = ctk.CTkLabel(tb, text="", font=ctk.CTkFont("Arial", 10, "bold"),
                                      text_color=WHITE, corner_radius=10)
        self.doc_badge.grid(row=0, column=1, padx=10, pady=6, sticky="e")

        # Content
        cont = ctk.CTkFrame(self.main, fg_color="transparent")
        cont.grid(row=2, column=0, sticky="nsew")
        cont.grid_columnconfigure(0, weight=3); cont.grid_columnconfigure(1, weight=2)
        cont.grid_rowconfigure(0, weight=1)

        self.chat_scroll = ctk.CTkScrollableFrame(cont, fg_color=CARD2, corner_radius=0)
        self.chat_scroll.grid(row=0, column=0, sticky="nsew")
        self.chat_scroll.grid_columnconfigure(0, weight=1)

        self.prev_panel = ctk.CTkFrame(cont, fg_color=CARD, corner_radius=0)
        self.prev_panel.grid(row=0, column=1, sticky="nsew", padx=(1, 0))
        self.prev_panel.grid_columnconfigure(0, weight=1)
        self.prev_panel.grid_rowconfigure(1, weight=1)
        self.prev_title = ctk.CTkLabel(self.prev_panel, text=self._t("preview"),
                                       font=ctk.CTkFont("Arial", 11, "bold"), text_color=GRAY1)
        self.prev_title.grid(row=0, column=0, padx=12, pady=(10, 4), sticky="w")
        self.prev_lbl = ctk.CTkLabel(self.prev_panel, text=self._t("preview_sub"),
                                     font=ctk.CTkFont("Arial", 11), text_color=GRAY1,
                                     anchor="center")
        self.prev_lbl.grid(row=1, column=0, padx=8, pady=8, sticky="nsew")
        self.open_btn = ctk.CTkButton(self.prev_panel, text=self._t("open_in"),
                                      command=self._open_in_office,
                                      state="disabled", height=30,
                                      font=ctk.CTkFont("Arial", 11, "bold"),
                                      fg_color=BLUE, hover_color=BLUE2, corner_radius=8)
        self.open_btn.grid(row=2, column=0, padx=12, pady=(0, 10), sticky="ew")

        # Input
        inp_f = ctk.CTkFrame(self.main, height=132, corner_radius=0, fg_color=CARD)
        inp_f.grid(row=3, column=0, sticky="ew")
        inp_f.grid_columnconfigure(0, weight=1); inp_f.grid_propagate(False)

        # Selector tipo documento
        doc_sel = ctk.CTkFrame(inp_f, fg_color="transparent")
        doc_sel.grid(row=0, column=0, padx=12, pady=(8,3), sticky="ew")
        self._type_lbl = ctk.CTkLabel(doc_sel,
                                       text=("Formato (opcional):" if self.lang[0]=="es"
                                             else "Format (optional):"),
                                       font=ctk.CTkFont("Arial",11,"bold"), text_color=GRAY1)
        self._type_lbl.pack(side="left", padx=(0,8))
        self.doc_btns = {}
        for dtype, label, col in [("auto","✨ Auto",BLUE),
                                   ("word","📄 Word",WORD_C),
                                   ("excel","📊 Excel",EXCEL_C),
                                   ("powerpoint","📑 PowerPoint",PPT_C)]:
            btn = ctk.CTkButton(doc_sel, text=label, height=28, corner_radius=14,
                               font=ctk.CTkFont("Arial",11,"bold"),
                               fg_color=col if dtype==self.selected_doc_type else "transparent",
                               text_color=WHITE if dtype==self.selected_doc_type else GRAY1,
                               border_width=1,
                               border_color=col if dtype==self.selected_doc_type else BORDER,
                               hover_color=CARD2,
                               command=lambda d=dtype: self._select_doc_type(d))
            btn.pack(side="left", padx=4)
            self.doc_btns[dtype] = btn

        # Chips rápidas
        cf = ctk.CTkFrame(inp_f, fg_color="transparent")
        cf.grid(row=1, column=0, padx=12, pady=(0,3), sticky="ew")
        for ch, col in [("Contrato",GRAY1),("Inventario",GRAY1),("Nómina",GRAY1),
                         ("Presupuesto",GRAY1),("Pitch deck",GRAY1),("Informe",GRAY1)]:
            ctk.CTkButton(cf, text=ch, height=22, corner_radius=11,
                          font=ctk.CTkFont("Arial",10), fg_color="transparent",
                          border_width=1, border_color=BORDER, text_color=col,
                          hover_color=CARD2,
                          command=lambda t=ch: self._chip(t)).pack(side="left", padx=2)

        ir = ctk.CTkFrame(inp_f, fg_color="transparent")
        ir.grid(row=2, column=0, padx=12, pady=(0, 8), sticky="ew")
        ir.grid_columnconfigure(0, weight=1)

        self.inp = ctk.CTkTextbox(ir, height=54, corner_radius=12,
                                  font=ctk.CTkFont("Arial", 13),
                                  border_width=1, border_color=BORDER,
                                  fg_color=CARD2, text_color=GRAY1)
        self.inp.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.inp.insert("1.0", self._t("inp_ph"))
        self.inp.bind("<FocusIn>",      lambda e: self._clr_ph())
        self.inp.bind("<Control-Return>", lambda e: self._send())
        self.inp.bind("<Command-Return>", lambda e: self._send())

        self.send_btn = ctk.CTkButton(ir, text="↑", width=54, height=54,
                                      command=self._send, corner_radius=12,
                                      font=ctk.CTkFont("Arial", 22, "bold"),
                                      fg_color=BLUE, hover_color=BLUE2)
        self.send_btn.grid(row=0, column=1)
        self.hint_lbl = ctk.CTkLabel(ir, text=self._t("cmd_hint"),
                                     font=ctk.CTkFont("Arial", 9), text_color=GRAY2)
        self.hint_lbl.grid(row=1, column=0, sticky="w", padx=2)

    # ── HELPERS ───────────────────────────────────────────────────────────────
    def _set_lang(self, l):
        self.lang[0] = l
        self.bes.configure(fg_color=BLUE if l=="es" else "transparent",
                           text_color=WHITE if l=="es" else GRAY2)
        self.ben.configure(fg_color=BLUE if l=="en" else "transparent",
                           text_color=WHITE if l=="en" else GRAY2)
        self.new_btn.configure(text=self._t("new_conv"))
        self.convs_lbl.configure(text=self._t("convs"))
        if hasattr(self,"_type_lbl"):
            self._type_lbl.configure(text="Formato (opcional):" if l=="es" else "Format (optional):")
        self.prev_title.configure(text=self._t("preview"))
        self.open_btn.configure(text=self._t("open_in"))
        self.hint_lbl.configure(text=self._t("cmd_hint"))
        self._set_state(self.app_state)

    def _clr_ph(self):
        if self.inp.get("1.0", "end-1c") == self._t("inp_ph"):
            self.inp.delete("1.0", "end")
            self.inp.configure(text_color=WHITE)

    def _chip(self, t):
        self.inp.delete("1.0", "end")
        self.inp.insert("1.0", t)
        self.inp.configure(text_color=WHITE)
        self.inp.focus()

    def _user_menu(self):
        plan = self.user.get("plan", "free")
        pc   = PLAN_COLORS.get(plan, GRAY1)
        pr   = PLAN_PRICES if self.lang[0] == "es" else PLAN_PRICES_EN
        m = ctk.CTkToplevel(self)
        m.overrideredirect(True)
        m.configure(fg_color=CARD2)
        m.attributes("-topmost", True)
        x = self.user_btn.winfo_rootx()
        y = self.user_btn.winfo_rooty() + self.user_btn.winfo_height() + 4
        m.geometry(f"220x240+{x-100}+{y}")
        ctk.CTkFrame(m, height=1, fg_color=BORDER).pack(fill="x")
        h = ctk.CTkFrame(m, fg_color=CARD, corner_radius=0); h.pack(fill="x")
        nm = f"{self.user.get('name','')} {self.user.get('last_name','')}".strip()
        ctk.CTkLabel(h, text=nm, font=ctk.CTkFont("Arial", 13, "bold"),
                     text_color=WHITE, anchor="w").pack(padx=14, pady=(10, 1), fill="x")
        ctk.CTkLabel(h, text=self.user.get("email", ""),
                     font=ctk.CTkFont("Arial", 10), text_color=GRAY1,
                     anchor="w").pack(padx=14, pady=(0, 4), fill="x")
        ctk.CTkLabel(h, text=f"{PLAN_NAMES_ES.get(plan,'')} · {pr.get(plan,'')}",
                     font=ctk.CTkFont("Arial", 10, "bold"), text_color=pc,
                     anchor="w").pack(padx=14, pady=(0, 8), fill="x")
        ctk.CTkFrame(m, height=1, fg_color=BORDER).pack(fill="x")

        def item(icon, text_key, cmd, danger=False):
            col = RED if danger else GRAY1
            ctk.CTkButton(m, text=f"{icon}  {self._t(text_key)}", anchor="w", height=36,
                          corner_radius=0, font=ctk.CTkFont("Arial", 11),
                          fg_color="transparent", hover_color=CARD,
                          text_color=col,
                          command=lambda: (m.destroy(), cmd())).pack(fill="x", padx=2, pady=1)

        item("👤", "profile",  lambda: self._open_profile())
        item("⚙",  "settings", lambda: self._open_settings())
        item("📁", "my_docs",  lambda: self._open_docs_window())
        item("❓", "help_menu", lambda: self._open_help())
        item("ℹ️", "about_menu",lambda: self._open_about())
        # v14.8 — Cancelar suscripción directo (solo para planes pagos).
        # El método ya existía en v14.7 pero ningún botón lo invocaba.
        if plan in ("pro", "biz"):
            ctk.CTkFrame(m, height=1, fg_color=BORDER).pack(fill="x", pady=3)
            item("✕", "cancel_sub", self._cancel_subscription_directly, danger=True)
        ctk.CTkFrame(m, height=1, fg_color=BORDER).pack(fill="x", pady=3)
        item("→",  "logout",   self._logout, danger=True)
        m.bind("<FocusOut>", lambda e: m.destroy())
        m.focus_set()

    # ── PROFILE ───────────────────────────────────────────────────────────────
    def _open_profile(self):
        w = ctk.CTkToplevel(self)
        w.title("WEP AI — " + self._t("profile"))
        w.geometry("440x520")
        w.resizable(False, False)
        w.configure(fg_color=BG)
        w.grab_set(); w.lift(); w.focus_force()
        w.attributes("-topmost", True)

        # Header
        hf = ctk.CTkFrame(w, fg_color=CARD, corner_radius=0)
        hf.pack(fill="x")
        ctk.CTkLabel(hf, text=self._t("profile"),
                     font=ctk.CTkFont("Arial", 16, "bold"),
                     text_color=WHITE).pack(pady=(16, 2))
        ctk.CTkLabel(hf, text=self.user.get("email", ""),
                     font=ctk.CTkFont("Arial", 11), text_color=GRAY1).pack(pady=(0, 12))

        # Avatar circle
        av_txt = (self.user.get("name","?")[:1] + self.user.get("last_name","")[:1]).upper() or "?"
        ctk.CTkLabel(hf, text=av_txt, width=60, height=60,
                     font=ctk.CTkFont("Arial", 22, "bold"),
                     fg_color=BLUE, text_color=WHITE,
                     corner_radius=30).pack(pady=(0, 16))

        # Form
        card = ctk.CTkFrame(w, fg_color=CARD, corner_radius=14,
                             border_width=1, border_color=BORDER)
        card.pack(padx=24, pady=16, fill="x")
        inn = ctk.CTkFrame(card, fg_color="transparent")
        inn.pack(padx=20, pady=16, fill="x")

        lbl_es = {"name_lbl":"Nombre","last_lbl":"Apellido",
                  "email_lbl":"Correo electrónico","pass_lbl":"Nueva contraseña (opcional)"}
        lbl_en = {"name_lbl":"First name","last_lbl":"Last name",
                  "email_lbl":"Email address","pass_lbl":"New password (optional)"}
        lbls = lbl_es if self.lang[0]=="es" else lbl_en

        fields = {}
        row2 = ctk.CTkFrame(inn, fg_color="transparent"); row2.pack(fill="x", pady=(0,10))
        row2.grid_columnconfigure((0,1), weight=1)
        for i, (key, ph) in enumerate([("name_lbl", self.user.get("name","")),
                                        ("last_lbl",  self.user.get("last_name",""))]):
            ctk.CTkLabel(row2, text=lbls[key],
                         font=ctk.CTkFont("Arial", 10, "bold"),
                         text_color=GRAY1, anchor="w").grid(row=0, column=i, sticky="w",
                                                             padx=(0,8) if i==0 else 0)
            e = ctk.CTkEntry(row2, height=38, corner_radius=9, font=ctk.CTkFont("Arial", 12),
                             fg_color=CARD2, border_color=BORDER, border_width=1,
                             text_color=WHITE)
            e.insert(0, ph)
            e.grid(row=1, column=i, sticky="ew", padx=(0,8) if i==0 else 0, pady=(4,0))
            fields[key] = e

        for key, ph in [("email_lbl", self.user.get("email","")), ("pass_lbl", "")]:
            ctk.CTkLabel(inn, text=lbls[key],
                         font=ctk.CTkFont("Arial", 10, "bold"),
                         text_color=GRAY1, anchor="w").pack(fill="x", pady=(8,4))
            e = ctk.CTkEntry(inn, height=38, corner_radius=9, font=ctk.CTkFont("Arial", 12),
                             fg_color=CARD2, border_color=BORDER, border_width=1,
                             text_color=WHITE,
                             show="●" if key=="pass_lbl" else None)
            e.insert(0, ph)
            e.pack(fill="x")
            fields[key] = e

        msg_lbl = ctk.CTkLabel(inn, text="", font=ctk.CTkFont("Arial", 11), text_color=GREEN)
        msg_lbl.pack(fill="x", pady=(8,0))

        def save():
            import hashlib
            nm = fields["name_lbl"].get().strip()
            ls = fields["last_lbl"].get().strip()
            em = fields["email_lbl"].get().strip()
            ps = fields["pass_lbl"].get().strip()
            if not nm or not ls or not em:
                msg_lbl.configure(text="⚠ " + ("Completa todos los campos" if self.lang[0]=="es" else "Fill in all fields"),
                                  text_color=RED); return
            con = db._con(); cur = con.cursor()
            if ps:
                ph = hashlib.sha256(ps.encode()).hexdigest()
                cur.execute("UPDATE users SET name=?,last_name=?,email=?,password_hash=? WHERE id=?",
                            (nm, ls, em, ph, self.user["id"]))
            else:
                cur.execute("UPDATE users SET name=?,last_name=?,email=? WHERE id=?",
                            (nm, ls, em, self.user["id"]))
            con.commit(); con.close()
            self.user["name"] = nm; self.user["last_name"] = ls; self.user["email"] = em
            # Update UI
            self.user_btn.configure(text=f"👤  {nm}  ▾")
            msg_lbl.configure(text="✓ " + ("Guardado correctamente" if self.lang[0]=="es"
                                            else "Saved successfully"), text_color=GREEN)

        ctk.CTkButton(inn, text="Guardar cambios" if self.lang[0]=="es" else "Save changes",
                      height=42, corner_radius=10,
                      font=ctk.CTkFont("Arial", 13, "bold"),
                      fg_color=BLUE, hover_color=BLUE2, command=save).pack(fill="x", pady=(12,0))

    # ── BILLING (v14.6) ────────────────────────────────────────────────────────
    def _refresh_billing_banner(self):
        """v14.7 — Muestra/oculta el banner según el estado de la suscripción.
        Llamado: al iniciar el app principal y cuando se sospecha cambio de
        estado (post visita al portal, polling)."""
        try:
            status = db.get_user_field(self.user["id"], "subscription_status")
        except Exception:
            status = None
        lbl = self.lang[0]

        if status == "past_due":
            # Pago vencido — banner ámbar urgente
            if lbl == "es":
                msg = "⚠ Tu último pago falló. Actualizá tu tarjeta o tu acceso se suspenderá en 7 días."
                btn = "Actualizar tarjeta"
            else:
                msg = "⚠ Your last payment failed. Update your card or your access will be suspended in 7 days."
                btn = "Update card"
            self.banner_frame.configure(fg_color=AMBER)
            self.banner_lbl.configure(text=msg)
            self.banner_btn.configure(text=btn)
            self.banner_frame.grid()
        elif status == "cancelled":
            # Cancelada — banner gris informativo (todavía con acceso hasta fin de período)
            if lbl == "es":
                msg = "Tu suscripción fue cancelada. Mantenés acceso hasta el fin del período actual."
                btn = "Reactivar"
            else:
                msg = "Your subscription has been cancelled. You retain access until the end of the current period."
                btn = "Reactivate"
            self.banner_frame.configure(fg_color="#3a3a3c")
            self.banner_lbl.configure(text=msg, text_color=WHITE)
            self.banner_btn.configure(text=btn, fg_color=BLUE, hover_color=BLUE2)
            self.banner_frame.grid()
        else:
            # Active o sin estado especial — sin banner
            self.banner_frame.grid_remove()

    def _cancel_subscription_directly(self):
        """v14.7 — Cancela directo sin abrir el portal completo. Confirmación
        obligatoria. La cancelación toma efecto al final del período actual
        (el usuario mantiene acceso hasta entonces). El webhook event
        customer.subscription.deleted disparará el email + downgrade."""
        from tkinter import messagebox
        lbl = self.lang[0]
        sid = db.get_user_field(self.user["id"], "stripe_subscription_id")
        if not sid:
            messagebox.showinfo("WEP AI",
                ("No encontramos suscripción activa." if lbl == "es"
                 else "No active subscription found."))
            return

        # Confirmación
        confirm = messagebox.askyesno(
            "WEP AI — " + ("Confirmar cancelación" if lbl == "es"
                            else "Confirm cancellation"),
            ("¿Estás seguro de cancelar tu suscripción?\n\n"
             "Mantenés acceso hasta el final del período actual. "
             "Después tu cuenta vuelve al plan Free." if lbl == "es" else
             "Are you sure you want to cancel your subscription?\n\n"
             "You'll keep access until the end of the current period. "
             "After that your account reverts to the Free plan."))
        if not confirm:
            return

        try:
            from storage import stripe_client
        except ImportError:
            stripe_client = None
        if not stripe_client or not stripe_client.is_configured():
            messagebox.showerror("WEP AI",
                ("Stripe no configurado. Cancelá desde el portal o "
                 "contactá support@wepai.app." if lbl == "es" else
                 "Stripe not configured. Cancel from the portal or "
                 "contact support@wepai.app."))
            return

        if stripe_client.cancel_subscription(sid):
            messagebox.showinfo("WEP AI",
                ("Cancelación programada. Vas a recibir un email de "
                 "confirmación y mantenés acceso hasta el fin del período."
                 if lbl == "es" else
                 "Cancellation scheduled. You'll receive a confirmation email "
                 "and keep access until the end of the period."))
            # Refrescar banner
            try:
                self._refresh_billing_banner()
            except Exception:
                pass
        else:
            messagebox.showerror("WEP AI",
                ("No pudimos procesar la cancelación. Reintentá desde el portal."
                 if lbl == "es" else
                 "Could not process cancellation. Please try from the portal."))

    def _open_billing_portal(self):
        """Abre el Stripe Customer Portal en el navegador para que el usuario
        gestione su tarjeta, cambie de plan, vea facturas o cancele."""
        import webbrowser
        from tkinter import messagebox
        lbl = self.lang[0]

        # Necesitamos el stripe_customer_id del usuario (lo populó el webhook)
        customer_id = db.get_user_field(self.user["id"], "stripe_customer_id")
        if not customer_id:
            messagebox.showinfo(
                "WEP AI",
                ("No encontramos información de facturación. Si pagaste "
                 "recientemente, esperá unos minutos a que se sincronice y "
                 "volvé a intentarlo. Si el problema persiste, escribí a "
                 "support@wepai.app."
                 if lbl == "es" else
                 "We couldn't find your billing information. If you paid "
                 "recently, wait a few minutes for it to sync and try again. "
                 "If the problem persists, contact support@wepai.app."))
            return

        try:
            from storage import stripe_client
        except ImportError:
            stripe_client = None

        if not stripe_client or not stripe_client.is_configured():
            messagebox.showinfo(
                "WEP AI",
                ("Para gestionar tu suscripción contactá support@wepai.app."
                 if lbl == "es" else
                 "To manage your subscription, contact support@wepai.app."))
            return

        url = stripe_client.create_billing_portal_session(customer_id)
        if not url:
            messagebox.showinfo(
                "WEP AI",
                ("No pudimos abrir el portal. Reintentá en unos minutos."
                 if lbl == "es" else
                 "Could not open the billing portal. Please try again in a few minutes."))
            return
        webbrowser.open(url)

    # ── SETTINGS ──────────────────────────────────────────────────────────────
    def _open_settings(self):
        import subprocess
        w = ctk.CTkToplevel(self)
        w.title("WEP AI — " + self._t("settings"))
        w.geometry("460x520")
        w.resizable(False, False)
        w.configure(fg_color=BG)
        w.grab_set(); w.lift(); w.focus_force()
        w.attributes("-topmost", True)

        # Header
        hf = ctk.CTkFrame(w, fg_color=CARD, corner_radius=0)
        hf.pack(fill="x")
        ctk.CTkLabel(hf, text="⚙  " + self._t("settings"),
                     font=ctk.CTkFont("Arial", 16, "bold"),
                     text_color=WHITE).pack(pady=(14,2))
        ctk.CTkLabel(hf, text=("Preferencias de la aplicación" if self.lang[0]=="es"
                               else "Application preferences"),
                     font=ctk.CTkFont("Arial", 11), text_color=GRAY1).pack(pady=(0,12))

        scroll = ctk.CTkScrollableFrame(w, fg_color="transparent", corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        def section(title):
            f = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=12,
                             border_width=1, border_color=BORDER)
            f.pack(padx=20, pady=(0,10), fill="x")
            ctk.CTkLabel(f, text=title, font=ctk.CTkFont("Arial", 10, "bold"),
                         text_color=GRAY1, anchor="w").pack(padx=16, pady=(10,4), fill="x")
            return f

        # ── 1. IDIOMA ──
        sf1 = section("Idioma / Language")
        lrow = ctk.CTkFrame(sf1, fg_color="transparent")
        lrow.pack(padx=16, pady=(0,12), fill="x")
        lang_var = ctk.StringVar(value=self.lang[0])

        def on_lang_change():
            # Preview: highlight selected option
            pass

        for lbl, val in [("🇪🇸  Español","es"), ("🇺🇸  English","en")]:
            ctk.CTkRadioButton(
                lrow, text=lbl, variable=lang_var, value=val,
                font=ctk.CTkFont("Arial", 13),
                text_color=WHITE, fg_color=BLUE, hover_color=BLUE2,
                command=on_lang_change
            ).pack(side="left", padx=(0,24))

        # ── 2. MODELO ──
        sf2 = section("Modelo de IA / AI Model")
        mrow = ctk.CTkFrame(sf2, fg_color="transparent"); mrow.pack(padx=16, pady=(0,12), fill="x")
        ctk.CTkLabel(mrow, text="Claude Sonnet 4.6",
                     font=ctk.CTkFont("Arial", 13), text_color=WHITE).pack(side="left")
        ctk.CTkLabel(mrow, text="  ●  " + ("Activo" if self.lang[0]=="es" else "Active"),
                     font=ctk.CTkFont("Arial", 12), text_color=GREEN).pack(side="left")
        ctk.CTkLabel(sf2, text="claude-sonnet-4-6",
                     font=ctk.CTkFont("Arial", 10), text_color=GRAY2, anchor="w").pack(
            padx=16, pady=(0,10), fill="x")

        # ── 3. CARPETA ──
        sf3 = section("Carpeta de documentos / Documents folder")
        docs_path = os.path.expanduser("~/Documents/WEP_AI")
        os.makedirs(docs_path, exist_ok=True)
        frow = ctk.CTkFrame(sf3, fg_color="transparent"); frow.pack(padx=16, pady=(0,12), fill="x")
        ctk.CTkLabel(frow, text=docs_path, font=ctk.CTkFont("Arial", 11),
                     text_color=GRAY1, anchor="w").pack(side="left", fill="x", expand=True)
        ctk.CTkButton(frow, text="Abrir" if self.lang[0]=="es" else "Open",
                      width=70, height=30, corner_radius=8,
                      font=ctk.CTkFont("Arial", 11, "bold"),
                      fg_color=BLUE, hover_color=BLUE2,
                      command=lambda: subprocess.run(["open", docs_path])).pack(side="right")

        # ── 4. API KEY ──
        sf4 = section("API Key de Anthropic / Anthropic API Key")
        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        arow = ctk.CTkFrame(sf4, fg_color="transparent"); arow.pack(padx=16, pady=(0,4), fill="x")
        if api_key:
            masked = "sk-ant-..." + api_key[-6:]
            ctk.CTkLabel(arow, text=masked, font=ctk.CTkFont("Arial", 12),
                         text_color=GREEN).pack(side="left")
            ctk.CTkLabel(arow, text="  ✓  " + ("Configurada" if self.lang[0]=="es" else "Configured"),
                         font=ctk.CTkFont("Arial", 11), text_color=GREEN).pack(side="left")
        else:
            ctk.CTkLabel(arow, text="⚠  " + ("No configurada — agrega ANTHROPIC_API_KEY al entorno"
                                               if self.lang[0]=="es"
                                               else "Not configured — add ANTHROPIC_API_KEY to environment"),
                         font=ctk.CTkFont("Arial", 11), text_color=AMBER,
                         wraplength=380, justify="left").pack(fill="x")

        # Muestra créditos restantes si hay API key
        if api_key:
            try:
                ctk.CTkLabel(sf4,
                             text=("Cuenta activa en console.anthropic.com" if self.lang[0]=="es"
                                   else "Account active at console.anthropic.com"),
                             font=ctk.CTkFont("Arial", 10), text_color=GRAY2, anchor="w").pack(
                    padx=16, pady=(0,10), fill="x")
            except Exception as e:
                print(f"[WEP AI] settings label render failed: {e}")
        else:
            ctk.CTkButton(sf4,
                          text="console.anthropic.com →",
                          height=28, corner_radius=7, width=200,
                          font=ctk.CTkFont("Arial", 10, "bold"),
                          fg_color="transparent", border_width=1, border_color=AMBER,
                          text_color=AMBER,
                          command=lambda: subprocess.run(["open", "https://console.anthropic.com"])
                          ).pack(padx=16, pady=(0,10), anchor="w")

        # ── 5. PLAN ──
        plan = self.user.get("plan","free")
        pc = PLAN_COLORS.get(plan, GRAY1)
        pr = PLAN_PRICES if self.lang[0]=="es" else PLAN_PRICES_EN
        sf5 = section("Plan activo / Active plan")
        prow = ctk.CTkFrame(sf5, fg_color="transparent"); prow.pack(padx=16, pady=(0,12), fill="x")
        ctk.CTkLabel(prow, text=f"{PLAN_NAMES_ES.get(plan,'')} · {pr.get(plan,'')}",
                     font=ctk.CTkFont("Arial", 13, "bold"), text_color=pc).pack(side="left")
        ctk.CTkLabel(prow, text=("  — Para cambiar de plan visita wepai.app" if self.lang[0]=="es"
                                  else "  — Visit wepai.app to change plan"),
                     font=ctk.CTkFont("Arial", 10), text_color=GRAY2).pack(side="left")

        # ── BOTÓN GUARDAR ──
        bot = ctk.CTkFrame(w, fg_color="transparent")
        bot.pack(padx=20, pady=12, fill="x")

        def save_settings():
            new_l = lang_var.get()
            changed = new_l != self.lang[0]
            if changed:
                self.lang[0] = new_l
                self._set_lang(new_l)
            w.destroy()
            if changed:
                from tkinter import messagebox
                messagebox.showinfo("WEP AI",
                    "Idioma cambiado a Español" if new_l=="es" else "Language changed to English")

        bot.grid_columnconfigure((0,1), weight=1)
        ctk.CTkButton(bot, text="Cancelar" if self.lang[0]=="es" else "Cancel",
                      height=42, corner_radius=10,
                      font=ctk.CTkFont("Arial", 13, "bold"),
                      fg_color="transparent", border_width=1, border_color=BORDER,
                      text_color=GRAY1, hover_color=CARD2,
                      command=w.destroy).grid(row=0, column=0, padx=(0,8), sticky="ew")
        ctk.CTkButton(bot, text="Guardar" if self.lang[0]=="es" else "Save",
                      height=42, corner_radius=10,
                      font=ctk.CTkFont("Arial", 13, "bold"),
                      fg_color=BLUE, hover_color=BLUE2,
                      command=save_settings).grid(row=0, column=1, sticky="ew")

    # ── MY DOCUMENTS ──────────────────────────────────────────────────────────
    def _open_docs_window(self):
        import subprocess
        # Abrir Finder
        path = os.path.expanduser("~/Documents/WEP_AI")
        os.makedirs(path, exist_ok=True)
        subprocess.run(["open", path])

        # Ventana en-app con lista de documentos
        w = ctk.CTkToplevel(self)
        w.title("WEP AI — " + self._t("my_docs"))
        w.geometry("520x480")
        w.resizable(False, True)
        w.configure(fg_color=BG)
        w.grab_set(); w.lift(); w.focus_force()
        w.attributes("-topmost", True)

        ctk.CTkLabel(w, text=self._t("my_docs"),
                     font=ctk.CTkFont("Arial", 17, "bold"), text_color=WHITE).pack(pady=(20,4))
        ctk.CTkLabel(w, text=path, font=ctk.CTkFont("Arial", 10), text_color=GRAY1).pack(pady=(0,14))

        # Listar archivos
        doc_scroll = ctk.CTkScrollableFrame(w, fg_color=BG, corner_radius=0)
        doc_scroll.pack(fill="both", expand=True, padx=20, pady=(0,10))

        icon_map  = {".docx":"📄 Word",".xlsx":"📊 Excel",".pptx":"📑 PowerPoint"}
        color_map = {".docx":WORD_C,   ".xlsx":EXCEL_C,  ".pptx":PPT_C}

        files = []
        if os.path.exists(path):
            files = sorted(
                [f for f in os.listdir(path) if f.endswith((".docx",".xlsx",".pptx"))],
                key=lambda f: os.path.getmtime(os.path.join(path,f)), reverse=True)

        if not files:
            ctk.CTkLabel(doc_scroll,
                         text=("No hay documentos generados aún.\nCrea tu primer documento con WEP AI." if self.lang[0]=="es"
                                else "No documents generated yet.\nCreate your first document with WEP AI."),
                         font=ctk.CTkFont("Arial", 13), text_color=GRAY2,
                         justify="center").pack(pady=40)
        else:
            for fname in files:
                ext   = os.path.splitext(fname)[1]
                label = icon_map.get(ext, "📄")
                col   = color_map.get(ext, GRAY1)
                fpath = os.path.join(path, fname)
                mtime = datetime.fromtimestamp(os.path.getmtime(fpath)).strftime("%d/%m/%Y %H:%M")
                size  = f"{os.path.getsize(fpath)//1024} KB"

                row = ctk.CTkFrame(doc_scroll, fg_color=CARD, corner_radius=10,
                                   border_width=1, border_color=BORDER)
                row.pack(fill="x", pady=4)
                row.grid_columnconfigure(1, weight=1)

                ctk.CTkLabel(row, text=label.split()[0], font=ctk.CTkFont("Arial", 24)).grid(
                    row=0, column=0, rowspan=2, padx=12, pady=10, sticky="w")
                ctk.CTkLabel(row, text=fname, font=ctk.CTkFont("Arial", 11, "bold"),
                             text_color=WHITE, anchor="w").grid(row=0, column=1, sticky="w", padx=4, pady=(10,2))
                ctk.CTkLabel(row, text=f"{label.split(None,1)[1]}  ·  {mtime}  ·  {size}",
                             font=ctk.CTkFont("Arial", 10), text_color=GRAY1, anchor="w").grid(
                    row=1, column=1, sticky="w", padx=4, pady=(0,8))

                def open_file(p=fpath):
                    subprocess.run(["open", p])

                ctk.CTkButton(row, text="Abrir" if self.lang[0]=="es" else "Open",
                              width=64, height=30, corner_radius=7,
                              font=ctk.CTkFont("Arial", 10, "bold"),
                              fg_color=col, hover_color=BLUE2,
                              command=open_file).grid(row=0, column=2, rowspan=2,
                                                       padx=12, pady=10)

        # Botón Abrir en Finder
        ctk.CTkButton(w, text=("Abrir carpeta en Finder" if self.lang[0]=="es" else "Open folder in Finder"),
                      height=40, corner_radius=10, width=220,
                      font=ctk.CTkFont("Arial", 12, "bold"),
                      fg_color=CARD2, hover_color=CARD, border_width=1, border_color=BORDER,
                      text_color=GRAY1,
                      command=lambda: subprocess.run(["open", path])).pack(pady=(0,16))


    def _show_onboarding(self):
        w = ctk.CTkToplevel(self)
        w.title("WEP AI"); w.geometry("580x500")
        w.resizable(False, False); w.configure(fg_color=BG)
        w.grab_set(); w.lift(); w.focus_force(); w.attributes("-topmost", True)
        l = self.lang[0]
        top_ = ctk.CTkFrame(w, fg_color="transparent"); top_.pack(pady=(28,4))
        row_ = ctk.CTkFrame(top_, fg_color="transparent"); row_.pack()
        for ch, col in [("W",WORD_C),("E",EXCEL_C),("P",PPT_C)]:
            ctk.CTkLabel(row_,text=ch,font=ctk.CTkFont("Arial",44,"bold"),text_color=col).pack(side="left")
        ctk.CTkLabel(row_,text=" AI",font=ctk.CTkFont("Arial",26,"normal"),text_color=WHITE).pack(side="left")
        ctk.CTkLabel(w,text="Bienvenido/a a WEP AI" if l=="es" else "Welcome to WEP AI",
                     font=ctk.CTkFont("Arial",18,"bold"),text_color=WHITE).pack(pady=(8,3))
        ctk.CTkLabel(w,text="Tu asistente de documentos con IA para macOS" if l=="es" else "Your AI document assistant for macOS",
                     font=ctk.CTkFont("Arial",12),text_color=GRAY1).pack(pady=(0,20))
        cards_f = ctk.CTkFrame(w, fg_color="transparent"); cards_f.pack(padx=28, fill="x")
        cards_f.grid_columnconfigure((0,1,2), weight=1)
        caps = [
            ("📄","Word",WORD_C,"Contratos, cartas,\nlegales e informes" if l=="es" else "Contracts, letters,\nlegal & reports"),
            ("📊","Excel",EXCEL_C,"Inventarios, nóminas,\npresupuestos y KPIs" if l=="es" else "Inventory, payroll,\nbudgets & KPIs"),
            ("📑","PowerPoint",PPT_C,"Pitch decks,\npresentaciones, training" if l=="es" else "Pitch decks,\npresentations, training"),
        ]
        for i,(ico,name,col,desc) in enumerate(caps):
            cf_ = ctk.CTkFrame(cards_f,fg_color=CARD,corner_radius=13,border_width=1,border_color=BORDER)
            cf_.grid(row=0,column=i,padx=6,sticky="nsew")
            ctk.CTkLabel(cf_,text=ico,font=ctk.CTkFont("Arial",30)).pack(pady=(14,4))
            ctk.CTkLabel(cf_,text=name,font=ctk.CTkFont("Arial",13,"bold"),text_color=col).pack(pady=(0,4))
            ctk.CTkLabel(cf_,text=desc,font=ctk.CTkFont("Arial",10),text_color=GRAY1,justify="center").pack(padx=10,pady=(0,14))
        tip = ctk.CTkFrame(w,fg_color=CARD2,corner_radius=11,border_width=1,border_color=BORDER)
        tip.pack(padx=28,pady=(18,0),fill="x")
        ctk.CTkLabel(tip,
                     text=("💡 Tip: Sé específico. En vez de 'un contrato', di: 'Contrato de arrendamiento\npara bodega en Medellín por 18 meses con opción de renovación automática'" if l=="es"
                            else "💡 Tip: Be specific. Instead of 'a contract', say: 'Commercial warehouse lease\nin Miami for 18 months with automatic renewal option'"),
                     font=ctk.CTkFont("Arial",10),text_color=GRAY1,justify="center",wraplength=500).pack(padx=16,pady=12)
        ctk.CTkButton(w,text="Empezar ahora →" if l=="es" else "Get started →",
                      height=46,corner_radius=12,width=200,
                      font=ctk.CTkFont("Arial",14,"bold"),fg_color=BLUE,hover_color=BLUE2,
                      command=lambda:(w.destroy(),self._prefs.update({"first_launch":False}),
                                      self._save_prefs(),self._new_chat())).pack(pady=18)

    def _open_help(self):
        w = ctk.CTkToplevel(self)
        w.title("WEP AI — " + self._t("help_menu"))
        w.geometry("540x580"); w.resizable(False, True)
        w.configure(fg_color=BG); w.grab_set(); w.lift(); w.focus_force(); w.attributes("-topmost", True)
        l = self.lang[0]
        hf = ctk.CTkFrame(w, fg_color=CARD, corner_radius=0); hf.pack(fill="x")
        ctk.CTkLabel(hf,text="❓  "+self._t("help_menu"),font=ctk.CTkFont("Arial",16,"bold"),text_color=WHITE).pack(pady=(14,2))
        ctk.CTkLabel(hf,text=("Aprende a sacarle el máximo provecho a WEP AI" if l=="es" else "Learn how to get the most out of WEP AI"),
                     font=ctk.CTkFont("Arial",11),text_color=GRAY1).pack(pady=(0,12))
        scroll = ctk.CTkScrollableFrame(w, fg_color=BG, corner_radius=0)
        scroll.pack(fill="both", expand=True)
        faqs = [
            ("¿Cómo creo un documento?" if l=="es" else "How do I create a document?",
             "Selecciona Word, Excel o PowerPoint con los botones de la barra de entrada, escribe lo que necesitas con todos los detalles y presiona Cmd+Enter." if l=="es"
             else "Select Word, Excel or PowerPoint using the input bar buttons, describe what you need in detail and press Cmd+Enter."),
            ("¿Cómo logro mejores resultados?" if l=="es" else "How do I get better results?",
             "Sé específico. En vez de 'un contrato', di: 'Contrato de arrendamiento comercial para bodega de 200m² en Medellín por 18 meses con cláusula de renovación automática y depósito de 2 meses.'" if l=="es"
             else "Be specific. Instead of 'a contract', say: 'Commercial warehouse lease agreement for 2000 sq ft in Miami for 18 months with automatic renewal clause and 2-month deposit.'"),
            ("¿Puedo corregir el documento?" if l=="es" else "Can I correct the document?",
             "Sí. Una vez abierto en Office, responde en el chat con los cambios que necesitas. Por ejemplo: 'Agrega una cláusula de confidencialidad' o 'Cambia el encabezado a azul'." if l=="es"
             else "Yes. Once open in Office, reply in the chat with the changes you need. For example: 'Add a confidentiality clause' or 'Change the header to blue'."),
            ("¿Dónde se guardan mis documentos?" if l=="es" else "Where are my documents saved?",
             "En ~/Documents/WEP_AI en tu Mac. Puedes cambiar esta carpeta en Configuración → Carpeta de documentos." if l=="es"
             else "In ~/Documents/WEP_AI on your Mac. You can change this folder in Settings → Documents folder."),
            ("¿Qué hace la API Key?" if l=="es" else "What does the API Key do?",
             "Conecta WEP AI con Claude (el modelo de IA de Anthropic). La obtienes en console.anthropic.com y la configuras en Configuración → API Key." if l=="es"
             else "It connects WEP AI to Claude (Anthropic's AI model). Get it at console.anthropic.com and configure it in Settings → API Key."),
        ]
        for q,a in faqs:
            qf = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=11, border_width=1, border_color=BORDER)
            qf.pack(padx=18, pady=(0,8), fill="x")
            ctk.CTkLabel(qf,text=f"{'P' if l=='es' else 'Q'}:  {q}",font=ctk.CTkFont("Arial",12,"bold"),
                         text_color=BLUE,anchor="w",wraplength=460).pack(padx=16,pady=(12,4),fill="x")
            ctk.CTkLabel(qf,text=a,font=ctk.CTkFont("Arial",11),text_color=GRAY1,
                         anchor="w",wraplength=460,justify="left").pack(padx=16,pady=(0,12),fill="x")
        # Atajos
        sf = ctk.CTkFrame(scroll, fg_color=CARD, corner_radius=11, border_width=1, border_color=BORDER)
        sf.pack(padx=18, pady=(0,18), fill="x")
        ctk.CTkLabel(sf,text="⌨️  "+("Atajos de teclado" if l=="es" else "Keyboard shortcuts"),
                     font=ctk.CTkFont("Arial",12,"bold"),text_color=WHITE,anchor="w").pack(padx=16,pady=(12,8),fill="x")
        for k,d in [("⌘+Enter","Enviar / Send"),("⌘+N","Nueva conv. / New conv."),("⌘+,","Configuración / Settings")]:
            r = ctk.CTkFrame(sf, fg_color="transparent"); r.pack(padx=16,pady=3,fill="x")
            ctk.CTkLabel(r,text=k,font=ctk.CTkFont("Arial",11,"bold"),text_color=WHITE,
                         fg_color=CARD2,corner_radius=5,width=90).pack(side="left",padx=(0,12))
            ctk.CTkLabel(r,text=d.split("/")[0 if l=="es" else 1].strip(),
                         font=ctk.CTkFont("Arial",11),text_color=GRAY1).pack(side="left")
        ctk.CTkFrame(sf, height=8, fg_color="transparent").pack()

    def _open_about(self):
        w = ctk.CTkToplevel(self)
        w.title("WEP AI"); w.geometry("380x430")
        w.resizable(False, False); w.configure(fg_color=BG)
        w.grab_set(); w.lift(); w.focus_force(); w.attributes("-topmost", True)
        l = self.lang[0]
        row_ = ctk.CTkFrame(w, fg_color="transparent"); row_.pack(pady=(28,4))
        rl = ctk.CTkFrame(row_, fg_color="transparent"); rl.pack()
        for ch, col in [("W",WORD_C),("E",EXCEL_C),("P",PPT_C)]:
            ctk.CTkLabel(rl,text=ch,font=ctk.CTkFont("Arial",40,"bold"),text_color=col).pack(side="left")
        ctk.CTkLabel(rl,text=" AI",font=ctk.CTkFont("Arial",24,"normal"),text_color=WHITE).pack(side="left")
        ctk.CTkLabel(w,text="Versión 1.1 · Fase 1" if l=="es" else "Version 1.1 · Phase 1",
                     font=ctk.CTkFont("Arial",11),text_color=GRAY1).pack(pady=(4,20))
        info_f = ctk.CTkFrame(w, fg_color=CARD, corner_radius=13, border_width=1, border_color=BORDER)
        info_f.pack(padx=28, fill="x", pady=(0,14))
        rows_ = [("Motor de IA" if l=="es" else "AI Engine","Claude Sonnet 4 · Anthropic"),
                 ("Versión" if l=="es" else "Version","1.1.0"),
                 ("Plataforma" if l=="es" else "Platform","macOS 13+"),
                 ("Documentos" if l=="es" else "Documents","Word · Excel · PowerPoint"),
                 ("Base de datos" if l=="es" else "Database","SQLite · Local"),
                 ("Hecho con" if l=="es" else "Built with","Python · CustomTkinter")]
        for i,(k,v) in enumerate(rows_):
            rf_ = ctk.CTkFrame(info_f, fg_color="transparent"); rf_.pack(fill="x", padx=16, pady=4)
            ctk.CTkLabel(rf_,text=k,font=ctk.CTkFont("Arial",11,"bold"),text_color=GRAY1,width=140,anchor="w").pack(side="left")
            ctk.CTkLabel(rf_,text=v,font=ctk.CTkFont("Arial",11),text_color=WHITE,anchor="w").pack(side="left")
            if i<len(rows_)-1: ctk.CTkFrame(info_f,height=0.5,fg_color=BORDER).pack(fill="x",padx=16)
        ctk.CTkLabel(w,text="© 2026 WEP AI · Word · Excel · PowerPoint con IA",
                     font=ctk.CTkFont("Arial",9),text_color=GRAY3).pack(pady=4)

    def _logout(self):
        # Ventana de confirmación personalizada
        w = ctk.CTkToplevel(self)
        w.title("WEP AI")
        w.geometry("340x200")
        w.resizable(False, False)
        w.configure(fg_color=BG)
        w.grab_set(); w.lift(); w.focus_force()
        w.attributes("-topmost", True)

        ctk.CTkLabel(w, text="→", font=ctk.CTkFont("Arial", 32), text_color=RED).pack(pady=(24,4))
        ctk.CTkLabel(w, text=self._t("confirm_logout"),
                     font=ctk.CTkFont("Arial", 14, "bold"), text_color=WHITE).pack(pady=(0,4))
        nm = self.user.get("name","")
        ctk.CTkLabel(w, text=(f"Sesión de {nm}" if self.lang[0]=="es" else f"{nm}'s session"),
                     font=ctk.CTkFont("Arial", 11), text_color=GRAY1).pack(pady=(0,20))

        btn_row = ctk.CTkFrame(w, fg_color="transparent"); btn_row.pack(padx=24, fill="x")
        btn_row.grid_columnconfigure((0,1), weight=1)

        ctk.CTkButton(btn_row,
                      text="Cancelar" if self.lang[0]=="es" else "Cancel",
                      height=40, corner_radius=10,
                      font=ctk.CTkFont("Arial", 12, "bold"),
                      fg_color="transparent", border_width=1, border_color=BORDER,
                      text_color=GRAY1, hover_color=CARD2,
                      command=w.destroy).grid(row=0, column=0, padx=(0,6), sticky="ew")

        def do_logout():
            w.destroy(); self.destroy(); run()

        ctk.CTkButton(btn_row,
                      text="Cerrar sesión" if self.lang[0]=="es" else "Sign out",
                      height=40, corner_radius=10,
                      font=ctk.CTkFont("Arial", 12, "bold"),
                      fg_color=RED, hover_color="#cc0000", text_color=WHITE,
                      command=do_logout).grid(row=0, column=1, padx=(6,0), sticky="ew")

    # ── CONVERSATIONS ─────────────────────────────────────────────────────────
    def _conv_menu(self, cid, title):
        l = self.lang[0]
        m = ctk.CTkToplevel(self); m.overrideredirect(True)
        m.configure(fg_color=CARD2); m.attributes("-topmost", True)
        m.geometry(f"185x78+{self.winfo_rootx()+220}+{self.winfo_rooty()+280}")
        ctk.CTkFrame(m, height=1, fg_color=BORDER).pack(fill="x")
        def rename_c():
            m.destroy()
            d = ctk.CTkInputDialog(text=("Nuevo nombre:" if l=="es" else "New name:"), title="WEP AI")
            new_title = d.get_input()
            if new_title and new_title.strip():
                db.update_conversation_title(cid, new_title.strip()); self._load_conversations()
        def delete_c():
            m.destroy()
            if messagebox.askyesno("WEP AI", self._t("confirm_delete")):
                con=db._con(); cur=con.cursor()
                cur.execute("DELETE FROM messages WHERE conversation_id=?",(cid,))
                cur.execute("DELETE FROM documents WHERE conversation_id=?",(cid,))
                cur.execute("DELETE FROM conversations WHERE id=?",(cid,))
                con.commit(); con.close()
                self._load_conversations(); self._new_chat()
        for icon,txt,cmd in [("✏️",self._t("rename_conv"),rename_c),("🗑️",self._t("delete_conv"),delete_c)]:
            ctk.CTkButton(m, text=f"{icon}  {txt}", anchor="w", height=34,
                          corner_radius=0, font=ctk.CTkFont("Arial",11),
                          fg_color="transparent", hover_color=CARD,
                          text_color=GRAY1, command=cmd).pack(fill="x", padx=2, pady=1)
        m.bind("<FocusOut>", lambda e: m.destroy()); m.focus_set()

    def _load_conversations(self):
        for b in self.conv_buttons: b.destroy()
        self.conv_buttons.clear()
        for cid, title, updated in db.get_recent_conversations(self.user["id"]):
            ds = updated[:10] if updated else ""
            frm = ctk.CTkFrame(self.conv_scroll, fg_color="transparent")
            frm.pack(fill="x", padx=4, pady=1)
            b = ctk.CTkButton(frm, text=f"{title[:26]}\n{ds}",
                              anchor="w", height=46, corner_radius=8,
                              font=ctk.CTkFont("Arial", 11),
                              fg_color="transparent", text_color=GRAY1,
                              hover_color=CARD2,
                              command=lambda c=cid, t=title: self._load_conv(c, t))
            b.pack(side="left", fill="x", expand=True)
            mb = ctk.CTkButton(frm, text="···", width=26, height=26,
                               corner_radius=6, font=ctk.CTkFont("Arial",10,"bold"),
                               fg_color="transparent", text_color=GRAY2, hover_color=CARD2,
                               command=lambda c=cid,t=title: self._conv_menu(c,t))
            mb.pack(side="right", padx=2)
            self.conv_buttons.append(frm)

    def _new_chat(self):
        self.conversation_id  = db.new_conversation(self.user["id"])
        self.chat_history     = []
        self.current_doc_path = self.current_gen_data = self.current_doc_type = None
        self.chat_title.configure(text=""); self.doc_badge.configure(text="")
        self._clr_area(); self._welcome(); self._load_conversations()
        self.open_btn.configure(state="disabled")
        self.prev_lbl.configure(image=None, text=self._t("preview_sub"))
        self._set_state("idle")

    def _load_conv(self, cid, title):
        self.conversation_id = cid
        self.chat_history    = db.get_messages(cid)
        self.chat_title.configure(text=title)
        self._clr_area()
        for msg in self.chat_history:
            self._bubble(msg["role"], msg["content"])
        docs = db.get_documents(cid)
        if docs:
            _, dt, dp, _ = docs[0]
            if os.path.exists(dp):
                self.current_doc_path = dp; self.current_doc_type = dt
                self._badge(dt); self.open_btn.configure(state="normal")

    def _clr_area(self):
        for w in self.chat_scroll.winfo_children(): w.destroy()

    def _welcome(self):
        f = ctk.CTkFrame(self.chat_scroll, fg_color="transparent")
        f.grid(sticky="ew", padx=28, pady=(28,10))
        f.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(f, text=("¿Qué vamos a crear hoy?" if self.lang[0]=="es" else "What are we creating today?"),
                     font=ctk.CTkFont("Arial",20,"bold"), text_color=WHITE).grid(pady=(0,5))
        ctk.CTkLabel(f, text=("Elige el tipo de documento y descríbelo con todos los detalles."
                               if self.lang[0]=="es"
                               else "Choose the document type and describe it in detail."),
                     font=ctk.CTkFont("Arial",13), text_color=GRAY1, justify="center").grid(pady=(0,18))

        # Tarjetas de tipo
        cards_f = ctk.CTkFrame(f, fg_color="transparent"); cards_f.grid(sticky="ew", pady=(0,18))
        cards_f.grid_columnconfigure((0,1,2), weight=1)
        doc_types = [
            ("word","📄",WORD_C,"Word",
             "Contratos, cartas, legales, políticas" if self.lang[0]=="es" else "Contracts, letters, legal, policies"),
            ("excel","📊",EXCEL_C,"Excel",
             "Inventarios, nóminas, presupuestos, KPIs" if self.lang[0]=="es" else "Inventory, payroll, budgets, KPIs"),
            ("powerpoint","📑",PPT_C,"PowerPoint",
             "Pitch decks, presentaciones, capacitaciones" if self.lang[0]=="es" else "Pitch decks, presentations, training"),
        ]
        for i,(dtype,ico,col,name,desc) in enumerate(doc_types):
            is_sel = (dtype == self.selected_doc_type)
            cf2 = ctk.CTkFrame(cards_f, fg_color=CARD2 if is_sel else CARD,
                              corner_radius=12, border_width=2 if is_sel else 1,
                              border_color=col if is_sel else BORDER, cursor="hand2")
            cf2.grid(row=0, column=i, padx=6, sticky="nsew")
            ctk.CTkLabel(cf2, text=ico, font=ctk.CTkFont("Arial",26)).pack(pady=(14,4))
            ctk.CTkLabel(cf2, text=name, font=ctk.CTkFont("Arial",13,"bold"), text_color=col).pack(pady=(0,4))
            ctk.CTkLabel(cf2, text=desc, font=ctk.CTkFont("Arial",9),
                         text_color=GRAY1, wraplength=140, justify="center").pack(padx=10, pady=(0,12))
            cf2.bind("<Button-1>", lambda e,d=dtype: (self._select_doc_type(d), self._new_chat()))
            for child in cf2.winfo_children():
                child.bind("<Button-1>", lambda e,d=dtype: (self._select_doc_type(d), self._new_chat()))

        # Ejemplos
        examples = (EXAMPLES_ES if self.lang[0]=="es" else EXAMPLES_EN).get(self.selected_doc_type,[])
        ef = ctk.CTkFrame(f, fg_color="transparent"); ef.grid(sticky="ew")
        ctk.CTkLabel(ef, text=("💡 Ejemplos para empezar — haz clic en uno:"
                                if self.lang[0]=="es" else "💡 Examples to get started — click one:"),
                     font=ctk.CTkFont("Arial",11,"bold"), text_color=GRAY1, anchor="w").pack(fill="x", pady=(0,6))
        for ex in examples[:4]:
            exf = ctk.CTkFrame(ef, fg_color=CARD, corner_radius=8,
                               border_width=1, border_color=BORDER, cursor="hand2")
            exf.pack(fill="x", pady=3)
            lbl = ctk.CTkLabel(exf, text=f"  💬  {ex}", font=ctk.CTkFont("Arial",11),
                               text_color=GRAY1, anchor="w", cursor="hand2")
            lbl.pack(fill="x", padx=8, pady=8)
            def use_example(txt=ex):
                try:
                    self.inp.delete("1.0","end")
                    self.inp.insert("1.0",txt)
                    self.inp.configure(text_color=WHITE)
                except Exception as e:
                    print(f"[WEP AI] use_example failed: {e}")
            exf.bind("<Button-1>", lambda e,fn=use_example: fn())
            lbl.bind("<Button-1>",  lambda e,fn=use_example: fn())
            exf.bind("<Enter>", lambda e,fr=exf: fr.configure(fg_color=CARD2))
            exf.bind("<Leave>", lambda e,fr=exf: fr.configure(fg_color=CARD))
            lbl.bind("<Enter>",  lambda e,fr=exf: fr.configure(fg_color=CARD2))
            lbl.bind("<Leave>",  lambda e,fr=exf: fr.configure(fg_color=CARD))

    def _bubble(self, role, text, new=False):
        isu = role == "user"
        rf  = ctk.CTkFrame(self.chat_scroll, fg_color="transparent")
        rf.grid(sticky="ew", padx=12, pady=5)
        rf.grid_columnconfigure(0 if isu else 1, weight=1)
        init = ("Tú" if self.lang[0]=="es" else "You") if isu else "AI"
        av   = ctk.CTkLabel(rf, text=init[:2], width=30, height=30,
                            font=ctk.CTkFont("Arial", 10, "bold"),
                            fg_color=BLUE if isu else TEAL,
                            text_color=WHITE, corner_radius=15)
        av.grid(row=0, column=1 if isu else 0,
                padx=(6, 0) if isu else (0, 6), sticky="ne" if isu else "nw")
        lines = text.count('\n') + len(text) // 50 + 1
        b = ctk.CTkTextbox(rf, font=ctk.CTkFont("Arial", 12),
                           fg_color=BLUE if isu else CARD,
                           text_color=WHITE if isu else GRAY1,
                           corner_radius=12, wrap="word",
                           activate_scrollbars=False, border_width=0,
                           height=max(38, lines * 19 + 16))
        b.insert("1.0", text); b.configure(state="disabled")
        b.grid(row=0, column=0 if isu else 1, sticky="e" if isu else "w",
               padx=(60, 0) if isu else (0, 60))
        if new:
            self.chat_scroll.update()
            self.chat_scroll._parent_canvas.yview_moveto(1.0)

    def _typing(self):
        f = ctk.CTkFrame(self.chat_scroll, fg_color="transparent")
        f.grid(sticky="ew", padx=12, pady=5)
        ctk.CTkLabel(f, text="AI", width=30, height=30,
                     font=ctk.CTkFont("Arial", 10, "bold"),
                     fg_color=TEAL, text_color=WHITE,
                     corner_radius=15).grid(row=0, column=0, padx=(0, 6), sticky="nw")
        ctk.CTkLabel(f, text="● ● ●", font=ctk.CTkFont("Arial", 15),
                     text_color=GRAY2).grid(row=0, column=1, padx=4, pady=8)
        self.chat_scroll.update()
        self.chat_scroll._parent_canvas.yview_moveto(1.0)
        return f

    def _doc_card(self, doc_type, title, filepath):
        cb, _ = {"word": (BLUE, WHITE), "excel": (GREEN, WHITE),
                  "powerpoint": (AMBER, WHITE)}.get(doc_type, (BLUE, WHITE))
        icons = {"word": "📄", "excel": "📊", "powerpoint": "📑"}
        labs  = {"word": "Word", "excel": "Excel", "powerpoint": "PowerPoint"}
        f = ctk.CTkFrame(self.chat_scroll, fg_color=CARD, corner_radius=12,
                         border_width=1, border_color=BORDER)
        f.grid(sticky="ew", padx=12, pady=5)
        f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(f, text=icons.get(doc_type, "📄"),
                     font=ctk.CTkFont("Arial", 28)).grid(
            row=0, column=0, rowspan=2, padx=12, pady=12, sticky="w")
        ctk.CTkLabel(f, text=title,
                     font=ctk.CTkFont("Arial", 12, "bold"),
                     text_color=WHITE, anchor="w").grid(
            row=0, column=1, padx=4, pady=(12, 2), sticky="w")
        ctk.CTkLabel(f, text=f"{labs.get(doc_type,'Doc')} · {os.path.basename(filepath)}",
                     font=ctk.CTkFont("Arial", 10), text_color=GRAY1, anchor="w").grid(
            row=1, column=1, padx=4, pady=(0, 4), sticky="w")
        lbl = ("Abrir en" if self.lang[0]=="es" else "Open in") + " " + labs.get(doc_type,"")
        ctk.CTkButton(f, text=lbl, height=30, corner_radius=8,
                      font=ctk.CTkFont("Arial", 10, "bold"),
                      fg_color=cb, hover_color=BLUE2,
                      command=self._open_in_office).grid(
            row=0, column=2, rowspan=2, padx=12, pady=12)

        # v15.16.3 — MEJORA 4: sello de verificación legal. Hace visible el
        # diferenciador del producto: muestra que el documento se generó con
        # datos legales del país (y, si aplica, verificados/actualizados por IA).
        self._verif_badge(f, doc_type)

        self.chat_scroll.update()
        self.chat_scroll._parent_canvas.yview_moveto(1.0)

    def _verif_badge(self, parent, doc_type):
        """Muestra un sello verde con los marcadores legales del país usados."""
        try:
            gd  = self.current_gen_data or {}
            datos = gd.get("datos", {}) if isinstance(gd, dict) else {}
            pais_in = (datos.get("pais") or gd.get("pais") or "").strip()
            if not pais_in:
                return  # sin país detectado, no mostramos sello
            from office.legal_config import _get_legal_config, LEGAL_CONFIG
            # resolver cfg con overlay (incluye actualizaciones persistentes)
            code = pais_in.upper() if pais_in.upper() in LEGAL_CONFIG else None
            cfg = _get_legal_config(pais_in) if not code else _get_legal_config(code)
            pais_nombre = cfg.get("pais", pais_in)
            # construir lista de marcadores legales relevantes según el tipo de doc
            marcas = []
            for k in ("ley_laboral", "iva_label", "autoridad_fiscal", "seguridad_social"):
                v = cfg.get(k)
                if v:
                    # acortar nombres largos: usar siglas entre paréntesis si existen
                    import re as _re
                    m = _re.search(r"\(([^)]+)\)", str(v))
                    marcas.append(m.group(1) if m else str(v).split()[0])
            iva = cfg.get("iva")
            if iva:
                marcas.append(f"{cfg.get('iva_label','IVA')} {iva}")
            actualizado = cfg.get("_overlay_applied") or cfg.get("_fiscal_verified")
            mes = datetime.now().strftime("%b %Y")
            detalle = " · ".join(dict.fromkeys(marcas[:4]))  # sin duplicados, máx 4
            titulo_v = (("✓ Verificado con leyes vigentes de " + pais_nombre)
                        if self.lang[0]=="es"
                        else ("✓ Verified against current laws of " + pais_nombre))
            if actualizado:
                titulo_v += (" · actualizado por IA" if self.lang[0]=="es" else " · AI-updated")

            badge = ctk.CTkFrame(parent, fg_color="#0d2818", corner_radius=8,
                                 border_width=1, border_color="#1a5c3a")
            badge.grid(row=2, column=0, columnspan=3, padx=12, pady=(0,12), sticky="ew")
            ctk.CTkLabel(badge, text=titulo_v,
                         font=ctk.CTkFont("Arial", 10, "bold"), text_color=GREEN,
                         anchor="w").pack(fill="x", padx=10, pady=(7,1))
            if detalle:
                ctk.CTkLabel(badge, text=f"{detalle} · {mes}",
                             font=ctk.CTkFont("Arial", 9), text_color=GRAY1,
                             anchor="w").pack(fill="x", padx=10, pady=(0,7))
        except Exception:
            pass  # nunca romper la generación por el sello

    # ── SEND ──────────────────────────────────────────────────────────────────
    #
    # Patrón thread-safe para Tkinter (v14.2):
    # ─────────────────────────────────────────
    # Tkinter NO es thread-safe. Cualquier mutación de widget desde un hilo
    # distinto del main loop puede causar freezes, race conditions o crashes
    # (especialmente en macOS).
    #
    # Patrón aplicado:
    #   1. Métodos `_send`, `_start_agent` corren en MAIN THREAD. Hacen
    #      todas las preparaciones de UI (estados, bubbles iniciales,
    #      indicadores de typing) y luego lanzan un worker.
    #   2. Workers `_process_worker`, `_agent_worker` corren en BACKGROUND.
    #      Hacen sólo trabajo de red/CPU/subprocess (brain.chat, generación
    #      de documentos, screenshots). NUNCA tocan widgets.
    #   3. Resultados se entregan al main thread vía `self.after(0, cb, *args)`.
    #      Los callbacks `_on_*` reciben los datos y actualizan la UI.
    #
    # Esta separación elimina toda llamada Tkinter desde threads no-principales.
    # ─────────────────────────────────────────────────────────────────────

    def _send(self, e=None):
        text = self.inp.get("1.0", "end-1c").strip()
        if not text or text == self._t("inp_ph"): return
        if self.app_state not in ["idle", "presenting"]: return
        self.inp.delete("1.0", "end")
        self.send_btn.configure(state="disabled")
        self._bubble("user", text, new=True)
        db.save_message(self.conversation_id, "user", text)
        self.chat_history.append({"role": "user", "content": text})
        if len(self.chat_history) == 1:
            threading.Thread(target=self._upd_title, args=(text,), daemon=True).start()

        # Preparación de UI ANTES de lanzar el worker (main thread).
        self._set_state("thinking")
        typ = self._typing()

        # Snapshot del historial para el worker — evita race si main thread modifica.
        history_snapshot = list(self.chat_history[:-1])

        threading.Thread(
            target=self._process_worker,
            args=(text, history_snapshot, typ),
            daemon=True,
        ).start()

    def _upd_title(self, text):
        """Background thread: genera título via API y refresca el UI."""
        try:
            t = brain.generate_title(text)
        except Exception as e:
            print(f"[WEP AI] generate_title failed: {e}")
            return
        try:
            db.update_conversation_title(self.conversation_id, t)
        except Exception as e:
            print(f"[WEP AI] update_conversation_title failed: {e}")
            return
        # UI siempre en main thread.
        self.after(0, lambda: self.chat_title.configure(text=t))
        self.after(0, self._load_conversations)

    # ── PROCESS (chat API call) ───────────────────────────────────────────────
    def _process_worker(self, txt, history, typ):
        """BACKGROUND thread. Sólo llama a la API; entrega resultado al main thread."""
        try:
            # v14.6 — pasar el idioma actual de la UI como fallback
            res = brain.chat(history, txt, ui_lang=self.lang[0])
        except Exception as e:
            self.after(0, self._on_chat_error, typ, e)
            return
        self.after(0, self._on_chat_done, typ, res)

    def _on_chat_error(self, typ, err):
        """MAIN thread. Cancela el indicador y muestra el error."""
        try:
            typ.destroy()
        except Exception:
            pass
        self._bubble("assistant", f"{self._t('err_api')}: {err}", new=True)
        self._set_state("idle")
        self.send_btn.configure(state="normal")

    def _on_chat_done(self, typ, res):
        """MAIN thread. Procesa la respuesta del modelo."""
        try:
            typ.destroy()
        except Exception:
            pass
        if res["text"]:
            self._bubble("assistant", res["text"], new=True)
            db.save_message(self.conversation_id, "assistant", res["full_text"])
            self.chat_history.append({"role": "assistant", "content": res["full_text"]})
        if res["generate"] and res["gen_data"]:
            self._start_agent(res["gen_data"])
        else:
            self._set_state("idle")
            self.send_btn.configure(state="normal")

    # ── AGENT (document generation) ───────────────────────────────────────────
    def _start_agent(self, gen_data):
        """MAIN thread. Verifica cupo, prepara UI, lanza el worker."""
        # Atomic check+consume — antes era check_can_generate + increment_docs_used.
        can, reason = db.check_and_consume_doc(self.user["id"])
        if not can:
            self._bubble("assistant", self._t("limit_reached"), new=True)
            self._show_upgrade()
            self._set_state("idle")
            self.send_btn.configure(state="normal")
            return

        dt = gen_data.get("tipo", "word").lower()
        if dt not in CONTROLLERS: dt = "word"
        titulo = gen_data.get("titulo", "Documento")
        labs = {"word": "Word", "excel": "Excel", "powerpoint": "PowerPoint"}

        self._set_state("generating")
        msg = (f"Generando {labs[dt]}: «{titulo}»..." if self.lang[0]=="es"
               else f"Generating {labs[dt]}: «{titulo}»...")
        self._bubble("assistant", msg, new=True)

        threading.Thread(
            target=self._agent_worker,
            args=(gen_data, dt, titulo),
            daemon=True,
        ).start()

    def _agent_worker(self, gen_data, dt, titulo):
        """
        BACKGROUND thread. Hace trabajo pesado (generación, apertura,
        screenshot, análisis vision). UI updates SIEMPRE vía self.after.
        """
        gfn, ofn, _ = CONTROLLERS[dt]

        # ── Fase 1: Generar el documento ─────────────────────────────────
        try:
            if getattr(self, "_attached_excel", None) and dt == "excel":
                from office.excel_controller import improve_excel, analyze_excel
                attached = self._attached_excel
                diag = analyze_excel(attached)
                n_issues = len(diag.get("issues", []))
                n_sugg = len(diag.get("suggestions", []))
                fp, bug_report = improve_excel(
                    attached,
                    gen_data.get("instrucciones", ""),
                    gen_data,
                )
                diag_txt = ""
                if n_issues:
                    diag_txt += f"\n\n\U0001f50d **Diagnostico:** {n_issues} errores detectados"
                if n_sugg:
                    diag_txt += f" · {n_sugg} sugerencias"
                bug_report = diag_txt + "\n\n" + bug_report
                self.after(0, self._reset_attachment)
            else:
                result = gfn(gen_data)
                if isinstance(result, tuple):
                    fp, bug_report = result
                else:
                    fp, bug_report = result, ""
        except Exception as e:
            self.after(0, self._on_agent_error, e)
            return

        # ── Fase 2: Persistir metadata (SQLite es thread-safe) ───────────
        try:
            db.save_document(self.conversation_id, self.user["id"], titulo, dt, fp)
        except Exception as e:
            print(f"[WEP AI] save_document failed: {e}")

        # Sincronizar con UI: doc path actual + contador
        self.after(0, self._on_doc_generated, dt, titulo, fp, gen_data)

        # ── Fase 3: Abrir en Office ──────────────────────────────────────
        self.after(0, lambda: self._set_state("opening"))
        try:
            ofn(fp)  # subprocess, thread-safe
        except Exception as e:
            print(f"[WEP AI] open_in_office failed: {e}")
        time.sleep(3)

        # ── Fase 4: Captura de pantalla ──────────────────────────────────
        self.after(0, lambda: self._set_state("capturing"))
        try:
            ss = vision.capture_office_window(APP_NAMES.get(dt, "Microsoft Word"))
        except Exception as e:
            print(f"[WEP AI] capture_office_window failed: {e}")
            ss = None
        self.after(0, self._on_screenshot_captured, dt)

        # ── Fase 5: Análisis visual (Claude Vision API) ──────────────────
        vc = ""
        if ss:
            try:
                vc = f"\n\n_{brain.analyze_screenshot(ss, dt, gen_data.get('instrucciones',''))}_"
            except Exception as e:
                print(f"[WEP AI] analyze_screenshot failed: {e}")

        # ── Fase 6: Bubble final + transición a 'presenting' ────────────
        self.after(0, self._on_agent_done, dt, titulo, fp, vc, bug_report)

    # ── Agent main-thread callbacks ───────────────────────────────────────────
    def _reset_attachment(self):
        """MAIN thread."""
        try:
            self.attach_btn.configure(fg_color=CARD2, text="📎")
            self._attach_label.configure(text="")
        except Exception as e:
            print(f"[WEP AI] _reset_attachment widget update failed: {e}")
        self._attached_excel = None

    def _on_agent_error(self, err):
        """MAIN thread."""
        self._bubble("assistant", f"{self._t('err_gen')}: {err}", new=True)
        self._set_state("idle")
        self.send_btn.configure(state="normal")

    def _on_doc_generated(self, dt, titulo, fp, gen_data):
        """MAIN thread. Actualiza estado del doc actual y el contador."""
        self.current_doc_path = fp
        self.current_gen_data = gen_data
        self.current_doc_type = dt
        self._upd_docs_counter()

    def _on_screenshot_captured(self, dt):
        """MAIN thread."""
        self._upd_preview()
        self._badge(dt)
        self.open_btn.configure(state="normal")

    def _on_agent_done(self, dt, titulo, fp, vision_comment, bug_report):
        """MAIN thread. Bubble final + cierre del flujo de generación."""
        labs = {"word": "Word", "excel": "Excel", "powerpoint": "PowerPoint"}
        self._doc_card(dt, titulo, fp)
        fq = self._t("feedback_q")
        dr = self._t("doc_ready")
        bug_section = ""
        if bug_report:
            bug_section = f"\n\n---\n⚙️ **Reporte técnico del documento:**\n{bug_report}"
        fb = (f"✅ Tu {labs[dt]} «{titulo}» está {dr}.{vision_comment}{bug_section}\n\n{fq}"
              if self.lang[0]=="es"
              else f"✅ Your {labs[dt]} «{titulo}» is {dr}.{vision_comment}\n\n{fq}")
        self._bubble("assistant", fb, new=True)
        db.save_message(self.conversation_id, "assistant", fb)
        self.chat_history.append({"role": "assistant", "content": fb})
        self._set_state("presenting")
        self.send_btn.configure(state="normal")

    def _show_upgrade(self):
        f = ctk.CTkFrame(self.chat_scroll, fg_color="#1c1000", corner_radius=10,
                         border_width=1, border_color="#78350f")
        f.grid(sticky="ew", padx=12, pady=5)
        used = db.get_docs_used(self.user["id"])
        ctk.CTkLabel(f, text=self._t("limit_title"),
                     font=ctk.CTkFont("Arial", 12, "bold"),
                     text_color=AMBER, anchor="w").pack(padx=12, pady=(10, 4), fill="x")
        ctk.CTkLabel(f, text=self._t("limit_msg").format(used),
                     font=ctk.CTkFont("Arial", 11), text_color="#fcd34d",
                     wraplength=420, anchor="w").pack(padx=12, pady=(0, 4), fill="x")
        ctk.CTkButton(f, text=self._t("upgrade_btn"), height=34, corner_radius=8,
                      font=ctk.CTkFont("Arial", 11, "bold"),
                      fg_color=BLUE, hover_color=BLUE2,
                      command=lambda: messagebox.showinfo(
                          "WEP AI", "Próximamente: wepai.app/upgrade")).pack(
            padx=12, pady=(0, 10), fill="x")
        self.chat_scroll.update()
        self.chat_scroll._parent_canvas.yview_moveto(1.0)

    def _upd_docs_counter(self):
        if self.user.get("plan", "free") == "free" and hasattr(self, "docs_lbl"):
            used = db.get_docs_used(self.user["id"])
            col  = RED if used >= 4 else AMBER
            self.docs_lbl.configure(text=f"Docs: {used}/5", text_color=col)

    def _upd_preview(self):
        img = vision.get_preview_image()
        if img:
            pw = self.prev_panel.winfo_width() - 24
            ph = self.prev_panel.winfo_height() - 80
            if pw > 50 and ph > 50: img.thumbnail((pw, ph), Image.LANCZOS)
            ci = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
            self.photo_ref = ci
            self.prev_lbl.configure(image=ci, text="")
        else:
            self.prev_lbl.configure(text=self._t("preview_sub"), image=None)

    def _badge(self, dt):
        labs = {"word": "Word", "excel": "Excel", "powerpoint": "PowerPoint"}
        cols = {"word": BLUE,   "excel": GREEN,    "powerpoint": AMBER}
        self.doc_badge.configure(text=f"  {labs.get(dt,'')}  ", fg_color=cols.get(dt, BLUE))

    def _open_in_office(self):
        if self.current_doc_path and os.path.exists(self.current_doc_path):
            _, ofn, _ = CONTROLLERS.get(self.current_doc_type, CONTROLLERS["word"])
            ofn(self.current_doc_path)
        else:
            messagebox.showwarning("WEP AI", self._t("no_doc"))

    def _set_state(self, state):
        self.app_state = state
        l = self.lang[0]
        SM = {
            "idle":       ("● " + T(l,"ready"),      BLUE),
            "thinking":   ("● " + T(l,"thinking"),    AMBER),
            "generating": ("● " + T(l,"generating"),  "#a78bfa"),
            "opening":    ("● " + T(l,"opening"),     TEAL),
            "capturing":  ("● " + T(l,"capturing"),   "#ec4899"),
            "presenting": ("● " + T(l,"presenting"),  GREEN),
            "correcting": ("● " + T(l,"correcting"),  AMBER),
        }
        lbl, col = SM.get(state, ("● " + T(l,"ready"), BLUE))
        self.status_lbl.configure(text=lbl, text_color=col)


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
def run():
    db.init_db()
    login = LoginWindow()
    login.mainloop()
    if login.result:
        user, lang = login.result
        app = WEPAIApp(user=user, lang=lang)
        app.mainloop()

if __name__ == "__main__":
    run()
