import tkinter as tk
from tkinter import ttk
import time

# Types de licence
# OS   = Open Source (code source libre)
# PROP = Proprietaire (licence commerciale)
# FREE = Gratuit (pas open source, mais gratuit)
# FREEM = Freemium (version gratuite limitee, payant pour toutes les fonctionnalites)

LICENCE_LABELS = {
    "OS":    "Open Source",
    "PROP":  "Proprietaire",
    "FREE":  "Gratuit",
    "FREEM": "Freemium",
}

LICENCE_COLORS = {
    "OS":    "#00ff88",
    "PROP":  "#ff6b6b",
    "FREE":  "#ffd700",
    "FREEM": "#ff9f43",
}

LABS = {
    "Laboratoire Arts": {
        "neon": "#00ff88", "icon": "ART",
        "sections": {
            "Tous les ordinateurs (Adobe CC)": [
                ("After Effect", "PROP"), ("Animate", "PROP"),
                ("Audition", "PROP"), ("Bridge", "FREE"),
                ("Dimension", "PROP"), ("DNG Converter", "FREE"),
                ("Dreamweaver", "PROP"), ("Illustrator", "PROP"),
                ("InDesign", "PROP"), ("Lightroom", "PROP"),
                ("Media Encoder", "PROP"), ("Photoshop", "PROP"),
                ("Prelude", "PROP"), ("Premiere", "PROP"),
            ],
            "PC Windows 11": [
                ("Antidote", "PROP"), ("Autodesk Autocad", "PROP"),
                ("Audacity", "OS"), ("Autodesk Maya", "PROP"),
                ("D5 (Rhino)", "PROP"), ("Ms Office", "PROP"),
                ("OpenShot", "OS"), ("Reaper (x64)", "PROP"),
                ("Revit", "PROP"), ("Rhino", "PROP"),
                ("Shotcut", "OS"), ("Sketchup", "FREEM"),
                ("Solidworks", "PROP"), ("Ultimaker Cura", "OS"),
            ],
            "MAC OS": [
                ("Antidote", "PROP"), ("Audacity", "OS"),
                ("DaVinci Resolve", "FREEM"), ("EndNote", "PROP"),
                ("GarageBand", "FREE"), ("Gimp", "OS"),
                ("Glyphs", "PROP"), ("Krita", "OS"),
                ("Ms Office", "PROP"), ("ShotCut", "OS"),
                ("Xmind 8", "FREEM"), ("Zotero", "OS"),
            ],
        },
    },
    "Laboratoire Communication": {
        "neon": "#00cfff", "icon": "COM",
        "sections": {
            "MAC OS X 13 (Ventura)": [
                ("Ableton Live 12", "PROP"), ("Adobe Acrobat DC", "PROP"),
                ("Adobe After Effects", "PROP"), ("Adobe Audition", "PROP"),
                ("Adobe Animate", "PROP"), ("Adobe Aero", "FREE"),
                ("Adobe Bridge", "FREE"), ("Adobe Character Animator", "PROP"),
                ("Adobe Dimension", "PROP"), ("Adobe Illustrator", "PROP"),
                ("Adobe InDesign", "PROP"), ("Adobe Incopy", "PROP"),
                ("Adobe Lightroom Classic", "PROP"),
                ("Adobe Media Encoder", "PROP"),
                ("Adobe Photoshop", "PROP"), ("Adobe Premiere Pro", "PROP"),
                ("Adobe Premiere Rush", "PROP"), ("Adobe XD", "PROP"),
                ("Antidote", "PROP"), ("Apercu", "FREE"),
                ("Arduino IDE", "OS"), ("Audacity", "OS"),
                ("Automator", "FREE"), ("Avid Media Composer", "PROP"),
                ("Balados", "FREE"), ("BBEdit", "PROP"),
                ("Blackmagic Proxy Generator Lite", "FREE"),
                ("Blackmagic RAW", "FREE"), ("Blender", "OS"),
                ("DaVinci Resolve", "FREEM"), ("Dictaphone", "FREE"),
                ("FaceTime", "FREE"), ("Firefox", "OS"),
                ("Freeform", "FREE"), ("Gimp", "OS"),
                ("GitHub Desktop", "OS"), ("GrandMA3", "PROP"),
                ("iMovie", "FREE"), ("Keynote", "FREE"),
                ("Mail", "FREE"), ("Max", "PROP"),
                ("Maxon Cinema 4D", "PROP"), ("Melodyne 5", "PROP"),
                ("Messages", "FREE"), ("Microsoft Excel", "PROP"),
                ("Microsoft PowerPoint", "PROP"),
                ("Microsoft Teams", "FREEM"), ("Microsoft Word", "PROP"),
                ("Microsoft Outlook", "PROP"),
                ("Microsoft One Note", "FREE"), ("Millumin5", "PROP"),
                ("Multidictionnaire", "PROP"), ("Musique", "FREE"),
                ("Notes", "FREE"), ("Numbers", "FREE"),
                ("OBS", "OS"), ("OneDrive", "FREEM"),
                ("Pages", "FREE"), ("Photo Booth", "FREE"),
                ("Photos", "FREE"), ("Plans", "FREE"),
                ("Pro Tools", "PROP"), ("QLab", "FREEM"),
                ("QuickTime Player", "FREE"), ("REAPER", "PROP"),
                ("Safari", "FREE"), ("Shotcut", "OS"),
                ("SketchUp", "FREEM"), ("Storyboarder", "OS"),
                ("TextEdit", "FREE"), ("TouchDesigner", "FREEM"),
                ("Transfert d'images", "FREE"), ("TwistedWave", "PROP"),
                ("VLC", "OS"), ("VueScan", "PROP"),
                ("Xcode", "FREE"),
                ("Youlean Loudness Meter", "FREEM"),
                ("zoom.us", "FREEM"), ("Zotero", "OS"),
            ],
        },
    },
    "Laboratoire Sciences": {
        "neon": "#bf5fff", "icon": "SCI",
        "sections": {
            "Bureautique": [
                ("Adobe Reader", "FREE"), ("Antidote", "PROP"),
                ("EndNote x20.4.1", "PROP"),
                ("Equitrac express client", "PROP"), ("Gimp", "OS"),
                ("LibreOffice", "OS"), ("MiKteX", "OS"),
                ("Microsoft Office 2021", "PROP"),
                ("Open office", "OS"), ("Paint.NET", "FREE"),
                ("NetSupportSchool", "PROP"), ("Texworks", "OS"),
                ("Zotero", "OS"),
            ],
            "Chimie": [
                ("BSL Analysis", "PROP"), ("Cn3D", "FREE"),
                ("Cytoscape", "OS"), ("ChemDraw Prime 22", "PROP"),
                ("Gaussian", "PROP"), ("GaussView", "PROP"),
                ("GenoPro", "PROP"), ("Molmol", "FREE"),
                ("Orbital viewer", "FREE"), ("PyMOL", "OS"),
                ("SimUtext (Simbio)", "PROP"),
            ],
            "Informatique": [
                ("Anaconda", "FREEM"), ("Android studio", "FREE"),
                ("Aptana Studio 3", "OS"),
                ("Apache NetBeans IDE", "OS"), ("BlueJ", "OS"),
                ("BlueGriffon", "OS"),
                ("Eclipse c, PHP, Java", "OS"), ("Emacs", "OS"),
                ("Edupython", "OS"), ("FileZilla", "OS"),
                ("Git", "OS"), ("Intelli IDEA Community", "OS"),
                ("IIs Express", "FREE"), ("Java", "FREE"),
                ("JEdit", "OS"), ("Jenkins", "OS"),
                ("Microsoft Visual Studio code", "FREE"),
                ("Mobaxterm", "FREEM"), ("Node.js", "OS"),
                ("Notepad++", "OS"),
                ("Oracle JDeveloper Studio", "FREE"),
                ("Oracle-OraClient10g home1", "PROP"),
                ("Pep", "OS"), ("Prolog-SWI", "OS"),
                ("PuTTY", "OS"), ("Python", "OS"),
                ("Pyscripter", "OS"), ("SQL Plus", "PROP"),
                ("SQL Developer", "FREE"), ("Spyder", "OS"),
                ("VirtualBox", "OS"), ("VNC Viewer", "FREEM"),
                ("Vi (editeur de texte)", "OS"), ("Win SCP", "OS"),
                ("X2go", "OS"), ("XMLmind", "FREEM"),
                ("7-Zip File Manager", "OS"),
            ],
            "Unix": [
                ("Alex (gem)", "OS"), ("AMC (Java seulement)", "OS"),
                ("Android-Studio", "FREE"), ("Ant", "OS"),
                ("Aptana Studio", "OS"), ("Arduino", "OS"),
                ("Atom", "OS"), ("Autotools", "OS"),
                ("Blender", "OS"), ("BlueGriffon", "OS"),
                ("BlueJ", "OS"), ("Build-essential", "OS"),
                ("Cairo", "OS"), ("Ccache", "OS"),
                ("ChimeraX", "FREEM"), ("Clang", "OS"),
                ("Clang-dev", "OS"), ("Clang-Tools", "OS"),
                ("Clang-Format", "OS"), ("Clang-Tidy", "OS"),
                ("Clion", "PROP"), ("Cmake", "OS"),
                ("CodeBlocks", "OS"), ("Cordova", "OS"),
                ("Cunit", "OS"), ("Curl", "OS"),
                ("Data display Debugger", "OS"), ("Devhelp", "OS"),
                ("Dia", "OS"), ("Docker", "FREEM"),
                ("Doctest", "OS"), ("Docutils", "OS"),
                ("Doxygen", "OS"), ("Eclipse", "OS"),
                ("Emacs", "OS"), ("FileZilla", "OS"),
                ("G++", "OS"), ("Gcc", "OS"),
                ("Gdb", "OS"), ("Gdebi", "OS"),
                ("Geany", "OS"), ("Gimp", "OS"),
                ("Git", "OS"),
            ],
            "Mathematique": [
                ("GeoGebra 6", "OS"), ("Graphic Calculus 2.0", "FREE"),
                ("Graphing Calculator", "PROP"),
                ("IBM SPSS Statistics", "PROP"), ("Inkscape", "OS"),
                ("JMP Pro", "PROP"), ("Krita", "OS"),
                ("MATLAB", "PROP"), ("R x64", "OS"),
                ("R i386", "OS"), ("Rstudio", "OS"),
                ("SAS 9.4 (francais)", "PROP"),
                ("Scilab (64 bit)", "OS"), ("Tinkerplot", "PROP"),
            ],
            "Autres Departements": [
                ("ArcGis", "PROP"), ("Gpower", "FREE"),
                ("Hot2000", "FREE"), ("HotPotatoes", "FREE"),
                ("ImageJ", "OS"), ("IoGas", "PROP"),
                ("Justinmind", "FREEM"), ("Kinovea", "OS"),
                ("Lego Mindstorms Edu EV3", "PROP"),
                ("Leapfrog", "PROP"), ("OpenShot video", "OS"),
                ("PHREEQC", "FREE"), ("ProteoWizard", "OS"),
                ("QGIS Desktop", "OS"), ("RETScreen 4", "FREE"),
                ("RETScreen Plus", "PROP"),
                ("Scratch-desktop", "OS"), ("Shotcut", "OS"),
                ("Ucinet", "PROP"), ("VLC Media Player", "OS"),
                ("Vortex10", "PROP"),
            ],
        },
    },
    "Sciences de l'education": {
        "neon": "#ff6b6b", "icon": "EDU",
        "sections": {
            "PC Windows 11": [
                ("7-Zip", "OS"), ("ActiveInspire", "PROP"),
                ("Adobe Digital Edition", "FREE"), ("Antidote", "PROP"),
                ("Algodoo", "FREEM"), ("Audacity", "OS"),
                ("Balabolka", "FREE"), ("Clavicom NG+", "FREE"),
                ("Dia", "OS"), ("EndNote", "PROP"),
                ("Foxit reader", "FREEM"),
                ("Lego Ev3 Education", "PROP"),
                ("Lego Ev3 Classrooom", "PROP"), ("Gimp", "OS"),
                ("HotPotatoes", "FREE"),
                ("IBM SPSS statistics version 30", "PROP"),
                ("Kompozer", "OS"), ("LibreOffice", "OS"),
                ("Listen N Write", "FREE"),
                ("LightBurn (imprimantes 3D)", "PROP"),
                ("Minicraft education", "PROP"),
                ("MS Office LTSC Pro 2021", "PROP"),
                ("NetLogo", "OS"), ("Netlogo 3D", "OS"),
                ("Notepad++", "OS"), ("NetSupportSchool", "PROP"),
                ("Nvivo 15", "PROP"), ("Nvu", "OS"),
                ("OBS Studio", "OS"), ("OO4Kids", "OS"),
                ("Open-Shot Video editor", "OS"),
                ("OpenStat", "OS"),
                ("Enregistreur Panopto", "PROP"),
                ("Presenter IPEVO", "FREE"), ("Paint 3D", "FREE"),
                ("Qgis", "OS"), ("Quandary", "FREE"),
                ("R", "OS"), ("RStudio", "OS"),
                ("Revo Uninstaller", "FREEM"), ("Shotcut", "OS"),
                ("Smart NoteBook (Lumino)", "PROP"),
                ("Stellarium", "OS"), ("Tinn-R", "OS"),
                ("Twine", "OS"), ("Visualizer LTSE", "FREE"),
                ("VLC", "OS"), ("WEDO Education", "PROP"),
                ("Weft QDA", "OS"), ("Xmind", "FREEM"),
                ("ZOOM", "FREEM"), ("Zotero", "OS"),
                ("Zotero extension", "OS"), ("WinScan2pd", "FREE"),
            ],
        },
    },
    "Sciences de la gestion": {
        "neon": "#ffd700", "icon": "GEST",
        "sections": {
            "PC Windows 11": [
                ("Adobe Acrobat (64-bit)", "PROP"),
                ("Adobe After Effects 2025", "PROP"),
                ("Adobe Creative Cloud", "PROP"),
                ("Adobe Illustrator 2025", "PROP"),
                ("Adobe InDesign 2025", "PROP"),
                ("Adobe Lightroom Classic", "PROP"),
                ("Adobe Media Encoder 2025", "PROP"),
                ("Adobe Photoshop 2025", "PROP"),
                ("Adobe Premiere Pro 2025", "PROP"),
                ("Anaconda3 2024.10-1", "FREEM"),
                ("Miniconda3 py313", "OS"),
                ("Antidote 12", "PROP"),
                ("Apache NetBeans version 26", "OS"),
                ("ArcGIS Pro", "PROP"), ("Arena", "PROP"),
                ("Audacity 3.7.4", "OS"),
                ("Autodesk AutoCAD 2025", "PROP"),
                ("Autodesk Revit 2024", "PROP"),
                ("Beyond 2020 Professional Browser", "PROP"),
                ("BlueJ", "OS"),
                ("CaseWare IDEA 13.0", "PROP"),
                ("EndNote 21", "PROP"), ("Google Chrome", "FREE"),
                ("IBM SPSS Statistics", "PROP"),
                ("IntelliJ IDEA 2025.1.1.1", "FREEM"),
                ("MapInfo Pro 2023 (64-bit)", "PROP"),
                ("MATLAB R2025a", "PROP"),
                ("Microsoft Edge", "FREE"),
                ("Microsoft OneDrive", "FREEM"),
                ("Microsoft PowerBI Desktop", "FREEM"),
                ("Microsoft Project", "PROP"),
                ("Microsoft Visio", "PROP"),
                ("Microsoft Visual Studio Code", "FREE"),
                ("Microsoft Office LTSC Professionnel Plus 2021", "PROP"),
                ("Mozilla Firefox", "OS"),
                ("Notepad++ (64-bit)", "OS"),
                ("NVivo 15", "PROP"),
                ("OpenOffice 4.1.15", "OS"),
                ("PyCharm 2025.1.1.1", "FREEM"),
                ("Python 3.14.2 (64-bit)", "OS"),
                ("QGIS 3.40.7 Bratislava", "OS"),
                ("R for Windows 4.5.0", "OS"),
                ("RStudio", "OS"),
                ("Safe Exam Browser (x64)", "OS"),
                ("SAP GUI for Windows 7.70", "PROP"),
                ("SketchUp 2026", "FREEM"),
                ("Stata 15", "PROP"),
                ("Tableau 2025.2", "PROP"),
                ("Tableau Prep Builder 2025.2", "PROP"),
                ("V-Ray for SketchUp", "PROP"),
                ("Zotero", "OS"), ("Shotcut", "OS"),
            ],
        },
    },
    "Sciences humaines & Langues": {
        "neon": "#ff9f43", "icon": "SHL",
        "sections": {
            "PC Windows 10": [
                ("Adobe Acrobat Reader XI Pro", "PROP"),
                ("Adobe CS3 Illustrator", "PROP"),
                ("Adobe CS3 Photoshop", "PROP"),
                ("Adobe Digital Edition v 4.5.11", "FREE"),
                ("Antidote 12", "PROP"),
                ("ArcGIS Pro v 3.5", "PROP"),
                ("Antconc v 4.2.4", "FREE"),
                ("Audacity v 3.5.1", "OS"),
                ("BlueSky Statistics v 10.3.4", "FREEM"),
                ("Catalyst v 2223", "PROP"),
                ("CloudCompare v 2.13", "OS"),
                ("CMap Tools v 6.04", "FREEM"),
                ("Dictionnaire des difficultes du francais moderne", "PROP"),
                ("Dictionnaire multidictionnaire", "PROP"),
                ("Dictionnaire Le Visuel", "PROP"),
                ("Elan v 6.8", "OS"),
                ("EndNote v 21.5", "PROP"),
                ("Express Scribe v 10.14", "FREEM"),
                ("File Maker Pro 11", "PROP"),
                ("FME Workbench v 2021.0", "PROP"),
                ("Foxit Reader v 2024.3.0.26795", "FREEM"),
                ("Gephi v 0.10", "OS"),
                ("Gimp v 2.10", "OS"),
                ("Google Earth Pro v 7.3.6", "FREE"),
                ("GPower v 3.1.9", "FREE"),
                ("GoldVarb v 3.0b3x", "FREE"),
                ("HyperBase 10", "PROP"),
                ("Inkscape v 1.1", "OS"),
                ("Iramutek v 0.7 alpha 2", "OS"),
                ("Jamovi v 2.2.5", "OS"),
                ("Jasp v 0.13.1", "OS"),
                ("Lexico v 5.13", "PROP"),
                ("Libre Office v 24.8.3.2", "OS"),
                ("Miktex v 4.4", "OS"),
                ("MS Office LTSC Pro 2021", "PROP"),
                ("MyStat v 12", "PROP"),
                ("Navigateur Firefox", "OS"),
                ("Navigateur Google Chrome", "FREE"),
                ("Navigateur Edge", "FREE"),
                ("NetSupportSchool 14", "PROP"),
                ("NotePad ++ v 8.4.4", "OS"),
                ("Nvivo 15", "PROP"),
                ("Pajek XXL et 3XL v 5.17", "FREE"),
                ("PSPP v 1.2.0", "OS"),
                ("Python", "OS"), ("Qgis v 3.28.8", "OS"),
                ("QuickTime 7", "FREE"), ("R v 4.1.1", "OS"),
                ("Revo Uninstaller v 2.4.5", "FREEM"),
                ("RStudio v 2023.06.3 build 561", "OS"),
                ("Sonal v 2.0.99", "OS"), ("SPSS 30", "PROP"),
                ("Spyder v 5.5.0 IDE", "OS"),
                ("Suite NCH v 10.14", "PROP"),
                ("Ucinet v 6.799", "PROP"),
                ("VLC v 3.0.20", "OS"),
                ("Weka v 3.8.3", "OS"),
                ("XMind 8 update 9", "FREEM"),
                ("ZOOM Workplace v 6.2.10", "FREEM"),
                ("Zotero v 6.0.30", "OS"),
                ("Zotero extension dans Firefox", "OS"),
                ("Zotero extension dans Word", "OS"),
                ("Zotero extension dans LibreOffice", "OS"),
            ],
        },
    },
}

BG     = "#0a0a0f"
PANEL  = "#0f0f1a"
PANEL2 = "#13131f"
BORDER = "#1a1a2e"
TEXT   = "#c8d6e5"
MUTED  = "#576574"
TAG_BG = "#16213e"
MATCH  = "#1f3a6e"
WHITE  = "#ffffff"


def hex_mix(hex_col, alpha, br=10, bg_=10, bb=15):
    """Melange une couleur hex avec le fond sombre."""
    c = hex_col.lstrip('#')
    r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
    return "#{:02x}{:02x}{:02x}".format(
        int(br + (r - br) * alpha),
        int(bg_ + (g - bg_) * alpha),
        int(bb + (b - bb) * alpha))


def neon_btn(parent, text, color, command, active=False):
    """Bouton style neon avec Frame+Label."""
    bg_col = hex_mix(color, 0.18) if active else PANEL2
    fg_col = color
    bd_col = color if active else hex_mix(color, 0.4)

    outer = tk.Frame(parent, bg=bd_col, padx=1, pady=1)
    inner = tk.Frame(outer, bg=bg_col)
    inner.pack(fill="both", expand=True)

    dot = tk.Label(inner, text="◆", bg=bg_col,
                   fg=color if active else MUTED,
                   font=("Courier", 7))
    dot.pack(side="left", padx=(6, 2), pady=6)

    lbl = tk.Label(inner, text=text, bg=bg_col,
                   fg=fg_col if active else MUTED,
                   font=("Courier", 8, "bold"), pady=6, padx=4)
    lbl.pack(side="left", padx=(0, 6))

    def on_enter(_):
        inner.config(bg=hex_mix(color, 0.12))
        dot.config(bg=hex_mix(color, 0.12), fg=color)
        lbl.config(bg=hex_mix(color, 0.12), fg=color)

    def on_leave(_):
        inner.config(bg=bg_col)
        dot.config(bg=bg_col, fg=color if active else MUTED)
        lbl.config(bg=bg_col, fg=fg_col if active else MUTED)

    def on_click(_):
        command()

    for w in (outer, inner, dot, lbl):
        w.bind("<Enter>", on_enter)
        w.bind("<Leave>", on_leave)
        w.bind("<Button-1>", on_click)

    return outer


class LaboSoft(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("LaboSoft Pro - Laboratoires Informatiques UQAM")
        self.configure(bg=BG)
        self.geometry("1350x820")
        self.minsize(950, 600)
        self._active_lab = None
        self._search_var = tk.StringVar()
        self._filter_licence = None  # None = tous, ou "OS", "PROP", etc.
        self._ready = False
        self._tab_widgets = {}
        self._filter_widgets = {}
        self._build_ui()
        self._ready = True
        self._show_lab(None)
        self._search_var.trace_add("write", self._on_search)
        self._tick()

    # -- Horloge ---------------------------------------------------------------
    def _tick(self):
        self._clock_lbl.config(
            text="SYS-TIME: " + time.strftime("%Y-%m-%d  //  %H:%M:%S"))
        self.after(1000, self._tick)

    # -- Construction UI -------------------------------------------------------
    def _build_ui(self):
        # Titre
        top = tk.Frame(self, bg="#07070f")
        top.pack(fill="x")
        left = tk.Frame(top, bg="#07070f")
        left.pack(side="left", padx=12, pady=8)
        tk.Label(left, text="◈", bg="#07070f", fg="#00cfff",
                 font=("Courier", 18, "bold")).pack(side="left", padx=(0, 8))
        tk.Label(left, text="LABOSOFT", bg="#07070f", fg=WHITE,
                 font=("Courier", 14, "bold")).pack(side="left")
        tk.Label(left, text="  PRO", bg="#07070f", fg="#00cfff",
                 font=("Courier", 10)).pack(side="left")

        right = tk.Frame(top, bg="#07070f")
        right.pack(side="right", padx=12, pady=8)
        self._clock_lbl = tk.Label(right, text="", bg="#07070f", fg=MUTED,
                                   font=("Courier", 8))
        self._clock_lbl.pack(side="left", padx=12)

        uqam_out = tk.Frame(right, bg="#00cfff", padx=1, pady=1)
        uqam_out.pack(side="left")
        uqam_in = tk.Frame(uqam_out, bg=PANEL2)
        uqam_in.pack()
        tk.Label(uqam_in, text="◆ UQAM", bg=PANEL2, fg="#00cfff",
                 font=("Courier", 8, "bold"), padx=10, pady=4).pack()

        tk.Frame(self, bg="#00cfff", height=1).pack(fill="x")

        # Status bar
        st = tk.Frame(self, bg=PANEL, pady=4, padx=14)
        st.pack(fill="x")
        tk.Label(st, text="◆ SYSTEME OPERATIONNEL", bg=PANEL,
                 fg="#00ff88", font=("Courier", 8)).pack(side="left")
        tk.Label(st, text="   ◆ LABORATOIRES ACTIFS", bg=PANEL,
                 fg="#00cfff", font=("Courier", 8)).pack(side="left")
        tk.Label(st, text="   ◆ PRET", bg=PANEL,
                 fg=MUTED, font=("Courier", 8)).pack(side="left")
        self._count_lbl = tk.Label(st, text="", bg=PANEL, fg=MUTED,
                                   font=("Courier", 8))
        self._count_lbl.pack(side="right")
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        # -- Onglets neon ------------------------------------------------------
        tab_outer = tk.Frame(self, bg=PANEL2, pady=10, padx=10)
        tab_outer.pack(fill="x")

        b = neon_btn(tab_outer, "TOUS LES LABOS", "#00cfff",
                     command=lambda: self._show_lab(None), active=True)
        b.pack(side="left", padx=4)
        self._tab_widgets[None] = b

        for name, lab in LABS.items():
            col = lab["neon"]
            icon = lab["icon"]
            btn = neon_btn(tab_outer, icon, col,
                           command=lambda k=name: self._show_lab(k),
                           active=False)
            btn.pack(side="left", padx=4)
            self._tab_widgets[name] = btn

        tk.Frame(self, bg="#1a1a2e", height=1).pack(fill="x")

        # -- Filtre par licence ------------------------------------------------
        lic_bar = tk.Frame(self, bg=PANEL2, pady=6, padx=12)
        lic_bar.pack(fill="x")
        tk.Label(lic_bar, text="◈ FILTRE LICENCE:", bg=PANEL2, fg=MUTED,
                 font=("Courier", 8, "bold")).pack(side="left", padx=(0, 8))

        # Bouton TOUS
        all_btn = neon_btn(lic_bar, "TOUS", "#00cfff",
                           command=lambda: self._set_licence_filter(None),
                           active=True)
        all_btn.pack(side="left", padx=3)
        self._filter_widgets[None] = all_btn

        for lic_key, lic_label in LICENCE_LABELS.items():
            col = LICENCE_COLORS[lic_key]
            fb = neon_btn(lic_bar, lic_label, col,
                          command=lambda k=lic_key: self._set_licence_filter(k),
                          active=False)
            fb.pack(side="left", padx=3)
            self._filter_widgets[lic_key] = fb

        # Legende
        legend = tk.Frame(lic_bar, bg=PANEL2)
        legend.pack(side="right")
        for lic_key in ("OS", "PROP", "FREE", "FREEM"):
            col = LICENCE_COLORS[lic_key]
            tk.Label(legend, text="●", bg=PANEL2, fg=col,
                     font=("Courier", 8)).pack(side="left", padx=(6, 1))
            tk.Label(legend, text=LICENCE_LABELS[lic_key], bg=PANEL2,
                     fg=MUTED, font=("Courier", 7)).pack(side="left")

        tk.Frame(self, bg="#1a1a2e", height=1).pack(fill="x")

        # -- Recherche ---------------------------------------------------------
        sb = tk.Frame(self, bg=PANEL2, pady=8, padx=12)
        sb.pack(fill="x")
        tk.Label(sb, text="◈", bg=PANEL2, fg="#00cfff",
                 font=("Courier", 12)).pack(side="left", padx=(0, 6))

        ef = tk.Frame(sb, bg="#0d0d1f", highlightthickness=1,
                      highlightbackground="#00cfff",
                      highlightcolor="#00cfff")
        ef.pack(side="left", fill="x", expand=True, ipady=2)
        self._entry = tk.Entry(ef, textvariable=self._search_var,
                               bg="#0d0d1f", fg="#00cfff",
                               insertbackground="#00cfff",
                               relief="flat", font=("Courier", 11))
        self._entry.pack(fill="x", expand=True, padx=8, ipady=4)
        self._entry.insert(0, "Rechercher un logiciel...")
        self._entry.config(fg=MUTED)
        self._entry.bind("<FocusIn>", self._focus_in)
        self._entry.bind("<FocusOut>", self._focus_out)
        tk.Frame(self, bg="#1a1a2e", height=1).pack(fill="x")

        # -- Canvas scroll -----------------------------------------------------
        cont = tk.Frame(self, bg=BG)
        cont.pack(fill="both", expand=True)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("N.Vertical.TScrollbar",
                        background="#1a1a2e", troughcolor=BG,
                        arrowcolor="#00cfff", bordercolor=BG,
                        darkcolor=BG, lightcolor=BG)

        self._canvas = tk.Canvas(cont, bg=BG, highlightthickness=0)
        vbar = ttk.Scrollbar(cont, orient="vertical",
                             command=self._canvas.yview,
                             style="N.Vertical.TScrollbar")
        self._canvas.configure(yscrollcommand=vbar.set)
        vbar.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._content = tk.Frame(self._canvas, bg=BG)
        self._win_id = self._canvas.create_window(
            (0, 0), window=self._content, anchor="nw")

        self._content.bind("<Configure>",
            lambda e: self._canvas.configure(
                scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
            lambda e: self._canvas.itemconfig(self._win_id, width=e.width))
        self._canvas.bind_all("<MouseWheel>",
            lambda e: self._canvas.yview_scroll(-1 * (e.delta // 120), "units"))

    # -- Mise a jour onglets ---------------------------------------------------
    def _set_tab_active(self, key):
        for k, outer in self._tab_widgets.items():
            try:
                inner = outer.winfo_children()[0]
                children = inner.winfo_children()
                dot = children[0]
                lbl = children[1]
                col = LABS[k]["neon"] if k is not None else "#00cfff"
            except Exception:
                continue

            is_active = (k == key)
            bg_col = hex_mix(col, 0.18) if is_active else PANEL2
            bd_col = col if is_active else hex_mix(col, 0.4)
            dot_fg = col if is_active else MUTED
            lbl_fg = col if is_active else MUTED

            outer.config(bg=bd_col)
            inner.config(bg=bg_col)
            dot.config(bg=bg_col, fg=dot_fg)
            lbl.config(bg=bg_col, fg=lbl_fg)

    def _set_filter_btn_active(self, key):
        for k, outer in self._filter_widgets.items():
            try:
                inner = outer.winfo_children()[0]
                children = inner.winfo_children()
                dot = children[0]
                lbl = children[1]
                col = LICENCE_COLORS.get(k, "#00cfff")
            except Exception:
                continue

            is_active = (k == key)
            bg_col = hex_mix(col, 0.18) if is_active else PANEL2
            bd_col = col if is_active else hex_mix(col, 0.4)
            dot_fg = col if is_active else MUTED
            lbl_fg = col if is_active else MUTED

            outer.config(bg=bd_col)
            inner.config(bg=bg_col)
            dot.config(bg=bg_col, fg=dot_fg)
            lbl.config(bg=bg_col, fg=lbl_fg)

    # -- Filtre licence --------------------------------------------------------
    def _set_licence_filter(self, key):
        self._filter_licence = key
        self._set_filter_btn_active(key)
        if not self._ready:
            return
        q = self._search_var.get()
        if q == "Rechercher un logiciel...":
            q = ""
        self._render(q)

    # -- Recherche -------------------------------------------------------------
    def _focus_in(self, _):
        if self._entry.get() == "Rechercher un logiciel...":
            self._entry.delete(0, "end")
            self._entry.config(fg="#00cfff")

    def _focus_out(self, _):
        if self._entry.get().strip() == "":
            self._entry.insert(0, "Rechercher un logiciel...")
            self._entry.config(fg=MUTED)

    def _on_search(self, *_):
        if not self._ready:
            return
        q = self._search_var.get()
        if q == "Rechercher un logiciel...":
            q = ""
        self._render(q)

    # -- Navigation ------------------------------------------------------------
    def _show_lab(self, key):
        self._active_lab = key
        self._set_tab_active(key)
        if not self._ready:
            return
        q = self._search_var.get()
        if q == "Rechercher un logiciel...":
            q = ""
        self._render(q)

    # -- Rendu -----------------------------------------------------------------
    def _render(self, query=""):
        for w in self._content.winfo_children():
            w.destroy()

        q = query.lower().strip()
        labs = {k: v for k, v in LABS.items()
                if self._active_lab is None or k == self._active_lab}

        cols = 2 if self._active_lab is None else 1
        wrapper = tk.Frame(self._content, bg=BG)
        wrapper.pack(fill="both", expand=True, padx=14, pady=14)

        col_frames = []
        for c in range(cols):
            f = tk.Frame(wrapper, bg=BG)
            f.pack(side="left", fill="both", expand=True,
                   padx=(0, 10 if c < cols - 1 else 0))
            col_frames.append(f)

        total = 0
        stats = {"OS": 0, "PROP": 0, "FREE": 0, "FREEM": 0}
        for i, (name, lab) in enumerate(labs.items()):
            card_total, card_stats = self._make_card(
                col_frames[i % cols], name, lab, q)
            total += card_total
            for k in stats:
                stats[k] += card_stats.get(k, 0)

        suffix = f'  pour "{query}"' if query else ""
        lic_suffix = ""
        if self._filter_licence:
            lic_suffix = f"  [{LICENCE_LABELS[self._filter_licence]}]"
        stat_str = (f"  (OS:{stats['OS']}  PROP:{stats['PROP']}"
                    f"  FREE:{stats['FREE']}  FREEM:{stats['FREEM']})")
        self._count_lbl.config(
            text=f"◆ {total} logiciel(s) affiche(s){suffix}{lic_suffix}{stat_str}")
        self._canvas.yview_moveto(0)

    def _make_card(self, parent, lab_name, lab, q):
        neon = lab["neon"]
        lic_filter = self._filter_licence

        outer = tk.Frame(parent, bg=neon, padx=1, pady=1)
        outer.pack(fill="x", pady=(0, 12))
        card = tk.Frame(outer, bg=PANEL2)
        card.pack(fill="both", expand=True)

        tk.Frame(card, bg=neon, height=2).pack(fill="x")

        hdr = tk.Frame(card, bg=PANEL2, padx=12, pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="◆", bg=PANEL2, fg=neon,
                 font=("Courier", 10, "bold")).pack(side="left", padx=(0, 6))
        tk.Label(hdr, text=lab_name.upper(), bg=PANEL2, fg=neon,
                 font=("Courier", 10, "bold")).pack(side="left")

        tk.Frame(card, bg=BORDER, height=1).pack(fill="x", padx=10)

        body = tk.Frame(card, bg=PANEL2, padx=10, pady=8)
        body.pack(fill="x")

        total = 0
        stats = {"OS": 0, "PROP": 0, "FREE": 0, "FREEM": 0}

        for sec_name, softs in lab["sections"].items():
            filtered = []
            for entry in softs:
                sw_name, sw_lic = entry
                if q and q not in sw_name.lower():
                    continue
                if lic_filter and sw_lic != lic_filter:
                    continue
                filtered.append(entry)

            if not filtered:
                continue

            total += len(filtered)
            for _, lic in filtered:
                stats[lic] = stats.get(lic, 0) + 1

            if len(lab["sections"]) > 1:
                sr = tk.Frame(body, bg=PANEL2)
                sr.pack(fill="x", pady=(6, 2))
                tk.Label(sr, text=f"  ◈ {sec_name}", bg=PANEL2, fg=neon,
                         font=("Courier", 8, "bold"),
                         anchor="w").pack(side="left")

            PER_ROW = 3
            for i in range(0, len(filtered), PER_ROW):
                row = tk.Frame(body, bg=PANEL2)
                row.pack(fill="x", pady=1)
                for sw_name, sw_lic in filtered[i:i + PER_ROW]:
                    is_match = bool(q and q in sw_name.lower())
                    lic_col = LICENCE_COLORS.get(sw_lic, TEXT)

                    bg = hex_mix(neon, 0.25) if is_match else TAG_BG
                    fg = neon if is_match else TEXT
                    brd = neon if is_match else BORDER

                    lf = tk.Frame(row, bg=brd, padx=1, pady=1)
                    lf.pack(side="left", padx=2, pady=1)
                    inner_f = tk.Frame(lf, bg=bg)
                    inner_f.pack(fill="both", expand=True)

                    # Nom du logiciel
                    tk.Label(inner_f, text=sw_name, bg=bg, fg=fg,
                             font=("Courier", 8),
                             padx=5, pady=2, width=24,
                             anchor="w").pack(side="left")

                    # Badge licence
                    lic_bg = hex_mix(lic_col, 0.2)
                    tk.Label(inner_f, text=sw_lic, bg=lic_bg, fg=lic_col,
                             font=("Courier", 6, "bold"),
                             padx=3, pady=1, width=5).pack(side="right",
                                                           padx=(0, 3),
                                                           pady=1)

        # Badge compteur
        badge_out = tk.Frame(hdr, bg=hex_mix(neon, 0.4), padx=1, pady=1)
        badge_out.pack(side="right")
        tk.Label(badge_out, text=f"  {total} logiciels  ",
                 bg=hex_mix(neon, 0.15), fg=neon,
                 font=("Courier", 8, "bold"), padx=4, pady=2).pack()

        return total, stats


if __name__ == "__main__":
    print("Demarrage LaboSoft Pro...")
    LaboSoft().mainloop()
