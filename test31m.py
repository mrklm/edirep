#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
                Edirep — Éditeur de répertoire téléphonique
- Tkinter UI : import VCF, édition modale, prévisualisation, exports TXT/ODT/ODS.
         Tri manuel des doublons, export livret PDF avec imposition.
- mode sombre, logo normal + logo inversé (crop configurable séparé pour clair/sombre).
                            klm novembre 2025
"""

import os
import sys
import re
import quopri
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from collections import defaultdict
from datetime import datetime

# Optional libraries
try:
    from odf.opendocument import OpenDocumentText, OpenDocumentSpreadsheet
    from odf.text import P
    from odf.table import Table, TableRow, TableCell
    ODF_AVAILABLE = True
except Exception:
    ODF_AVAILABLE = False

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

try:
    from PIL import Image, ImageTk, ImageOps
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False

# ------------------------- CONFIGURATION EDITABLE -------------------------
MIN_SPACES = 3

APP_WINDOW_TITLE = "Edirep"
MAIN_HEADER_TEXT = "Éditeur de répertoire téléphonique"
STATUS_DEFAULT_TEXT = "KLM - Edirep - v3.6.0"

BUTTON_LABELS = {
    'import_vcf': "Importer VCF",
    'export_txt': "Exporter TXT",
    'export_odt': "Exporter ODT",
    'export_ods': "Exporter ODS",
    'dedupe': "Tri doublons",
    'livret': "Édition Livret"
}

PDF_DEFAULTS = {
    'title_line1': "Répertoire téléphonique",
    'title_line2': "Insérez votre nom ici",
    'count_text': "{} contacts",
    'date_text': "Édité le {}",
    'cover_line1': '',
    'cover_line2': '',
    'back_line1': 'Édité avec Edirep v.3.6.0',
    'back_line2': 'KLM Software',
}

COVER_TITLES = {
    'cover_line1': PDF_DEFAULTS['cover_line1'],
    'cover_line2': PDF_DEFAULTS['cover_line2'],
    'back_line1':  PDF_DEFAULTS['back_line1'],
    'back_line2':  PDF_DEFAULTS['back_line2'],
}
FIXED_LETTER_SIZE = 23
FIXED_CONTACT_SIZE = 13

BTN_BG = '#116cc3'
BTN_FG = 'white'
BTN_ACTIVE_BG = '#0b5291'
BTN_ACTIVE_FG = 'white'
TTK_STYLE_NAME = "Blue.TButton"

LOGO_CROP_LIGHT = (0.05, 0.05, 0.05, 0.05)
LOGO_CROP_DARK  = (0.18, 0.18, 0.18, 0.18)
LOGO_MAX_UI_WIDTH = 75
LOGO_CROP_PDF = (0.15, 0.15, 0.15, 0.15)
LOGO_MAX_PDF_WIDTH = 50

# ------------------------- UTILITAIRES -------------------------

def unfold_lines(text):
    """Déplie les lignes VCF pliées (RFC2425 folding)."""
    text = re.sub(r'\r\n[ \t]+', '', text)
    text = re.sub(r'\n[ \t]+', '', text)
    return text

def decode_quoted_printable(value):
    """Décodage quoted-printable si nécessaire."""
    try:
        return quopri.decodestring(value).decode('utf-8', errors='replace')
    except Exception:
        return value

def format_phone(number):
    """Normalise un numéro FR (simplifié)."""
    if not number:
        return ''
    s = str(number).strip()
    s = s.replace('\u00A0', ' ')
    s = re.sub(r'\s+', '', s)
    s = re.sub(r'^\+33', '0', s)
    s = re.sub(r'^0033', '0', s)
    s = re.sub(r'^006', '06', s)
    s = re.sub(r'^007', '07', s)
    digits = re.sub(r'\D', '', s)
    if len(digits) == 10:
        return '-'.join([digits[i:i+2] for i in range(0, 10, 2)])
    return s

def parse_vcf(path):
    """
    Parse un fichier VCF simple et renvoie une liste de dicts:
    [{'name': '...', 'number': '...', 'enabled': None, 'widgets': {}}...]
    """
    with open(path, 'rb') as f:
        raw_bytes = f.read()
    try:
        raw = raw_bytes.decode('utf-8')
    except Exception:
        try:
            raw = raw_bytes.decode('latin-1')
        except Exception:
            raw = raw_bytes.decode('utf-8', errors='ignore')
    raw = unfold_lines(raw)
    cards = raw.split('BEGIN:VCARD')
    contacts = []
    for card in cards:
        if not card.strip():
            continue
        fn_match = re.search(r'(?mi)^\s*FN(?:;[^:]*)?:(.+)$', card)
        name = ''
        if fn_match:
            fn_value = fn_match.group(1).strip()
            header = fn_match.group(0)
            if re.search(r'ENCODING=QUOTED-PRINTABLE', header, flags=re.I):
                fn_value = decode_quoted_printable(fn_value)
            name = fn_value
        else:
            n_match = re.search(r'(?mi)^\s*N:([^\r\n]+)', card)
            if n_match:
                parts = n_match.group(1).split(';')
                family = parts[0].strip() if len(parts) > 0 else ''
                given = parts[1].strip() if len(parts) > 1 else ''
                name = (given + ' ' + family).strip()
        name = name.replace('\\,', ',').replace(';', ' ').strip()
        tels = re.findall(r'(?mi)^\s*TEL[^:]*:([^\r\n]+)', card)
        if not tels:
            continue
        chosen = None
        for t in tels:
            d = re.sub(r'\D', '', t)
            if d.startswith('06') or d.startswith('07'):
                chosen = t
                break
            if re.match(r'^\+33\s*(6|7)|^0033\s*(6|7)', t):
                chosen = t
                break
        if not chosen:
            chosen = tels[0]
        number = format_phone(chosen)
        if not name:
            name = number
        contacts.append({'name': name, 'number': number, 'enabled': None, 'widgets': {}})
    return contacts

def get_letter(name):
    if not name:
        return '#'
    ch = name[0].upper()
    return ch if ch.isalpha() else '#'

def make_colored_button(parent, text, command, bg=BTN_BG, fg=BTN_FG, active_bg=BTN_ACTIVE_BG):
    lbl = tk.Label(parent, text=text, bg=bg, fg=fg, bd=1, relief='raised', padx=8, pady=4)
    lbl.bind("<Button-1>", lambda e: command())
    lbl.bind("<Enter>", lambda e: lbl.config(bg=active_bg))
    lbl.bind("<Leave>", lambda e: lbl.config(bg=bg))
    def on_key(e):
        if e.keysym in ('Return', 'space'):
            command()
    lbl.bind("<Key>", on_key)
    lbl.configure(takefocus=1)
    return lbl

# ---------------- Helpers PDF / Imposition ----------------

def get_fold_lines(fold):
    """Retourne lignes de pliure relatives (0..1)."""
    if fold == 2:
        return [(0.5, 0, 0.5, 1)]
    if fold == 4:
        return [(0.5, 0, 0.5, 1), (0, 0.5, 1, 0.5)]
    if fold == 8:
        return [(0.25, 0, 0.25, 1), (0.5, 0, 0.5, 1), (0.75, 0, 0.75, 1), (0, 0.5, 1, 0.5)]
    return []

def make_logical_half_pages(contacts_enabled, contact_pt, heading_pt, page_h_pts, top_margin_pts, bottom_margin_pts):
    """
    Constitue des "moitiés logiques" (demi-page) contenant en-têtes de lettre et lignes de contact.
    Renvoie (halves, line_height).
    """
    grouped = defaultdict(list)
    for ct in contacts_enabled:
        grouped[get_letter(ct['name'])].append(ct)
    for k in grouped:
        grouped[k].sort(key=lambda x: x['name'].lower())
    
    line_height = int(contact_pt * 1.05)
    usable = page_h_pts - top_margin_pts - bottom_margin_pts
    halves = []
    curr = []
    used = 0

    def push():
        nonlocal curr, used
        if curr:
            halves.append(curr)
        curr = []
        used = 0

    for letter in sorted(grouped.keys()):
        heading_h = int(heading_pt * 1.0)
        pre_gap = max(int(heading_pt * 0.4), 2)
        if used + pre_gap + heading_h > usable * 0.75 and used != 0:
            push()
        if used == 0 and heading_h + pre_gap > usable:
            curr.append(('H', letter))
            used += heading_h
        else:
            if used != 0:
                curr.append(('B',))
                used += pre_gap
            curr.append(('H', letter))
            used += heading_h
        for ct in grouped[letter]:
            if used + line_height > usable:
                push()
                curr.append(('H', letter))
                used += heading_h
            curr.append(('L', f"{ct['name']}|||{ct['number']}"))
            used += line_height
    if curr:
        halves.append(curr)
    return halves, line_height

def imposition_sequence(n_halves):
    """
    Renvoie une séquence d'imposition pour n_halves logiques (1-indexed).
    Sortie : list of pairs (left_half_index, right_half_index) — indices 1..n_halves (0 = blank)
    """
    P = n_halves
    rem = P % 4
    if rem != 0:
        P += (4 - rem)
    sheets = P // 4
    out = []
    for i in range(sheets):
        Lf = P - 2 * i
        Rf = 1 + 2 * i
        out.append((Lf if Lf <= n_halves else 0, Rf if Rf <= n_halves else 0))
        Lb = 2 + 2 * i
        Rb = P - 1 - 2 * i
        out.append((Lb if Lb <= n_halves else 0, Rb if Rb <= n_halves else 0))
    return out

# ================================ MAIN APP ==================================

class KLMEditor(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_WINDOW_TITLE)
        self.geometry('1000x650')

        # data
        self.contacts = []

        # tk vars
        self.dark = tk.BooleanVar(value=False)
        self.letter_font_size = tk.IntVar(value=FIXED_LETTER_SIZE)
        self.contact_font_size = tk.IntVar(value=FIXED_CONTACT_SIZE)
        self.left_contact_font = tk.IntVar(value=12)

        # logo placeholders
        try:
            base = os.path.dirname(__file__)
        except Exception:
            base = os.getcwd()
        self.logo_path = os.path.join(base, "logo.png")
        self.logo_normal = None
        self.logo_inverted = None
        self.logo_label = None

        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass
        style.configure(TTK_STYLE_NAME, background=BTN_BG, foreground=BTN_FG, padding=6)
        style.map(TTK_STYLE_NAME, background=[('active', BTN_ACTIVE_BG), ('!disabled', BTN_BG)],
                  foreground=[('active', BTN_ACTIVE_FG), ('!disabled', BTN_FG)])

        self._load_logo_images()
        self.create_interface()
        self.apply_theme()

    def _load_logo_images(self):
        if not PIL_AVAILABLE:
            return
        if not os.path.exists(self.logo_path):
            return
        try:
            img = Image.open(self.logo_path).convert("RGBA")
        except Exception:
            return
        w, h = img.size
        left = int(w * LOGO_CROP_LIGHT[0])
        top = int(h * LOGO_CROP_LIGHT[1])
        right = int(w * (1.0 - LOGO_CROP_LIGHT[2]))
        bottom = int(h * (1.0 - LOGO_CROP_LIGHT[3]))
        try:
            img_light = img.crop((left, top, right, bottom))
        except Exception:
            img_light = img
        if img_light.width > LOGO_MAX_UI_WIDTH:
            ratio = LOGO_MAX_UI_WIDTH / img_light.width
            img_light = img_light.resize((LOGO_MAX_UI_WIDTH, int(img_light.height * ratio)), Image.LANCZOS)
        try:
            self.logo_normal = ImageTk.PhotoImage(img_light)
        except Exception:
            self.logo_normal = None
        try:
            rgb = img.convert("RGB")
            inverted = ImageOps.invert(rgb)
            if img.mode == "RGBA":
                alpha = img.split()[-1]
                inverted = Image.merge("RGBA", list(inverted.split()) + [alpha])
            else:
                inverted = inverted.convert("RGBA")
            w2, h2 = inverted.size
            left = int(w2 * LOGO_CROP_DARK[0])
            top = int(h2 * LOGO_CROP_DARK[1])
            right = int(w2 * (1.0 - LOGO_CROP_DARK[2]))
            bottom = int(h2 * (1.0 - LOGO_CROP_DARK[3]))
            try:
                img_dark = inverted.crop((left, top, right, bottom))
            except Exception:
                img_dark = inverted
            if img_dark.width > LOGO_MAX_UI_WIDTH:
                ratio = LOGO_MAX_UI_WIDTH / img_dark.width
                img_dark = img_dark.resize((LOGO_MAX_UI_WIDTH, int(img_dark.height * ratio)), Image.LANCZOS)
            self.logo_inverted = ImageTk.PhotoImage(img_dark)
        except Exception:
            self.logo_inverted = None

    def create_interface(self):
        self.button_bar = tk.Frame(self, bg='#1976d2', pady=6)
        self.button_bar.pack(fill='x')
        tk.Label(self.button_bar, text=MAIN_HEADER_TEXT, font=('Helvetica', 14, 'bold'),
                 bg='#1976d2', fg='#d0d0d0').pack(pady=2)
        if self.logo_normal:
            self.logo_label = tk.Label(self.button_bar, image=self.logo_normal, bg='#1976d2')
            self.logo_label.image = self.logo_normal
        else:
            self.logo_label = tk.Label(self.button_bar, bg='#1976d2', width=20, height=2)
        self.logo_label.pack(pady=(4,6))
        btn_frame = tk.Frame(self.button_bar, bg='#1976d2')
        btn_frame.pack()
        ttk.Button(btn_frame, text=BUTTON_LABELS['import_vcf'], command=self.load_vcf, style=TTK_STYLE_NAME).pack(side='left', padx=4)
        ttk.Button(btn_frame, text=BUTTON_LABELS['dedupe'], command=self.manual_remove_duplicates, style=TTK_STYLE_NAME).pack(side='left', padx=4)
        ttk.Button(btn_frame, text=BUTTON_LABELS['export_txt'], command=self.export_txt, style=TTK_STYLE_NAME).pack(side='left', padx=4)
        ttk.Button(btn_frame, text=BUTTON_LABELS['export_odt'], command=self.export_odt, style=TTK_STYLE_NAME).pack(side='left', padx=4)
        ttk.Button(btn_frame, text=BUTTON_LABELS['export_ods'], command=self.export_ods, style=TTK_STYLE_NAME).pack(side='left', padx=4)
        ttk.Button(btn_frame, text=BUTTON_LABELS['livret'],
                   command=lambda: LivretWindow(self, self.contacts, logo_path=self.logo_path),
                   style=TTK_STYLE_NAME).pack(side='left', padx=8)
        tk.Checkbutton(self.button_bar, text='🌙', variable=self.dark, command=self.apply_theme,
                       bg='#1976d2', fg='#d0d0d0', selectcolor='#1976d2', relief='flat').pack(side='right', padx=8)

        main_pw = tk.PanedWindow(self, orient='horizontal')
        main_pw.pack(fill='both', expand=True, pady=6)
        left_frame = tk.Frame(main_pw)
        self.left_canvas = tk.Canvas(left_frame, borderwidth=0)
        self.vscroll_left = tk.Scrollbar(left_frame, orient='vertical', command=self.left_canvas.yview)
        self.left_canvas.configure(yscrollcommand=self.vscroll_left.set)
        self.vscroll_left.pack(side='right', fill='y')
        self.left_canvas.pack(side='left', fill='both', expand=True)
        self.inner_frame = tk.Frame(self.left_canvas)
        self.left_canvas.create_window((0, 0), window=self.inner_frame, anchor='nw')
        self.inner_frame.bind('<Configure>', lambda e: self.left_canvas.configure(scrollregion=self.left_canvas.bbox('all')))
        if sys.platform.startswith('linux'):
            self.left_canvas.bind_all('<Button-4>', self._on_mousewheel_left)
            self.left_canvas.bind_all('<Button-5>', self._on_mousewheel_left)
        else:
            self.left_canvas.bind_all('<MouseWheel>', self._on_mousewheel_left)
        main_pw.add(left_frame, minsize=280)

        right_frame = tk.Frame(main_pw, padx=6)
        tk.Label(right_frame, text='Prévisualisation', font=('Helvetica', 14, 'bold')).pack(pady=(4,6))
        self.preview_text = tk.Text(right_frame, wrap='none')
        self.vscroll_right = tk.Scrollbar(right_frame, orient='vertical', command=self.preview_text.yview)
        self.preview_text.configure(yscrollcommand=self.vscroll_right.set)
        self.vscroll_right.pack(side='right', fill='y')
        self.preview_text.pack(side='left', fill='both', expand=True)
        if sys.platform.startswith('linux'):
            self.preview_text.bind_all('<Button-4>', self._on_mousewheel_preview)
            self.preview_text.bind_all('<Button-5>', self._on_mousewheel_preview)
        else:
            self.preview_text.bind_all('<MouseWheel>', self._on_mousewheel_preview)
        main_pw.add(right_frame, minsize=420)

        self.status_bar = tk.Frame(self, bg='#1976d2', height=26)
        self.status_bar.pack(fill='x', side='bottom')
        self.status_label = tk.Label(self.status_bar, text='Contacts : 0', bg='#1976d2', fg='#d0d0d0')
        self.status_label.pack(side='left', padx=6)
        self.version_label = tk.Label(self.status_bar, text=STATUS_DEFAULT_TEXT, bg='#1976d2', fg='#d0d0d0')
        self.version_label.pack(side='right', padx=6)

    def _on_mousewheel_left(self, event):
        if sys.platform.startswith('linux'):
            if event.num == 4:
                self.left_canvas.yview_scroll(-1, 'units')
            elif event.num == 5:
                self.left_canvas.yview_scroll(1, 'units')
        else:
            self.left_canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

    def _on_mousewheel_preview(self, event):
        if sys.platform.startswith('linux'):
            if event.num == 4:
                self.preview_text.yview_scroll(-1, 'units')
            elif event.num == 5:
                self.preview_text.yview_scroll(1, 'units')
        else:
            self.preview_text.yview_scroll(int(-1 * (event.delta / 120)), 'units')

    def apply_theme(self):
        dark = self.dark.get()
        try:
            if hasattr(self, 'logo_label') and self.logo_label is not None:
                if dark and self.logo_inverted:
                    self.logo_label.configure(image=self.logo_inverted)
                    self.logo_label.image = self.logo_inverted
                elif self.logo_normal:
                    self.logo_label.configure(image=self.logo_normal)
                    self.logo_label.image = self.logo_normal
        except Exception:
            pass
        bg = '#181818' if dark else '#1976d2'
        fg = '#d0d0d0'
        txt_bg = '#111111' if dark else 'white'
        txt_fg = 'white' if dark else 'black'
        try:
            self.configure(bg=bg)
            self.button_bar.configure(bg=bg)
            self.status_bar.configure(bg=bg)
            self.inner_frame.configure(bg=bg)
            if self.logo_label:
                self.logo_label.configure(bg=bg)
        except Exception:
            pass
        def recolor(w):
            try:
                w.configure(bg=bg, fg=fg)
            except Exception:
                pass
            for ch in w.winfo_children():
                recolor(ch)
        recolor(self)
        try:
            self.preview_text.configure(bg=txt_bg, fg=txt_fg, insertbackground=txt_fg)
        except Exception:
            pass
        self.update_left_list_fonts()
        self.update_preview()

    def load_vcf(self):
        path = filedialog.askopenfilename(filetypes=[('Fichier VCF', '*.vcf')])
        if not path:
            return
        new = parse_vcf(path)
        if not new:
            messagebox.showinfo('Aucun contact', 'Aucun contact trouvé dans ce fichier VCF.')
            return
        existing_keys = set((c['name'].strip().lower(), c['number'].strip()) for c in self.contacts)
        added = 0
        for ct in new:
            key = (ct['name'].strip().lower(), ct['number'].strip())
            if key in existing_keys:
                continue
            ct['enabled'] = tk.BooleanVar(self, value=True)
            ct['widgets'] = {}
            self.contacts.append(ct)
            existing_keys.add(key)
            added += 1
        if added == 0:
            messagebox.showinfo('Import', 'Aucun nouveau contact (tout était déjà présent).')
        else:
            self.sort_contacts()
            self.refresh_contact_list()
            messagebox.showinfo('Import', f'{added} contact(s) ajoutés.')

    def sort_contacts(self):
        self.contacts.sort(key=lambda c: c['name'].lower())

    def manual_remove_duplicates(self):
        groups = defaultdict(list)
        for c in self.contacts:
            groups[c['number']].append(c)
        dup_groups = [g for g in groups.values() if len(g) > 1]
        if not dup_groups:
            messagebox.showinfo('Tri doublons', 'Aucun doublon détecté.')
            return
        for group in dup_groups:
            number = group[0]['number']
            win = tk.Toplevel(self)
            win.title(f"Doublons: {number}")
            tk.Label(win, text=f"Numéro {number} a plusieurs noms :").pack(pady=6)
            choice = tk.StringVar(value=group[0]['name'])
            for g in group:
                tk.Radiobutton(win, text=g['name'], variable=choice, value=g['name']).pack(anchor='w', padx=12)
            def confirm():
                sel = choice.get()
                kept = None
                for g in group:
                    if g['name'] == sel:
                        kept = g
                        break
                for g in group:
                    try:
                        self.contacts.remove(g)
                    except ValueError:
                        pass
                if kept:
                    self.contacts.append(kept)
                win.destroy()
            ttk.Button(win, text='Valider', command=confirm, style=TTK_STYLE_NAME).pack(pady=8)
            win.grab_set()
            self.wait_window(win)
        self.sort_contacts()
        self.refresh_contact_list()
        messagebox.showinfo('Tri doublons', 'Tri manuel terminé.')

    def _open_edit_modal(self, contact):
        win = tk.Toplevel(self)
        win.title("Éditer le contact")
        win.transient(self)
        win.update_idletasks()
        win.grab_set()
        tk.Label(win, text="Nom :").grid(row=0, column=0, sticky='e', padx=8, pady=6)
        name_var = tk.StringVar(value=contact.get('name', ''))
        tk.Entry(win, textvariable=name_var, width=40).grid(row=0, column=1, padx=8, pady=6)
        tk.Label(win, text="Téléphone :").grid(row=1, column=0, sticky='e', padx=8, pady=6)
        num_var = tk.StringVar(value=contact.get('number', ''))
        tk.Entry(win, textvariable=num_var, width=25).grid(row=1, column=1, padx=8, pady=6, sticky='w')
        enabled_var = tk.BooleanVar(value=bool(contact.get('enabled') and contact['enabled'].get()))
        tk.Checkbutton(win, text="Garder ce Contact", variable=enabled_var).grid(row=2, column=1, sticky='w', padx=8, pady=6)
        def on_ok():
            new_name = name_var.get().strip()
            new_num = format_phone(num_var.get().strip())
            if not new_name:
                messagebox.showerror("Erreur", "Le nom ne peut être vide.")
                return
            contact['name'] = new_name
            contact['number'] = new_num
            if contact.get('enabled') is None:
                contact['enabled'] = tk.BooleanVar(self, value=enabled_var.get())
            else:
                contact['enabled'].set(enabled_var.get())
            self.sort_contacts()
            self.refresh_contact_list()
            self.update_preview()
            win.destroy()
        def on_cancel():
            win.destroy()
        btn_frame = tk.Frame(win)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=(4, 8))
        ttk.Button(btn_frame, text="Annuler", command=on_cancel).pack(side='left', padx=6)
        ok_btn = make_colored_button(btn_frame, "OK", on_ok)
        ok_btn.pack(side='left', padx=6)
        win.lift()
        name_entry = win.children.get(name_var._name)
        if name_entry:
            name_entry.focus_set()
        else:
            win.focus_set()

    def refresh_contact_list(self):
        for w in self.inner_frame.winfo_children():
            w.destroy()
        grouped = defaultdict(list)
        for c in self.contacts:
            grouped[get_letter(c['name'])].append(c)
        for letter in sorted(grouped.keys()):
            lbl = tk.Label(self.inner_frame, text=letter,
                           font=('Helvetica', max(14, int(self.letter_font_size.get())) , 'bold'))
            lbl.pack(anchor='w', pady=(10, 0), padx=4)
            for c in sorted(grouped[letter], key=lambda x: x['name'].lower()):
                row = tk.Frame(self.inner_frame, pady=2)
                row.pack(fill='x', padx=4)
                if not c.get('enabled'):
                    c['enabled'] = tk.BooleanVar(self, value=True)
                cb = tk.Checkbutton(row, variable=c['enabled'], command=self.update_preview,
                                    font=('Helvetica', 11, 'bold'))
                cb.pack(side='left')
                name_lbl = tk.Label(row, text=c['name'], anchor='w')
                name_lbl.pack(side='left', fill='x', expand=True, padx=(6, 4))
                name_lbl.bind("<Double-Button-1>", lambda e, ct=c: self._open_edit_modal(ct))
                num_lbl = tk.Label(row, text=c['number'])
                num_lbl.pack(side='right', padx=4)
                edit_btn = make_colored_button(row, 'Éditer', lambda ct=c: self._open_edit_modal(ct))
                edit_btn.pack(side='right', padx=4)
                c['widgets'] = {'check': cb, 'name_lbl': name_lbl, 'num_lbl': num_lbl}
        self.update_left_list_fonts()
        self.update_preview()
        self.status_label.config(text=f'Contacts : {len(self.contacts)}')

    def update_left_list_fonts(self):
        size = self.left_contact_font.get()
        for c in self.contacts:
            w = c.get('widgets') or {}
            if w.get('name_lbl'):
                try:
                    w['name_lbl'].configure(font=('Helvetica', size))
                except Exception:
                    pass
            if w.get('num_lbl'):
                try:
                    w['num_lbl'].configure(font=('Helvetica', size))
                except Exception:
                    pass
            if w.get('check'):
                try:
                    w['check'].configure(font=('Helvetica', max(9, int(size * 0.9))))
                except Exception:
                    pass

    def update_preview(self):
        self.preview_text.delete('1.0', 'end')
        grouped = defaultdict(list)
        for c in self.contacts:
            if c.get('enabled') and c['enabled'].get():
                grouped[get_letter(c['name'])].append(c)
        overall_max = 0
        for letter in grouped:
            for c in grouped[letter]:
                overall_max = max(overall_max, len(c['name']))
        max_name = min(overall_max, 30) if overall_max > 0 else 20
        for letter in sorted(grouped.keys()):
            self.preview_text.insert('end', f"{letter}\n", 'letter')
            for c in sorted(grouped[letter], key=lambda x: x['name'].lower()):
                name = c['name']
                display_name = name if len(name) <= max_name else name[:max_name - 3] + '...'
                pad = (max_name - len(display_name)) + MIN_SPACES
                line = f"{display_name}{' ' * pad}{c['number']}\n"
                self.preview_text.insert('end', line, 'contact')
            self.preview_text.insert('end', '\n')
        lsize = FIXED_LETTER_SIZE
        csize = FIXED_CONTACT_SIZE
        try:
            self.preview_text.tag_configure('letter', font=('Helvetica', lsize, 'bold'), spacing3=8)
            self.preview_text.tag_configure('contact', font=('Courier', csize))
        except Exception:
            pass

    def export_txt(self):
        if not self.contacts:
            messagebox.showinfo('TXT Export', 'Aucun contact.')
            return
        path = filedialog.asksaveasfilename(defaultextension='.txt', filetypes=[('TXT', '*.txt')])
        if not path:
            return
        grouped = defaultdict(list)
        for c in sorted(self.contacts, key=lambda x: x['name'].lower()):
            if c.get('enabled') and c['enabled'].get():
                grouped[get_letter(c['name'])].append(c)
        overall_max = 0
        for letter in grouped:
            for c in grouped[letter]:
                overall_max = max(overall_max, len(c['name']))
        max_name = min(overall_max, 30) if overall_max > 0 else 20
        with open(path, 'w', encoding='utf-8') as f:
            for letter in sorted(grouped.keys()):
                f.write(f"{letter}\n\n")
                for c in grouped[letter]:
                    name = c['name']
                    display_name = name if len(name) <= max_name else name[:max_name - 3] + '...'
                    pad = (max_name - len(display_name)) + MIN_SPACES
                    line = f"{display_name}{' ' * pad}{c['number']}\n"
                    f.write(line)
                f.write("\n")
        messagebox.showinfo('TXT Export', f'Fichier TXT généré : {path}')

    def export_odt(self):
        if not ODF_AVAILABLE:
            messagebox.showinfo('odfpy manquant', "Installez odfpy (pip install odfpy) pour exporter ODT.")
            return
        path = filedialog.asksaveasfilename(defaultextension='.odt', filetypes=[('ODT', '*.odt')])
        if not path:
            return
        doc = OpenDocumentText()
        doc.text.addElement(P(text='Répertoire téléphonique'))
        grouped = defaultdict(list)
        for c in sorted(self.contacts, key=lambda x: x['name'].lower()):
            if c.get('enabled') and c['enabled'].get():
                grouped[get_letter(c['name'])].append(c)
        for letter in sorted(grouped.keys()):
            doc.text.addElement(P(text=letter))
            for c in grouped[letter]:
                doc.text.addElement(P(text=f"{c['name']} — {c['number']}"))
            doc.text.addElement(P(text=""))
        doc.save(path)
        messagebox.showinfo('ODT Export', f'Fichier ODT généré : {path}')

    def export_ods(self):
        if not ODF_AVAILABLE:
            messagebox.showinfo('odfpy manquant', "Installez odfpy (pip install odfpy) pour exporter ODS.")
            return
        path = filedialog.asksaveasfilename(defaultextension='.ods', filetypes=[('ODS', '*.ods')])
        if not path:
            return
        doc = OpenDocumentSpreadsheet()
        table = Table(name="Contacts")
        doc.spreadsheet.addElement(table)
        header = TableRow()
        for title in ("Nom", "Numéro"):
            cell = TableCell()
            cell.addElement(P(text=title))
            header.addElement(cell)
        table.addElement(header)
        grouped = defaultdict(list)
        for c in sorted(self.contacts, key=lambda x: x['name'].lower()):
            if c.get('enabled') and c['enabled'].get():
                grouped[get_letter(c['name'])].append(c)
        for letter in sorted(grouped.keys()):
            row = TableRow()
            cell = TableCell(numbercolumnsSpanned="2")
            cell.addElement(P(text=f"— {letter} —"))
            row.addElement(cell)
            table.addElement(row)
            for c in grouped[letter]:
                row = TableRow()
                cell_name = TableCell()
                cell_name.addElement(P(text=c['name']))
                row.addElement(cell_name)
                cell_number = TableCell()
                cell_number.addElement(P(text=c['number']))
                row.addElement(cell_number)
                table.addElement(row)
        doc.save(path)
        messagebox.showinfo('ODS Export', f'Fichier ODS généré : {path}')

# ========================= LIVRETWINDOW =========================

class LivretWindow(tk.Toplevel):
    def __init__(self, master, contacts, logo_path=None):
        super().__init__(master)
        self.title("Édition Livret PDF")
        self.transient(master)
        self.grab_set()
        self.contacts = contacts
        self.logo_path = logo_path
        
        tk.Label(self, text='Titre (ligne 1) :').pack(anchor='w', padx=8, pady=(10,2))
        self.title_var = tk.StringVar(value=PDF_DEFAULTS['title_line1'])
        tk.Entry(self, textvariable=self.title_var, width=72).pack(padx=8)
        
        tk.Label(self, text='Ligne 2 (nom) :').pack(anchor='w', padx=8, pady=(8,2))
        self.name_var = tk.StringVar(value=PDF_DEFAULTS['title_line2'])
        tk.Entry(self, textvariable=self.name_var, width=72).pack(padx=8)
        
        tk.Label(self, text='Ligne 3 (nombre contacts) :').pack(anchor='w', padx=8, pady=(8,2))
        self.count_var = tk.StringVar(value=PDF_DEFAULTS['count_text'].format(self._enabled_count()))
        tk.Entry(self, textvariable=self.count_var, width=72, state='readonly', readonlybackground='#f0f0f0', fg='black').pack(padx=8)
        
        tk.Label(self, text='Ligne 4 (date) :').pack(anchor='w', padx=8, pady=(8,2))
        self.date_var = tk.StringVar(value=PDF_DEFAULTS['date_text'].format(datetime.now().strftime('%d %B %Y')))
        tk.Entry(self, textvariable=self.date_var, width=72, state='readonly', readonlybackground='#f0f0f0', fg='black').pack(padx=8)
        
        tk.Label(self, text='(PDF A4 en mode paysage)').pack(padx=8, pady=8)

        tk.Label(self, text="Type de pliage :").pack(anchor='w', padx=8, pady=(4,2))
        self.fold_var = tk.IntVar(value=2)
        fold_choices = [2,4,8]
        fold_menu = ttk.OptionMenu(self, self.fold_var, self.fold_var.get(), *fold_choices)
        fold_menu.pack(padx=8, anchor='w')

        self.canvas_width = 380
        self.canvas_height = 260
        self.illustration = tk.Canvas(self, width=self.canvas_width, height=self.canvas_height, 
                                      bg="white", highlightthickness=1, highlightbackground="#888")
        self.illustration.pack(pady=10)
        self.update_illustration()
        self.fold_var.trace_add("write", lambda *a: self.update_illustration())

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=(6,12))
        ttk.Button(btn_frame, text='Générer PDF Livret', command=self.generate_pdf, style=TTK_STYLE_NAME).pack(side='left', padx=8)
        ttk.Button(btn_frame, text='Annuler', command=self.destroy).pack(side='left', padx=8)
        
        info = "Logo trouvé" if (self.logo_path and os.path.exists(self.logo_path)) else "Aucun logo (logo.png manquant)"
        tk.Label(self, text=info, fg='gray').pack(pady=(0,8))

    def update_illustration(self, *args):
        self.illustration.delete("all")
        fold = self.fold_var.get()
        lines = get_fold_lines(fold)
        w = self.canvas_width
        h = self.canvas_height
        m = 6
        self.illustration.create_rectangle(m, m, w-m, h-m, outline='black', width=2)
        for x0, y0, x1, y1 in lines:
            self.illustration.create_line(x0 * w, y0 * h, x1 * w, y1 * h, dash=(4,3), fill='blue', width=1.5)

    def _enabled_count(self):
        return sum(1 for c in self.contacts if c.get('enabled') and c['enabled'].get())

    def generate_pdf(self):
        fold_type = self.fold_var.get()
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror('reportlab manquant', "Installe reportlab (pip install reportlab) pour générer le PDF.")
            return
        path = filedialog.asksaveasfilename(defaultextension='.pdf', filetypes=[('PDF', '*.pdf')])
        if not path:
            return
        contacts_enabled = [c for c in self.contacts if c.get('enabled') and c['enabled'].get()]
        if not contacts_enabled:
            messagebox.showinfo('Aucun contact', 'Aucun contact sélectionné.')
            return

        c = canvas.Canvas(path, pagesize=landscape(A4))
        pw, ph = landscape(A4)
        
        if fold_type == 2:
            self._generate_fold_2(c, pw, ph, contacts_enabled)
        elif fold_type == 4:
            self._generate_fold_4(c, pw, ph, contacts_enabled)
        elif fold_type == 8:
            self._generate_fold_8(c, pw, ph, contacts_enabled)
        else:
            messagebox.showerror("Pliage inconnu", f"Type de pliage inconnu: {fold_type}")
            return

        try:
            c.save()
            messagebox.showinfo("PDF généré", f"Fichier PDF généré : {path}")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Erreur écriture PDF", str(e))

    def _generate_fold_2(self, c, pw, ph, contacts_enabled):
        """
        Pliage 2 : Simple - Couverture + demi-pages successives
        """
        demi_w = pw / 2.0
        ui_heading = FIXED_LETTER_SIZE
        ui_contact = FIXED_CONTACT_SIZE
        pdf_contact_pt = max(8, int(ui_contact * 0.9))
        pdf_heading_pt = max(12, int(ui_heading * 0.9))

        left_margin = 12 * mm
        right_margin = 12 * mm
        top_margin = 18 * mm + pdf_heading_pt * 0.4
        bottom_margin = 12 * mm

        # --- PAGE 1 : COUVERTURE (droite) + 4ème DE COUV (gauche) ---
        left_center_x = demi_w / 2.0
        right_center_x = demi_w + (demi_w / 2.0)
        center_y = ph * 0.75

        # Couverture (droite)
        c.setFont("Helvetica-Bold", 18)
        c.drawCentredString(right_center_x, center_y, self.title_var.get())
        c.setFont("Helvetica", 12)
        c.drawCentredString(right_center_x, center_y - 30*mm, self.name_var.get())
        c.drawCentredString(right_center_x, center_y - 50*mm, self.count_var.get())
        c.drawCentredString(right_center_x, center_y - 65*mm, self.date_var.get())
        c.setFont("Helvetica-Bold", 11)
        c.drawCentredString(right_center_x, ph * 0.18, COVER_TITLES.get('cover_line1', ''))
        c.drawCentredString(right_center_x, ph * 0.14, COVER_TITLES.get('cover_line2', ''))

        # 4ème de couv (gauche)
        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString(left_center_x, ph * 0.72, COVER_TITLES.get('back_line1', ''))
        try:
            if self.logo_path and os.path.exists(self.logo_path):
                logo_w = 40 * mm
                logo_h = 40 * mm
                logo_x = left_center_x - (logo_w / 2.0)
                logo_y = ph * 0.52
                c.drawImage(self.logo_path, logo_x, logo_y, width=logo_w, height=logo_h,
                           preserveAspectRatio=True, mask='auto')
        except Exception:
            pass
        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString(left_center_x, ph * 0.45, COVER_TITLES.get('back_line2', ''))
        c.showPage()

        # --- PAGES INTÉRIEURES : Remplir demi-pages successivement ---
        # Grouper les contacts par lettre
        grouped = defaultdict(list)
        for ct in contacts_enabled:
            grouped[get_letter(ct['name'])].append(ct)
        for k in grouped:
            grouped[k].sort(key=lambda x: x['name'].lower())

        # Variables pour le remplissage
        usable_height = ph - top_margin - bottom_margin
        line_height = int(pdf_contact_pt * 1.07)
        heading_height = int(pdf_heading_pt * 1.15)
        
        current_x = 0  # 0 = gauche, demi_w = droite
        current_y = ph - top_margin
        
        def move_to_next_half():
            nonlocal current_x, current_y
            if current_x == 0:
                # Passer à la demi-page droite
                current_x = demi_w
                current_y = ph - top_margin
            else:
                # Nouvelle page A4
                c.showPage()
                current_x = 0
                current_y = ph - top_margin
        
        def check_space_and_draw(needed_height):
            nonlocal current_y
            if current_y - needed_height < bottom_margin:
                move_to_next_half()
                current_y = ph - top_margin
            return True
        
        # Dessiner les contacts
        for letter in sorted(grouped.keys()):
            # Vérifier si on a assez de place pour le titre de lettre
            check_space_and_draw(heading_height)
            
            inner_left = current_x + left_margin
            inner_right = current_x + demi_w - right_margin
            inner_width = inner_right - inner_left
            
            # Dessiner la lettre
            c.setFont('Helvetica-Bold', pdf_heading_pt)
            c.drawString(inner_left, current_y, letter)
            current_y -= heading_height
            
            # Dessiner les contacts de cette lettre
            for ct in grouped[letter]:
                check_space_and_draw(line_height)
                
                name = ct['name']
                number = ct['number']
                max_chars = 30
                if len(name) > max_chars:
                    name = name[:max_chars - 3] + '...'
                
                c.setFont('Courier', pdf_contact_pt)
                c.drawString(inner_left, current_y, name)
                number_x = inner_left + int(inner_width * 0.60)
                c.drawString(number_x, current_y, number)
                current_y -= line_height
            
            # Petit espace après chaque groupe de lettre
            current_y -= int(pdf_heading_pt * 0.4)

    def _generate_fold_4(self, c, pw, ph, contacts_enabled):
        """Pliage 4 : 8 pages sur 1 feuille (couv DANS l'imposition)"""
        max_name_len = max(len(ct['name']) for ct in contacts_enabled) if contacts_enabled else 20
        
        grouped = defaultdict(list)
        for ct in contacts_enabled:
            grouped[get_letter(ct['name'])].append(ct)
        
        formatted_lines = []
        for letter in sorted(grouped.keys()):
            if formatted_lines:
                formatted_lines.append(('space', '', '', '', {}))
            formatted_lines.append(('letter', letter, '', '', {'size': 16}))
            for ct in sorted(grouped[letter], key=lambda x: x['name'].lower()):
                formatted_lines.append(('contact', '', ct['name'], ct['number'], {'size': 10}))
            formatted_lines.append(('space', '', '', '', {}))

        def draw_text_block(cobj, x, y, w, h, content, font_size=10):
            leading = font_size * 1.2
            cur_y = y + h - leading
            char_width = font_size * 0.6
            name_width = max_name_len * char_width
            number_x = x + name_width + 10
            
            for line_type, text, name, number, style in content:
                if cur_y < y + 4:
                    break
                if line_type == 'space':
                    cur_y -= leading * 0.5
                elif line_type == 'letter':
                    size = style.get('size', 16)
                    cobj.setFont("Helvetica-Bold", size)
                    cobj.drawString(x + 4, cur_y, text)
                    cur_y -= leading * 1.2
                elif line_type == 'contact':
                    size = style.get('size', font_size)
                    cobj.setFont("Helvetica", size)
                    cobj.drawString(x + 4, cur_y, name)
                    cobj.drawString(number_x, cur_y, number)
                    cur_y -= leading

        def draw_zone_rotated(cobj, x, y, w, h, content, rotation, fs=9):
            cobj.saveState()
            if rotation == 90:
                cobj.translate(x + w, y)
                cobj.rotate(90)
                draw_text_block(cobj, 4, 4, h - 8, w - 8, content, fs)
            elif rotation == -90:
                cobj.translate(x, y + h)
                cobj.rotate(-90)
                draw_text_block(cobj, 4, 4, h - 8, w - 8, content, fs)
            cobj.restoreState()

        def draw_small_cover(cobj, x, y, w, h, rotation):
            cobj.saveState()
            if rotation == 90:
                cobj.translate(x + w, y)
                cobj.rotate(90)
                cx, cy = h/2, w/2
                cobj.setFont("Helvetica-Bold", 14)
                cobj.drawCentredString(cx, cy + 15, self.title_var.get())
                cobj.setFont("Helvetica", 9)
                cobj.drawCentredString(cx, cy, self.name_var.get())
            cobj.restoreState()

        def draw_small_back(cobj, x, y, w, h, rotation):
            cobj.saveState()
            if rotation == 90:
                cobj.translate(x + w, y)
                cobj.rotate(90)
                cx, cy = h/2, w/2
                cobj.setFont("Helvetica-Bold", 10)
                cobj.drawCentredString(cx, cy, COVER_TITLES.get('back_line1', ''))
            cobj.restoreState()

        qw = pw / 2.0
        qh = ph / 2.0
        n_pages = 6
        per_page = max(1, len(formatted_lines) // n_pages)
        pages_content = []
        idx = 0
        for i in range(n_pages):
            end_idx = idx + per_page if i < n_pages - 1 else len(formatted_lines)
            pages_content.append(formatted_lines[idx:end_idx])
            idx = end_idx

        # RECTO
        if len(pages_content) >= 4:
            draw_zone_rotated(c, 0, qh, qw, qh, pages_content[3], -90, 9)
        draw_small_back(c, qw, qh, qw, qh, 90)
        if len(pages_content) >= 3:
            draw_zone_rotated(c, 0, 0, qw, qh, pages_content[2], -90, 9)
        draw_small_cover(c, qw, 0, qw, qh, 90)
        c.showPage()

        # VERSO
        if len(pages_content) >= 6:
            draw_zone_rotated(c, 0, qh, qw, qh, pages_content[5], 90, 9)
        if len(pages_content) >= 5:
            draw_zone_rotated(c, qw, qh, qw, qh, pages_content[4], 90, 9)
        if len(pages_content) >= 1:
            draw_zone_rotated(c, 0, 0, qw, qh, pages_content[0], 90, 9)
        if len(pages_content) >= 2:
            draw_zone_rotated(c, qw, 0, qw, qh, pages_content[1], 90, 9)
        c.showPage()

    def _generate_fold_8(self, c, pw, ph, contacts_enabled):
        """Pliage 8 : 16 pages sur 1 feuille (couv DANS l'imposition)"""
        max_name_len = max(len(ct['name']) for ct in contacts_enabled) if contacts_enabled else 20
        
        grouped = defaultdict(list)
        for ct in contacts_enabled:
            grouped[get_letter(ct['name'])].append(ct)
        
        formatted_lines = []
        for letter in sorted(grouped.keys()):
            if formatted_lines:
                formatted_lines.append(('space', '', '', '', {}))
            formatted_lines.append(('letter', letter, '', '', {'size': 14}))
            for ct in sorted(grouped[letter], key=lambda x: x['name'].lower()):
                formatted_lines.append(('contact', '', ct['name'], ct['number'], {'size': 7}))
            formatted_lines.append(('space', '', '', '', {}))

        def draw_text_block(cobj, x, y, w, h, content, font_size=7):
            leading = font_size * 1.2
            cur_y = y + h - leading
            char_width = font_size * 0.6
            name_width = max_name_len * char_width
            number_x = x + name_width + 6
            
            for line_type, text, name, number, style in content:
                if cur_y < y + 4:
                    break
                if line_type == 'space':
                    cur_y -= leading * 0.5
                elif line_type == 'letter':
                    size = style.get('size', 14)
                    cobj.setFont("Helvetica-Bold", size)
                    cobj.drawString(x + 2, cur_y, text)
                    cur_y -= leading * 1.2
                elif line_type == 'contact':
                    size = style.get('size', font_size)
                    cobj.setFont("Helvetica", size)
                    cobj.drawString(x + 2, cur_y, name)
                    cobj.drawString(number_x, cur_y, number)
                    cur_y -= leading

        def draw_zone_rotated(cobj, x, y, w, h, content, rotation, fs=7):
            cobj.saveState()
            if rotation == 180:
                cobj.translate(x + w, y + h)
                cobj.rotate(180)
                draw_text_block(cobj, 2, 2, w - 4, h - 4, content, fs)
            else:
                draw_text_block(cobj, x + 2, y + 2, w - 4, h - 4, content, fs)
            cobj.restoreState()

        def draw_small_cover(cobj, x, y, w, h, rotation):
            cx, cy = x + w/2, y + h/2
            cobj.setFont("Helvetica-Bold", 12)
            cobj.drawCentredString(cx, cy + 10, self.title_var.get())
            cobj.setFont("Helvetica", 8)
            cobj.drawCentredString(cx, cy - 5, self.name_var.get())

        def draw_small_back(cobj, x, y, w, h, rotation):
            cx, cy = x + w/2, y + h/2
            cobj.setFont("Helvetica-Bold", 9)
            cobj.drawCentredString(cx, cy, COVER_TITLES.get('back_line1', ''))

        zw = pw / 4.0
        zh = ph / 2.0
        n_pages = 13
        per_page = max(1, len(formatted_lines) // n_pages)
        pages_content = []
        idx = 0
        for i in range(n_pages):
            end_idx = idx + per_page if i < n_pages - 1 else len(formatted_lines)
            pages_content.append(formatted_lines[idx:end_idx])
            idx = end_idx

        # RECTO
        if len(pages_content) >= 7:
            draw_zone_rotated(c, 0*zw, zh, zw, zh, pages_content[6], 180, 7)
        if len(pages_content) >= 6:
            draw_zone_rotated(c, 1*zw, zh, zw, zh, pages_content[5], 180, 7)
        if len(pages_content) >= 3:
            draw_zone_rotated(c, 2*zw, zh, zw, zh, pages_content[2], 180, 7)
        if len(pages_content) >= 10:
            draw_zone_rotated(c, 3*zw, zh, zw, zh, pages_content[9], 180, 7)
        
        draw_small_back(c, 0*zw, 0, zw, zh, 0)
        draw_small_cover(c, 1*zw, 0, zw, zh, 0)
        if len(pages_content) >= 2:
            draw_zone_rotated(c, 2*zw, 0, zw, zh, pages_content[1], 0, 7)
        if len(pages_content) >= 11:
            draw_zone_rotated(c, 3*zw, 0, zw, zh, pages_content[10], 0, 7)
        c.showPage()

        # VERSO
        if len(pages_content) >= 9:
            draw_zone_rotated(c, 0*zw, zh, zw, zh, pages_content[8], 180, 7)
        if len(pages_content) >= 4:
            draw_zone_rotated(c, 1*zw, zh, zw, zh, pages_content[3], 180, 7)
        if len(pages_content) >= 5:
            draw_zone_rotated(c, 2*zw, zh, zw, zh, pages_content[4], 180, 7)
        if len(pages_content) >= 8:
            draw_zone_rotated(c, 3*zw, zh, zw, zh, pages_content[7], 180, 7)
        
        if len(pages_content) >= 12:
            draw_zone_rotated(c, 0*zw, 0, zw, zh, pages_content[11], 0, 7)
        if len(pages_content) >= 1:
            draw_zone_rotated(c, 1*zw, 0, zw, zh, pages_content[0], 0, 7)
        if len(pages_content) >= 13:
            draw_zone_rotated(c, 3*zw, 0, zw, zh, pages_content[12], 0, 7)
        c.showPage()

# ========================= MAIN =========================
if __name__ == '__main__':
    app = KLMEditor()
    app.mainloop()