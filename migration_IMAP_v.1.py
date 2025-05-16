import os
import shutil
import subprocess
import imaplib
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import win32com.client
import csv
from exchangelib import Credentials, Account, DELEGATE, Message, Mailbox, HTMLBody
from datetime import datetime

# ------------------------- CONFIGURATION GLOBALE -------------------------
BACKUP_DIR = os.path.join(os.getcwd(), "outlook_backup")

# ------------------------- FONCTION D'AUTHENTIFICATION -------------------------
def get_credentials():
    email = simpledialog.askstring("Auth M365", "Adresse email M365:")
    if not email:
        messagebox.showwarning("Auth annulée", "Email non renseigné.")
        return None, None
    pwd = simpledialog.askstring("Auth M365", "Mot de passe M365:", show='*')
    if not pwd:
        messagebox.showwarning("Auth annulée", "Mot de passe non renseigné.")
        return None, None
    return email, pwd

# ------------------------- FONCTION DE LOG COLORE -------------------------
def log(text, level='INFO'):
    txt_log.config(state='normal')
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    txt_log.insert('end', f"{timestamp} – {text}\n", level)
    txt_log.see('end')
    txt_log.config(state='disabled')

# ------------------------- FONCTIONS PRINCIPALES -------------------------
def sauvegarder_profil_outlook():
    email, pwd = get_credentials()
    if not email: return
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        sig_src = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Signatures')
        sig_dst = os.path.join(BACKUP_DIR, 'Signatures')
        if os.path.exists(sig_src): shutil.copytree(sig_src, sig_dst, dirs_exist_ok=True)
        reg_export = os.path.join(BACKUP_DIR, 'OutlookProfiles.reg')
        subprocess.call(f'reg export "HKCU\\Software\\Microsoft\\Office" "{reg_export}" /y', shell=True)
        for root_dir, _, files in os.walk(os.environ['USERPROFILE']):
            for f in files:
                if f.lower().endswith(('.pst', '.ost')):
                    dst = os.path.join(BACKUP_DIR, 'Mails', f)
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    shutil.copy2(os.path.join(root_dir, f), dst)
        log("Profil Outlook sauvegardé.", 'SUCCESS')
    except Exception as e:
        log(f"Erreur sauvegarde: {e}", 'ERROR')

def reparer_pst():
    email, pwd = get_credentials()
    if not email: return
    path = filedialog.askopenfilename(title="Sélectionner scanpst.exe", filetypes=[("EXE","*.exe")])
    if not path: return
    count = 0
    for root_dir, _, files in os.walk(os.path.join(BACKUP_DIR, 'Mails')):
        for f in files:
            if f.lower().endswith('.pst'):
                subprocess.run([path, os.path.join(root_dir, f)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                count += 1
    log(f"{count} PST réparé(s).", 'SUCCESS')

def importer_pst():
    email, pwd = get_credentials()
    if not email: return
    pst = filedialog.askopenfilename(title="Sélectionner PST", filetypes=[("PST","*.pst")])
    if not pst: return
    try:
        ol = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        ol.AddStore(pst)
        log(f"PST importé : {os.path.basename(pst)}", 'SUCCESS')
    except Exception as e:
        log(f"Erreur import PST: {e}", 'ERROR')

def verifier_config():
    email, pwd = get_credentials()
    if not email: return
    try:
        import socket
        socket.gethostbyname('outlook.office365.com')
        # Test connexion M365
        creds = Credentials(username=email, password=pwd)
        Account(primary_smtp_address=email, credentials=creds, autodiscover=True, access_type=DELEGATE)
        log("Configuration M365 OK", 'SUCCESS')
    except Exception as e:
        log(f"Erreur config: {e}", 'ERROR')

def extraire_mails():
    email, pwd = get_credentials()
    if not email: return
    try:
        creds = Credentials(username=email, password=pwd)
        acct = Account(primary_smtp_address=email, credentials=creds, autodiscover=True, access_type=DELEGATE)
        msgs = acct.inbox.all()
        dest = os.path.join(BACKUP_DIR, "ExportMails")
        os.makedirs(dest, exist_ok=True)
        for i, item in enumerate(msgs):
            fn = f"{i}_{(item.subject or 'Sans_sujet')[:30]}.txt"
            with open(os.path.join(dest, fn), 'w', encoding='utf-8', errors='ignore') as f:
                f.write(item.body)
        log(f"{len(msgs)} mail(s) exporté(s).", 'SUCCESS')
    except Exception as e:
        log(f"Erreur extraction mails: {e}", 'ERROR')


def extraire_contacts():
    email, pwd = get_credentials()
    if not email: return
    try:
        creds = Credentials(username=email, password=pwd)
        acct = Account(primary_smtp_address=email, credentials=creds, autodiscover=True, access_type=DELEGATE)
        contacts = acct.contacts.all()
        dest = os.path.join(BACKUP_DIR, "ExportContacts")
        os.makedirs(dest, exist_ok=True)
        filename = os.path.join(dest, "contacts.csv")
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['FullName','Email1','BusinessPhone','MobilePhone'])
            for c in contacts:
                writer.writerow([c.display_name, c.email_addresses[0].email if c.email_addresses else '', c.business_phone, c.mobile_phone])
        log(f"{len(contacts)} contact(s) exporté(s).", 'SUCCESS')
    except Exception as e:
        log(f"Erreur extraction contacts: {e}", 'ERROR')



def extraire_calendrier():
    email, pwd = get_credentials()
    if not email: return
    try:
        creds = Credentials(username=email, password=pwd)
        acct = Account(primary_smtp_address=email, credentials=creds, autodiscover=True, access_type=DELEGATE)
        items = acct.calendar.all()
        dest = os.path.join(BACKUP_DIR, "ExportCalendar")
        os.makedirs(dest, exist_ok=True)
        filename = os.path.join(dest, "calendar.csv")
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Subject','Start','End','Location','Body'])
            for ev in items:
                writer.writerow([ev.subject, ev.start.strftime('%Y-%m-%d %H:%M:%S'), ev.end.strftime('%Y-%m-%d %H:%M:%S'), ev.location, ev.body])
        log(f"{len(items)} événement(s) exporté(s).", 'SUCCESS')
    except Exception as e:
        log(f"Erreur extraction calendrier: {e}", 'ERROR')

# Migration IMAP->M365 avec progression

def migrer_imap():
    email, pwd = get_credentials()
    if not email:
        return
    srv = entries[0].get().strip()
    user = entries[1].get().strip()
    pwd_imap = entries[2].get().strip()
    mail = email
    mp = pwd
    if not all([srv, user, pwd_imap, mail, mp]):
        messagebox.showerror("Erreur", "Veuillez remplir tous les champs IMAP/M365.")
        return
    try:
        imap = imaplib.IMAP4_SSL(srv)
        imap.login(user, pwd_imap)
        imap.select("INBOX")
        creds = Credentials(username=mail, password=mp)
        acct = Account(primary_smtp_address=mail, credentials=creds, autodiscover=True, access_type=DELEGATE)
        ids = imap.search(None, 'ALL')[1][0].split()
        total = len(ids)
        progress['maximum'] = total
        progress['value'] = 0
        for i, num in enumerate(ids):
            data = imap.fetch(num, '(RFC822)')[1][0][1]
            m = Message(
                account=acct,
                subject=f"[Migré] {num.decode()}",
                body=HTMLBody(data.decode('utf-8', 'ignore')),
                to_recipients=[Mailbox(email_address=mail)]
            )
            m.send_and_save()
            progress['value'] = i + 1
            percent_label.config(text=f"{int((i+1)/total*100)} %")
            root.update_idletasks()
        log(f"{total} mail(s) migré(s).", 'SUCCESS')
    except Exception as e:
        log(f"Erreur migration: {e}", 'ERROR')

# === Création de l'interface ===
root = tk.Tk()
root.title("Migration IMAP & Gestion M365")
root.geometry("750x550")
# Icône
script_dir = os.path.dirname(os.path.abspath(__file__))
icon_path = os.path.join(script_dir, 'col.ico')
if os.path.exists(icon_path):
    try: root.iconbitmap(icon_path)
    except: pass

# Couleurs
COL_BG = '#e8f8f0'; COL_FRAME='#ffffff'
COL_TAB_BG='#4caf50'; COL_TAB_FG='#ffffff'; COL_TAB_SEL_BG='#388e3c'
COL_BTN_BG='#81c784'; COL_BTN_FG='#ffffff'; COL_BTN_ACTIVE='#66bb6a'
COL_LOG_BG='#f7f7f7'; COL_LOG_FG='#2d2d2d'

# Menu
menubar = tk.Menu(root)
file_menu = tk.Menu(menubar, tearoff=0)
file_menu.add_command(label="Quitter", command=root.quit)
menubar.add_cascade(label="Fichier", menu=file_menu)
help_menu = tk.Menu(menubar, tearoff=0)
help_menu.add_command(label="À propos", command=lambda: messagebox.showinfo("À propos", "Migration IMAP & M365\nVersion 1.1\nDéveloppé par : David SALVADOR"))
menubar.add_cascade(label="Aide", menu=help_menu)
root.config(menu=menubar, bg=COL_BG)

# Style ttk
style = ttk.Style(); style.theme_use('clam')
style.configure('TNotebook', background=COL_BG)
style.configure('TNotebook.Tab', padding=[12,8], font=('Helvetica',10,'bold'), background=COL_TAB_BG, foreground=COL_TAB_FG)
style.map('TNotebook.Tab', background=[('selected',COL_TAB_SEL_BG)])
style.configure('TButton', padding=8, font=('Helvetica',10), background=COL_BTN_BG, foreground=COL_BTN_FG)
style.map('TButton', background=[('active',COL_BTN_ACTIVE)])
style.configure('TFrame', background=COL_FRAME); style.configure('TLabel', background=COL_FRAME)

# Notebook
nb = ttk.Notebook(root)
frame1 = ttk.Frame(nb); frame2 = ttk.Frame(nb)
nb.add(frame1, text="Gestion Outlook & Export")
nb.add(frame2, text="Migration IMAP → M365")
nb.pack(fill='both', expand=True, padx=10, pady=10)

# Frame1 Buttons
buttons1 = [
    ("Sauvegarder Profil", sauvegarder_profil_outlook),
    ("Réparer PST", reparer_pst),
    ("Importer PST", importer_pst),
    ("Vérifier Config", verifier_config),
    ("Extraire Mails", extraire_mails),
    ("Extraire Contacts", extraire_contacts),
    ("Extraire Calendrier", extraire_calendrier)
]
for i,(txt,cmd) in enumerate(buttons1):
    ttk.Button(frame1, text=txt, command=cmd).grid(row=i, column=0, sticky='ew', padx=5, pady=5)

# Frame2 entries + Button + Progressbar
labels = ['Serveur IMAP:','Utilisateur IMAP:','Mot de passe IMAP:','Email M365:','Mot de passe M365:']
entries = []
for idx, lbl in enumerate(labels):
    ttk.Label(frame2, text=lbl).grid(row=idx, column=0, sticky='w', padx=5, pady=5)
    ent = ttk.Entry(frame2, show='*' if 'Mot de passe' in lbl else '')
    ent.grid(row=idx, column=1, padx=5, pady=5)
    entries.append(ent)

bttn = ttk.Button(frame2, text="Migrer IMAP→M365", command=migrer_imap)
bttn.grid(row=5, column=0, columnspan=2, sticky='ew', padx=5, pady=(5,10))

progress = ttk.Progressbar(frame2, orient='horizontal', length=400, mode='determinate')
progress.grid(row=6, column=0, columnspan=2, pady=5)

percent_label = ttk.Label(frame2, text="0 %")
percent_label.grid(row=7, column=0, columnspan=2)

# Logs
txt_log = tk.Text(root, state='disabled', height=10, bg=COL_LOG_BG, fg=COL_LOG_FG, font=('Courier',9))
txt_log.pack(fill='both', expand=False, padx=10, pady=(0,10))
# Config tags couleurs
txt_log.tag_config('SUCCESS', foreground='green')
txt_log.tag_config('WARNING', foreground='orange')
txt_log.tag_config('ERROR', foreground='red')
txt_log.tag_config('INFO', foreground='black')

root.mainloop()
