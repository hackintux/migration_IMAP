# Migration IMAP README

Ce document présente le **README** du script **Migration IMAP** (Tkinter stylé et coloré).

---

## 🌟 Aperçu

Le **Migration IMAP** est une application en Python, offrant :

* **Gestion Outlook**

  * Sauvegarde complète du profil
  * Réparation automatique des fichiers PST
  * Importation rapide de fichiers PST
  * Vérification de la configuration réseau (DNS)
  * Extraction des emails au format texte
* **Migration IMAP → Microsoft 365**

  * Transfert de tous vos messages d’un serveur IMAP vers Exchange Online
  * Barre de progression et pourcentage en temps réel
* **Interface stylée** avec Tkinter et ttk, palette verte, menu et zone de logs
* **Icône personnalisée** (`cil.ico`)

---

## 🚀 Installation

1. **Cloner ou télécharger** ce dépôt dans un dossier local.
2. **Placer** votre **icône** `cil.ico` à la racine du dossier (au même niveau que le script `outlook_m365_toolkit_tk_styled.py`).
3. **Installer** les dépendances :

   ```bash
   pip install pywin32 exchangelib
   ```
4. **Exécuter** le script :

   ```bash
   python outlook_m365_toolkit_tk_styled.py
   ```

---

## 🖱️ Utilisation de l’interface

### Onglet **Gestion Outlook**

| Bouton                 | Description                                                                                               |
| ---------------------- | --------------------------------------------------------------------------------------------------------- |
| **Sauvegarder Profil** | Exporte et copie : signatures, clés de registre et fichiers `.pst`/`.ost` vers `outlook_backup/`.         |
| **Réparer PST**        | Lance `scanpst.exe` sur tous les `.pst` sauvegardés pour corriger les fichiers corrompus.                 |
| **Importer PST**       | Ouvre une boîte de dialogue pour sélectionner un `.pst` et l’ajoute dans votre profil Outlook actif.      |
| **Vérifier Config**    | Teste la résolution DNS de `outlook.office365.com` pour valider la connectivité Microsoft 365.            |
| **Extraire Mails**     | Extrait chaque email de la boîte de réception (ID 6) en fichier `.txt` dans `outlook_backup/ExportMails`. |

### Onglet **Migration IMAP**

* **Champs de saisie** :

  * **Serveur IMAP** : adresse du serveur (ex. `imap.votreservice.com`)
  * **Utilisateur IMAP** : identifiant de connexion IMAP
  * **Mot de passe IMAP** : mot de passe associé
  * **Email M365** : adresse de destination Exchange Online
  * **Mot de passe M365** : mot de passe du compte Microsoft 365
* **Migrer IMAP→M365** :

  * Lance la migration, avec :

    * **Barre de progression** indiquant l’avancement
    * **Pourcentage** mis à jour dynamiquement

---

## ❓ Qu’est-ce que le **Folder ID 6** ?

Dans l’API **MAPI** (utilisée par `win32com.client.Dispatch("Outlook.Application")`), chaque dossier par défaut d’Outlook possède un identifiant numérique (Folder ID). Les IDs les plus courants sont :

* **3** : Éléments envoyés
* **4** : Éléments supprimés
* **5** : Boîte d’entrée (Inbox) pour certaines versions/localisations
* **6** : **Boîte de réception (Inbox)** dans la plupart des configurations
* **9** : Éléments de calendrier
* **10** : Contacts
* etc.

**Folder ID 6** correspond donc à **la Boîte de réception (Inbox)**. C’est pourquoi, pour extraire ou lire vos messages entrants, le script utilise :

```python
inbox = outlook.GetDefaultFolder(6)
```

Ainsi, tous les messages de votre dossier principal "Inbox" sont traités.

---

> **Note** : selon la version/langue d’Outlook, l’ID du dossier Inbox peut varier (parfois `5`). Si vous ne récupérez pas vos emails, essayez d’utiliser l’ID `5`.

---

## 📂 Structure du projet

```
outlook_backup/           # Dossier de sortie pour sauvegarde et exports
col.ico                   # Icône personnalisée (format ICO)
outlook_m365_toolkit_tk_styled.py  # Script principal
README.md                 # Ce fichier
```

---

## 🔧 Personnalisation

* **Changer les couleurs** : modifier les constantes `COL_*` en haut du script.
* **Icône** : remplacer `col.ico` par une autre icône (dimension recommandée : 32×32).
* **Thème ttk** : changer `style.theme_use('clam')` pour un autre thème existant (`vista`, `alt`, ...).

Bonne utilisation et n’hésitez pas à proposer des améliorations ! 😊
