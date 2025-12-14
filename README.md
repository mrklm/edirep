# EDIREP

![Version](https://img.shields.io/badge/version-3.8.1-blue)
![License](https://img.shields.io/badge/license-GPL--3.0-green)
![Platform](https://img.shields.io/badge/platform-macOS%20%7C%20Linux%20%7C%20Windows-lightgrey)

Edirep sert à éditer votre répertoire téléphonique sous différentes formes.

On importe un fichier `.VCF` depuis un téléphone, on l'édite si besoin et on l'exporte :

1. **En fichier texte brut** → `.TXT`
2. **En fichier traitement de texte** → `.ODT`
3. **En fichier Tableur** → `.ODS`
4. **Sous différentes formes de livrets à imprimer** → `.PDF`

---

## 📥 Téléchargement

### Applications standalone (recommandé pour utilisateurs)
- **macOS** : [Edirep-macOS-v3.8.1.zip](#) _(à publier dans GitHub Releases)_
- **Windows** : _(à venir)_
- **Linux** : Voir installation depuis les sources ci-dessous

---

## 🚀 Installation

### Option 1 : Applications standalone (recommandé)
Téléchargez la version correspondante à votre système d'exploitation :
- **macOS** : [Edirep-macOS-v3.8.1.zip](#) _(à publier dans GitHub Releases)_
- **Windows** : _(à venir)_
- **Linux** : Voir installation depuis les sources ci-dessous

### Option 2 : Installation depuis les sources

1. **Clonez ce dépôt** :
```bash
   git clone https://github.com/mrklm/codes.git
   cd codes
```

2. **Créez un environnement virtuel et installez les dépendances** :
   
   **Sur macOS/Linux** :
```bash
   python3 -m venv venv
   source venv/bin/activate
   pip install reportlab python-docx odfpy
```
   
   **Sur Windows** :
```bash
   python -m venv venv
   venv\Scripts\activate
   pip install reportlab python-docx odfpy
```

3. **Lancez le programme** :
```bash
   python3 edirep.py
```

---

## 📜 Licence et conditions d'utilisation

Ce logiciel est distribué sous la **licence GNU GPL v3**, avec une clause supplémentaire interdisant tout usage commercial.

- **Vous êtes libre** de copier, modifier et redistribuer ce logiciel, sous réserve de respecter les termes de la licence.
- **Vous ne pouvez pas** utiliser ce logiciel ou ses dérivés à des fins commerciales sans autorisation écrite de l'auteur.
- Toute modification doit être redistribuée sous la même licence.

Pour plus de détails, consultez le fichier [LICENSE](LICENSE).

---

## 🛠️ Contribuer

Les contributions sont les bienvenues ! Pour contribuer :

1. Forkez ce dépôt.
2. Créez une branche pour votre fonctionnalité ou correction.
3. Ouvrez une *Pull Request* en expliquant vos modifications.

---

## ⚠️ Avertissement

Ce logiciel est fourni **sans garantie**. L'auteur décline toute responsabilité en cas de dommage ou de dysfonctionnement.

---

## 💡 Pourquoi ce projet est-il sous licence libre ?

Ce projet s'inscrit dans la philosophie du logiciel libre, promue par des associations comme [April](https://www.april.org/). Nous croyons que le partage des connaissances et des outils est essentiel pour une société numérique plus juste et transparente.

---

## 📬 Contact

Pour toute question ou demande d'autorisation (notamment pour un usage commercial), vous pouvez me contacter à l'adresse suivante :

**clementmorel@free.fr**

---

**✨ Bonne édition de répertoire avec Edirep !**