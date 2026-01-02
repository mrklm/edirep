# Changelog - Edirep

Toutes les modifications notables de ce projet seront documentées dans ce fichier.

Le format est basé sur [Keep a Changelog](https://keepachangelog.com/fr/1.0.0/),
et ce projet adhère au [Versioning Sémantique](https://semver.org/lang/fr/).

---

## [3.10] - 2026-01-02

### Ajouté

 - Fonction export des VCF
 - Ajout des adresses dans les repertoires (.TXT,.ODS,.ODT & .PDF)
 - Lignes vides ajoutées dans les repertoires au dessus et en dessous des adresses
   pour une meilleurs visibilité.

## [3.9] - 2026-01-02

### Corrigé

- Le titre (ou nom) s'affiche à nouveau sur les couvertures des 
  versions 4 & 8 plis.

## [3.8.1] - 2025-12-13

### Corrigé

- Le point d'inttérrogation en rouge sous Ubuntu
- La colonne édition noir sous Ubuntu
- Affichage du Logo sous Ubuntu  
- Placement logo dans les livret 4 & 8 plis

## [3.8.0] - 2025-12-13

### Ajouté

- Création d'un texte d'aide qui s'affiche dans la partie "visualisation" au demarage 
- Bouton pour afficher l'aide en mirroir du bouton "mode sombre"

## [3.7.0] - 2025-12-13

### Ajouté
- Changement du logo et modification de son emplacement

### Corrigé

- ajustement des titres 

## [3.6.2] - 2025-12-13

### Ajouté
- Numérotation des pages alternée gauche/droite dans les PDFs (pages paires à gauche, impaires à droite)
- Pas de numéro sur la couverture et la 4ème de couverture

### Corrigé
- Format des téléphones fixes français : 001/002/003/004/005 sont maintenant correctement convertis en 01/02/03/04/05

### Modifié
- Colonne de prévisualisation (droite) désormais en lecture seule (non-modifiable)

---

## [3.6.1] - 2025-12-12

### Corrigé
- Export ODS : correction de l'attribut `numbercolumnsspanned` (était en camelCase)
- Gestion des erreurs ODS plus détaillée pour faciliter le débogage

---

## [3.6.0] - 2025-12-12

### Ajouté
- **Système modulaire pour pliage 4 plis** :
  - Génération automatique de feuilles supplémentaires selon le nombre de contacts
  - Support demi-A4 (4 pages) ou A4 complet (8 pages) à insérer au milieu du livret
  - Calcul intelligent : ≤6 pages → 1 A4, 7-10 → 1 A4 + demi-A4, 11-14 → 2 A4, etc.
- Pointillés de pliage (vertical et horizontal) sur toutes les feuilles du pliage 4
- Pages blanches automatiquement placées en fin de livret après les contacts

### Corrigé
- **Imposition pliage 4 plis** complètement recalculée selon la méthode de pliage réelle :
  - Pli 1 : vertical, gauche sous droite
  - Pli 2 : horizontal, bas sous haut
  - Ordre alphabétique A→Z maintenant parfaitement respecté
- Rotations 90°/-90° correctement appliquées sur toutes les zones
- Mapping page → demi-pages (halves) corrigé pour éviter les contacts manquants
- Utilisation de `qw` au lieu de `qh` pour le calcul des demi-pages (car rotation 90°)

---

## [3.5.0] - 2025-12-11

### Ajouté
- **Pliage 2 plis fonctionnel** avec imposition correcte
- Utilisation de la fonction `make_logical_half_pages()` pour découpage intelligent du contenu
- Couverture et 4ème de couverture séparées de l'imposition pour le pliage 2

### Corrigé
- Seuil de remplissage des demi-pages augmenté de 75% à 90%
- Suppression de la répétition des titres de lettres lors du passage à une nouvelle demi-page
- Ordre alphabétique A→Z respecté dans le livret une fois plié

---

## [3.4.x] - Versions antérieures

### Fonctionnalités existantes
- Import de fichiers VCF (vCard)
- Export TXT, ODT, ODS avec ordre alphabétique
- Export PDF avec pliages 2, 4 et 8 plis
- Interface Tkinter avec prévisualisation en temps réel
- Gestion des doublons avec tri manuel
- Thème sombre/clair
- Normalisation des numéros de téléphone français
- Personnalisation des polices et tailles

---

## À venir

### Prévu
- [ ] Pliage 8 plis fonctionnel avec imposition correcte
- [ ] Amélioration du scroll sur la colonne de gauche (molette souris)
- [ ] Tests avec différents volumes de contacts (50, 500, 1000+)

### En réflexion
- [ ] Export direct vers Google Drive
- [ ] Aperçu PDF intégré dans l'interface
- [ ] Support d'autres formats d'import (CSV, Excel)

---

## Guide des types de changements

- **Ajouté** : Nouvelles fonctionnalités
- **Modifié** : Changements dans des fonctionnalités existantes
- **Déprécié** : Fonctionnalités bientôt supprimées
- **Supprimé** : Fonctionnalités retirées
- **Corrigé** : Corrections de bugs
- **Sécurité** : Corrections de vulnérabilités
