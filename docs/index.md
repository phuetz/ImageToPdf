# PDF Merger

**Fusionnez vos images, PDF et fichiers Markdown en un seul document PDF**

[![GitHub release](https://img.shields.io/github/v/release/phuetz/ImageToPdf)](https://github.com/phuetz/ImageToPdf/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://github.com/phuetz/ImageToPdf/blob/master/LICENSE)
[![.NET](https://img.shields.io/badge/.NET-8.0-purple)](https://dotnet.microsoft.com/)

---

## Pourquoi PDF Merger ?

PDF Merger est une application Windows légère et gratuite qui simplifie la création de documents PDF. Que vous ayez besoin de combiner des photos de vacances, de fusionner des factures scannées ou de compiler de la documentation, PDF Merger le fait en quelques clics.

### Caractéristiques principales

- **Simple et intuitif** : Interface style Windows Explorer familière
- **Glisser-déposer** : Ajoutez vos fichiers en les déposant directement
- **Multi-format** : Images (JPG, PNG, BMP, GIF, TIFF), PDF et Markdown
- **Portable** : Exécutable autonome, aucune installation requise
- **Open source** : Code source disponible sous licence MIT

---

## Téléchargement

<a href="https://github.com/phuetz/ImageToPdf/releases/latest" class="download-button">
  Télécharger PDF Merger
</a>

**Configuration requise** : Windows 10/11 (x64)

---

## Fonctionnalités

### Interface moderne

PDF Merger adopte le style visuel de Windows 11 avec une barre d'outils intuitive et 7 modes d'affichage différents :

- Très grandes icônes
- Grandes icônes
- Icônes moyennes
- Petites icônes
- Liste
- Détails
- Mosaïque

### Aperçu en temps réel

Visualisez vos fichiers avant de créer le PDF :

- **Images** : Aperçu avec zoom (25% à 200%)
- **PDF** : Navigation multi-pages
- **Markdown** : Rendu du contenu

### Mode ligne de commande

Automatisez vos tâches avec le mode CLI :

```bash
# Fusionner des fichiers
ImageToPdf.exe image1.jpg image2.png document.pdf resultat.pdf

# Avec wildcards
ImageToPdf.exe -o rapport.pdf C:\images\*.png
```

---

## Éditions

| Fonctionnalité | Lite | Standard | Pro |
|----------------|:----:|:--------:|:---:|
| Fusion d'images et PDF | ✓ | ✓ | ✓ |
| Glisser-déposer | ✓ | ✓ | ✓ |
| Mode ligne de commande | ✓ | ✓ | ✓ |
| Support Markdown | | ✓ | ✓ |
| Panneau d'aperçu avec zoom | | ✓ | ✓ |
| Conversion PDF → Word | | ✓ | ✓ |
| Numérotation des pages | | | ✓ |
| Filigrane personnalisable | | | ✓ |
| Choix du format de page | | | ✓ |

---

## Guide de démarrage rapide

1. **Téléchargez** la dernière version depuis GitHub
2. **Extrayez** le fichier ZIP
3. **Lancez** `ImageToPdf.exe`
4. **Ajoutez** vos fichiers (glisser-déposer ou bouton "Nouveau")
5. **Organisez** l'ordre avec les flèches ↑ ↓
6. **Créez** votre PDF en cliquant sur "Créer PDF"

### Raccourcis clavier

| Raccourci | Action |
|-----------|--------|
| `Ctrl+O` | Ajouter des fichiers |
| `Ctrl+S` | Créer le PDF |
| `Ctrl+P` | Afficher/masquer l'aperçu |
| `Suppr` | Supprimer la sélection |
| `F5` | Aperçu du résultat |

---

## Contribuer

PDF Merger est un projet open source. Les contributions sont les bienvenues !

- [Code source sur GitHub](https://github.com/phuetz/ImageToPdf)
- [Signaler un bug](https://github.com/phuetz/ImageToPdf/issues)
- [Proposer une fonctionnalité](https://github.com/phuetz/ImageToPdf/issues)

---

## Licence

PDF Merger est distribué sous licence MIT. Vous êtes libre de l'utiliser, le modifier et le redistribuer.

---

<footer>
  <p>Développé par <a href="https://github.com/phuetz">Patrice Huetz</a></p>
</footer>
