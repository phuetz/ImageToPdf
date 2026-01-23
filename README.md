# PDF Merger

Application Windows permettant de fusionner des images, fichiers PDF et documents Markdown en un seul fichier PDF.

## Éditions disponibles

PDF Merger est disponible en trois éditions :

| Fonctionnalité | Lite | Standard | Pro |
|----------------|:----:|:--------:|:---:|
| Fusion d'images (JPG, PNG, BMP, GIF, TIFF) | :white_check_mark: | :white_check_mark: | :white_check_mark: |
| Fusion de fichiers PDF | :white_check_mark: | :white_check_mark: | :white_check_mark: |
| Glisser-déposer | :white_check_mark: | :white_check_mark: | :white_check_mark: |
| Réorganisation des fichiers | :white_check_mark: | :white_check_mark: | :white_check_mark: |
| Mode ligne de commande | :white_check_mark: | :white_check_mark: | :white_check_mark: |
| Support Markdown | :x: | :white_check_mark: | :white_check_mark: |
| Panneau d'aperçu avec zoom | :x: | :white_check_mark: | :white_check_mark: |
| Conversion PDF → Word | :x: | :white_check_mark: | :white_check_mark: |
| Intégration PDFsam | :x: | :white_check_mark: | :white_check_mark: |
| Numérotation des pages | :x: | :x: | :white_check_mark: |
| Filigrane personnalisable | :x: | :x: | :white_check_mark: |
| Choix du format de page | :x: | :x: | :white_check_mark: |
| Paramètres de qualité d'image | :x: | :x: | :white_check_mark: |

### Téléchargement

- **[PDF Merger Lite](https://github.com/phuetz/ImageToPdf/releases)** - Léger et simple pour la fusion basique
- **[PDF Merger Standard](https://github.com/phuetz/ImageToPdf/releases)** - Fonctionnalités complètes pour un usage quotidien
- **[PDF Merger Pro](https://github.com/phuetz/ImageToPdf/releases)** - Toutes les options avancées

## Fonctionnalités

### Formats supportés

- **Images** : JPG, JPEG, PNG, BMP, GIF, TIFF
- **PDF** : Fusion de documents PDF existants (toutes les pages sont importées)
- **Markdown** : Conversion des fichiers .md en pages PDF (Standard et Pro)

### Interface style Windows Explorer

- Barre d'outils Windows 11 avec icônes 24x24
- **7 modes d'affichage** : Très grandes icônes, Grandes icônes, Icônes moyennes, Petites icônes, Liste, Détails, Mosaïque
- Miniatures dynamiques pour les images et types de fichiers
- Tri par colonnes (nom, type, date de modification)
- Panneau divisible avec splitter redimensionnable

### Aperçu de fichiers

- Aperçu des images en temps réel
- **Navigation multi-pages** pour les PDF (boutons précédent/suivant)
- **Zoom** de 25% à 200%
- Aperçu du contenu Markdown

### Historique et raccourcis

- **Fichiers récents** : Accès rapide aux derniers PDF créés
- **Raccourcis clavier** :
  - `Ctrl+O` : Ajouter des fichiers
  - `Ctrl+S` : Créer le PDF
  - `Ctrl+P` : Afficher/masquer l'aperçu
  - `Suppr` : Supprimer la sélection
  - `F5` : Aperçu du résultat

## Mode ligne de commande

PDF Merger peut être utilisé en ligne de commande pour l'automatisation :

```bash
# Afficher l'aide
ImageToPdf.exe --help

# Fusionner des fichiers (dernier argument = fichier de sortie)
ImageToPdf.exe image1.jpg image2.png document.pdf resultat.pdf

# Spécifier le fichier de sortie avec -o
ImageToPdf.exe -o resultat.pdf image1.jpg image2.png

# Mode verbose pour voir le détail du traitement
ImageToPdf.exe -v -o merged.pdf doc1.pdf doc2.pdf

# Utiliser des wildcards
ImageToPdf.exe -o rapport.pdf C:\images\*.png C:\docs\*.pdf
```

### Options

| Option | Description |
|--------|-------------|
| `-o, --output <fichier>` | Spécifier le fichier PDF de sortie |
| `-v, --verbose` | Afficher les détails du traitement |
| `-h, --help` | Afficher l'aide |

Sans argument, l'interface graphique est lancée.

## Prérequis

### Pour exécuter le binaire précompilé

- Windows 10/11 (x64)
- Aucune installation requise (application autonome)

### Pour compiler depuis les sources

- Windows 10/11
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) ou supérieur

## Installation

### Option 1 : Télécharger le binaire

Téléchargez la dernière version depuis la page [Releases](https://github.com/phuetz/ImageToPdf/releases).

### Option 2 : Compiler depuis les sources

```bash
cd ImageToPdf
dotnet restore
dotnet build -c Release
```

### Option 3 : Créer un exécutable autonome

```bash
cd ImageToPdf
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true
```

## Utilisation (Interface graphique)

1. Lancez l'application (double-clic ou sans arguments en ligne de commande)
2. Ajoutez des fichiers :
   - Cliquez sur **"Nouveau"** pour ouvrir le dialogue de sélection
   - Ou glissez-déposez des fichiers directement dans la fenêtre
3. Changez le mode d'affichage via le menu **"Affichage"**
4. Organisez l'ordre des fichiers avec les boutons **↑** et **↓**
5. Utilisez le panneau d'aperçu pour visualiser les fichiers
6. Cliquez sur **"Aperçu"** pour voir le résultat final
7. Cliquez sur **"Créer PDF"** pour générer le fichier

### Comportement par type de fichier

| Type | Comportement |
|------|-------------|
| Image | Une page par image, taille adaptée à l'image |
| PDF | Toutes les pages du PDF sont importées |
| Markdown | Converti en texte, rendu sur pages A4 |

## Structure du projet

```
ImageToPdf/
├── ImageToPdf.csproj    # Fichier projet .NET
├── Program.cs           # Point d'entrée et mode ligne de commande
├── MainForm.cs          # Interface utilisateur
└── PdfMerger.cs         # Logique de création PDF (partagée GUI/CLI)
```

## Dépendances

| Package | Version | Description |
|---------|---------|-------------|
| [PdfSharpCore](https://github.com/ststeiger/PdfSharpCore) | 1.3.65 | Génération et manipulation de PDF |
| [Markdig](https://github.com/xoofx/markdig) | 0.37.0 | Parser Markdown |
| [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) | 3.0.2 | Conversion PDF → Word |
| [iText7](https://github.com/itext/itext7-dotnet) | 8.0.2 | Extraction de texte PDF |

## Compilation cross-platform (depuis Linux/WSL)

```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:EnableWindowsTargeting=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true
```

## Licence

Ce projet est fourni tel quel, sans garantie.
