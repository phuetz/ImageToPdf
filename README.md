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
| Support Markdown | :x: | :white_check_mark: | :white_check_mark: |
| Panneau d'aperçu | :x: | :white_check_mark: | :white_check_mark: |
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

- **Images** : JPG, JPEG, PNG, BMP, GIF, TIFF
- **PDF** : Fusion de documents PDF existants (toutes les pages sont importées)
- **Markdown** : Conversion des fichiers .md en pages PDF (Standard et Pro)

### Interface

- Sélection multiple de fichiers via le dialogue de fichiers
- Glisser-déposer (drag & drop) directement dans la fenêtre
- Réorganisation des fichiers (monter/descendre dans la liste)
- Icônes de type de fichier (image, PDF, Markdown)
- Barre de progression pendant la conversion
- Panneau d'aperçu (Standard et Pro)

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

## Utilisation

1. Lancez l'application
2. Ajoutez des fichiers :
   - Cliquez sur **"Ajouter fichiers..."** pour ouvrir le dialogue de sélection
   - Ou glissez-déposez des fichiers directement dans la fenêtre
3. Organisez l'ordre des fichiers avec les boutons **"Monter"** et **"Descendre"**
4. Cliquez sur **"Créer le PDF"**
5. Choisissez l'emplacement et le nom du fichier PDF de sortie

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
├── Program.cs           # Point d'entrée de l'application
└── MainForm.cs          # Interface utilisateur et logique de conversion
```

## Dépendances

| Package | Version | Description |
|---------|---------|-------------|
| [PdfSharpCore](https://github.com/ststeiger/PdfSharpCore) | 1.3.65 | Génération et manipulation de PDF |
| [Markdig](https://github.com/xoofx/markdig) | 0.37.0 | Parser Markdown |

## Compilation cross-platform (depuis Linux/WSL)

```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:EnableWindowsTargeting=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true
```

## Licence

Ce projet est fourni tel quel, sans garantie.
