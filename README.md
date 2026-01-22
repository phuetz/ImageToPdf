# PDF Merger

Application Windows permettant de fusionner des images, fichiers PDF et documents Markdown en un seul fichier PDF.

## Fonctionnalités

- **Images** : JPG, JPEG, PNG, BMP, GIF, TIFF
- **PDF** : Fusion de documents PDF existants (toutes les pages sont importées)
- **Markdown** : Conversion des fichiers .md en pages PDF

### Interface

- Sélection multiple de fichiers via le dialogue de fichiers
- Glisser-déposer (drag & drop) directement dans la fenêtre
- Réorganisation des fichiers (monter/descendre dans la liste)
- Indicateur de type de fichier [IMG], [PDF], [MD]
- Barre de progression pendant la conversion

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
