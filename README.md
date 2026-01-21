# Image to PDF Converter

Application Windows permettant de convertir une ou plusieurs images en un fichier PDF.

## Fonctionnalités

- Sélection multiple d'images via le dialogue de fichiers
- Glisser-déposer (drag & drop) des images directement dans la fenêtre
- Formats supportés : JPG, JPEG, PNG, BMP, GIF, TIFF
- Réorganisation des images (monter/descendre dans la liste)
- Suppression d'images individuelles ou de la sélection
- Barre de progression pendant la conversion
- Chaque image devient une page du PDF (taille adaptée à l'image)

## Prérequis

### Pour exécuter le binaire précompilé

- Windows 10/11 (x64)
- Aucune installation requise (application autonome)

### Pour compiler depuis les sources

- Windows 10/11
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) ou supérieur

## Installation

### Option 1 : Utiliser le binaire précompilé

1. Téléchargez le contenu du dossier `publish/`
2. Copiez tous les fichiers sur votre machine Windows
3. Double-cliquez sur `ImageToPdf.exe`

### Option 2 : Compiler depuis les sources

```bash
cd ImageToPdf
dotnet restore
dotnet build -c Release
```

L'exécutable sera dans `bin/Release/net8.0-windows/`.

### Option 3 : Créer un exécutable autonome

```bash
cd ImageToPdf
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ../publish
```

## Utilisation

1. Lancez `ImageToPdf.exe`
2. Ajoutez des images :
   - Cliquez sur **"Ajouter images..."** pour ouvrir le dialogue de sélection
   - Ou glissez-déposez des images directement dans la fenêtre
3. Organisez l'ordre des images avec les boutons **"Monter"** et **"Descendre"**
4. Cliquez sur **"Convertir en PDF"**
5. Choisissez l'emplacement et le nom du fichier PDF
6. Le PDF est généré avec une page par image

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
| [PdfSharpCore](https://github.com/ststeiger/PdfSharpCore) | 1.3.65 | Bibliothèque de génération PDF cross-platform |

## Compilation cross-platform (depuis Linux/WSL)

Pour compiler depuis Linux ou WSL pour Windows :

```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:EnableWindowsTargeting=true
```

## Licence

Ce projet est fourni tel quel, sans garantie.
