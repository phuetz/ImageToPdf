# 07 — Administration : certificat, signature et publication

> **Pour qui ?** Ce guide s'adresse à l'administrateur de l'application et au **service informatique**. Si vous êtes un utilisateur de PDF Merger, vous pouvez l'ignorer — tout ce qui vous concerne est dans les guides 00 à 06.

## État des lieux (02/07/2026)

| Élément | État |
|---|---|
| Application déployée | **v2.9.3**, signée numériquement `CN=CCAS` (signature valide + horodatage) |
| Code source | `https://github.com/phuetz/ImageToPdf` (branche `master`), licence AGPL v3 |
| Certificat de signature | Auto-signé **CCAS**, valide jusqu'au **01/07/2029** |
| Clé privée | Magasin de certificats du poste de build (`Cert:\CurrentUser\My`) + fichier `signing\PDFMerger-CodeSigning.pfx` |
| Chaîne de publication GitHub (CI) | **En panne** (voir Action 2) — les releases GitHub servent d'anciens ZIP non signés |

**Il reste deux actions :** une pour le service informatique (la plus importante), une optionnelle côté GitHub.

---

## Action 1 — Installer le certificat CCAS sur les postes *(service informatique)*

**Objectif :** faire disparaître définitivement l'alerte bleue SmartScreen (« Éditeur inconnu ») au lancement de `ImageToPdf.exe` sur tous les postes.

L'exécutable est signé avec un certificat **auto-signé** : Windows ne lui fera confiance que si le certificat public est installé sur le poste. Le fichier à distribuer est :

```
signing\PDFMerger-CodeSigning-public.cer
```

Ce fichier ne contient **que la partie publique** du certificat : il peut circuler librement (mail, GPO, partage réseau).

### Par stratégie de groupe (recommandé)

Dans une GPO appliquée aux postes concernés, importer le `.cer` dans **les deux** magasins suivants (niveau ordinateur) :

1. `Configuration ordinateur → Stratégies → Paramètres Windows → Paramètres de sécurité → Stratégies de clé publique → **Autorités de certification racines de confiance**` → clic droit → *Importer* → sélectionner le `.cer`.
2. Même chemin → **Éditeurs approuvés** → *Importer* → même fichier.

### Manuellement, poste par poste (dépannage)

Dans une invite de commandes **administrateur** :

```bat
certutil -addstore -f Root PDFMerger-CodeSigning-public.cer
certutil -addstore -f TrustedPublisher PDFMerger-CodeSigning-public.cer
```

### Vérification

Sur un poste équipé : supprimer le fichier `%LocalAppData%\...` n'est pas nécessaire — il suffit de lancer `ImageToPdf.exe` : **aucune alerte ne doit apparaître**. Les propriétés du fichier → onglet *Signatures numériques* doivent montrer « CCAS » avec une signature valide.

### À noter pour plus tard

Le certificat **expire le 01/07/2029**. Prévoir de générer un nouveau certificat et de redéployer son `.cer` avant cette date (les exécutables déjà signés et horodatés resteront valides après l'expiration grâce à l'horodatage).

---

## Action 2 — Remettre la publication GitHub en état *(optionnel)*

Le workflow GitHub `Build and Sign Release` échoue depuis des mois à l'étape de signature : **« Could not authorize against SignPath API »** — le jeton d'API SignPath (secret `SIGNPATH_API_TOKEN`) ou l'identifiant d'organisation (`SIGNPATH_ORGANIZATION_ID`) est invalide ou expiré. Conséquence : la page *Releases* de GitHub propose encore d'anciens ZIP non signés.

Trois options, de la plus simple à la plus lourde :

### Option A — Signer en CI avec le certificat CCAS (recommandé)

Abandonner SignPath et signer directement dans le workflow avec le PFX :

1. Encoder le PFX en Base64 (sur le poste de build) :
   ```powershell
   [Convert]::ToBase64String([IO.File]::ReadAllBytes('signing\PDFMerger-CodeSigning.pfx')) | Set-Clipboard
   ```
2. Sur GitHub → *Settings → Secrets and variables → Actions* : créer `SIGNING_PFX` (coller la Base64) et `SIGNING_PFX_PASSWORD` (le mot de passe du PFX).
3. Remplacer le job `sign` de `.github/workflows/build-and-sign.yml` par un job `windows-latest` qui reconstitue le PFX et signe avec `Set-AuthenticodeSignature` (ou `signtool`), puis horodate (`http://timestamp.digicert.com`).

### Option B — Réparer SignPath

Se connecter sur `signpath.io`, régénérer un jeton d'API, et mettre à jour les secrets `SIGNPATH_API_TOKEN` / `SIGNPATH_ORGANIZATION_ID` du dépôt GitHub. Le reste du workflow est déjà correct (projet `ImageToPdf`, policy `release-signing`, configuration `initial`).

### Option C — Rester en publication manuelle

Continuer à signer localement (procédure ci-dessous) et publier la release à la main :

```powershell
Compress-Archive -Path .\publish\ImageToPdf.exe -DestinationPath PDFMerger-2.9.3-win-x64-signe.zip
gh release create v2.9.3 PDFMerger-2.9.3-win-x64-signe.zip --title "PDF Merger 2.9.3" --notes "Binaire signé CCAS."
```

⚠️ La création du tag `v2.9.3` déclenchera le workflow existant, qui échouera à l'étape de signature tant que l'option A ou B n'est pas faite. Sans conséquence : la release créée manuellement reste en ligne ; seul un ✗ rouge apparaîtra dans l'onglet *Actions*.

---

## Procédure — Publier une nouvelle version (recette complète)

À exécuter sur le poste qui possède le certificat (vérifiable avec :
`Get-ChildItem Cert:\CurrentUser\My | Where-Object Subject -match 'CN=CCAS'`).

```powershell
# 1. Récupérer le source à jour
git clone https://github.com/phuetz/ImageToPdf   # ou git pull dans un clone existant
cd ImageToPdf

# 2. Monter la version dans ImageToPdf/ImageToPdf.csproj
#    (<Version>, <FileVersion>, <AssemblyVersion>) puis committer/pousser.

# 3. Publier l'exécutable autonome (~82 Mo)
dotnet publish ImageToPdf/ImageToPdf.csproj -c Release -r win-x64 --self-contained true `
  -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:EnableCompressionInSingleFile=true -o .\publish

# 4. Signer + horodater
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object Thumbprint -eq '44A4CE96B17500F95F16BFB07C20470EAF5535AD'
Set-AuthenticodeSignature -FilePath .\publish\ImageToPdf.exe -Certificate $cert `
  -HashAlgorithm SHA256 -TimestampServer 'http://timestamp.digicert.com'

# 5. Vérifier AVANT de déployer
(Get-AuthenticodeSignature .\publish\ImageToPdf.exe).Status        # doit afficher : Valid
(Get-Item .\publish\ImageToPdf.exe).VersionInfo                    # bonne version + CCAS + PDF Merger
.\publish\ImageToPdf.exe -o test.pdf image1.jpg image2.jpg         # test rapide de fusion

# 6. Déployer
Copy-Item <dossier-de-distribution>\ImageToPdf.exe <dossier-de-distribution>\signing\ImageToPdf.vANCIENNE.bak.exe
Copy-Item .\publish\ImageToPdf.exe <dossier-de-distribution>\ImageToPdf.exe -Force
# puis mettre à jour la version et la date dans LISEZ-MOI.txt
```

**Si le certificat n'est plus dans le magasin** (nouveau poste, réinstallation) — l'importer depuis le PFX (mot de passe demandé) :

```powershell
Import-PfxCertificate -FilePath signing\PDFMerger-CodeSigning.pfx `
  -CertStoreLocation Cert:\CurrentUser\My -Password (Read-Host -AsSecureString 'Mot de passe du PFX')
```

---

## Sécurité — le fichier PFX

- `signing\PDFMerger-CodeSigning.pfx` contient la **clé privée** : quiconque possède ce fichier **et** son mot de passe peut signer des programmes au nom de « CCAS ».
- Ne jamais l'envoyer par mail, le copier sur un partage ouvert, ni le **committer sur GitHub** (y compris dans un secret en clair dans le workflow — utiliser les *GitHub Secrets*).
- Le `.cer` public, lui, peut circuler librement.
- `signing\` contient aussi les sauvegardes des anciens exécutables (`*.bak.exe`).

---

## Qui fait quoi — récapitulatif

| # | Action | Qui | Impact |
|---|--------|-----|--------|
| 1 | Déployer le `.cer` (GPO : Trusted Root + Trusted Publishers) | Service informatique | Plus d'alerte SmartScreen sur les postes |
| 2 | Réparer la signature CI (option A ou B) ou assumer le manuel (C) | Administrateur | Releases GitHub à jour et signées |
| — | Renouveler le certificat avant le **01/07/2029** | Administrateur + IT | Continuité de la signature |

---

*Guide créé le 02/07/2026 pour la v2.9.3.*
