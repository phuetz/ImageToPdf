# PDF Merger : Un outil open source pour fusionner vos documents

*Comment j'ai créé une application Windows simple pour résoudre un problème du quotidien*

---

## Le problème

Vous avez sûrement déjà vécu cette situation : vous devez envoyer plusieurs documents à quelqu'un — des photos, des scans, peut-être quelques PDF — et vous aimeriez tout regrouper dans un seul fichier. Les solutions existantes sont soit payantes, soit trop complexes, soit remplies de publicités.

C'est de cette frustration qu'est né **PDF Merger**.

## La solution

PDF Merger est une application Windows gratuite et open source qui fait une chose et la fait bien : fusionner des fichiers en un seul PDF.

### Ce que vous pouvez fusionner

- **Images** : JPG, PNG, BMP, GIF, TIFF
- **Documents PDF** : toutes les pages sont importées
- **Fichiers Markdown** : convertis en pages PDF formatées

### Comment ça marche

L'interface ressemble à l'explorateur Windows que vous connaissez déjà. Vous pouvez :

1. Glisser-déposer vos fichiers directement dans la fenêtre
2. Réorganiser l'ordre avec des flèches ou par glisser-déposer
3. Prévisualiser chaque fichier avant la fusion
4. Créer votre PDF en un clic

Pas d'inscription, pas de limite, pas de filigrane.

## Les choix techniques

PDF Merger est développé en C# avec .NET 8 et Windows Forms. J'ai fait le choix de la simplicité :

- **Exécutable autonome** : un seul fichier, aucune installation
- **Portable** : fonctionne depuis une clé USB
- **Léger** : démarre instantanément

Pour la génération PDF, j'utilise [PdfSharpCore](https://github.com/ststeiger/PdfSharpCore), une bibliothèque open source robuste et bien maintenue.

## Mode ligne de commande

Pour les utilisateurs avancés et l'automatisation, PDF Merger fonctionne aussi en ligne de commande :

```bash
# Fusionner des images et PDF
ImageToPdf.exe photo1.jpg photo2.jpg scan.pdf resultat.pdf

# Utiliser des wildcards
ImageToPdf.exe -o rapport.pdf *.png *.pdf

# Mode verbose
ImageToPdf.exe -v -o output.pdf input1.pdf input2.pdf
```

Idéal pour intégrer dans des scripts ou des workflows automatisés.

## Trois éditions pour différents besoins

J'ai créé trois éditions pour répondre à différents cas d'usage :

| Édition | Pour qui ? |
|---------|-----------|
| **Lite** | Fusion basique, léger et rapide |
| **Standard** | Usage quotidien avec aperçu et Markdown |
| **Pro** | Options avancées : numérotation, filigrane, formats |

Toutes les éditions sont gratuites et open source.

## Open source et transparent

Le code source complet est disponible sur GitHub sous licence MIT. Vous pouvez :

- Vérifier exactement ce que fait l'application
- Proposer des améliorations
- L'adapter à vos besoins
- Le redistribuer librement

## Téléchargement

PDF Merger est disponible gratuitement sur GitHub :

**[Télécharger PDF Merger](https://github.com/phuetz/ImageToPdf/releases)**

Configuration requise : Windows 10 ou 11 (64 bits)

---

## Conclusion

Parfois, les meilleurs outils sont les plus simples. PDF Merger ne prétend pas révolutionner quoi que ce soit — il résout juste un problème concret de manière efficace.

Si vous avez des suggestions ou trouvez des bugs, n'hésitez pas à ouvrir une issue sur GitHub. Les contributions sont les bienvenues !

---

*Patrice Huetz — Développeur passionné*

*[GitHub](https://github.com/phuetz) | [PDF Merger](https://github.com/phuetz/ImageToPdf)*
