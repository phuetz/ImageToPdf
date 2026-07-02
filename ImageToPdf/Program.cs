namespace ImageToPdf;

static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        // Les encodages Windows (1252…) ne sont pas inclus par défaut dans .NET :
        // indispensable pour lire les fichiers texte « ANSI » français.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Mode ligne de commande si des arguments sont fournis
        if (args.Length > 0)
        {
            return RunCommandLine(args);
        }

        // Mode GUI
        ApplicationConfiguration.Initialize();

        // Filet de sécurité : convertir toute exception non gérée en message clair
        // plutôt que d'afficher la boîte de dialogue de plantage .NET brute (JIT).
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (s, e) => ShowUnhandledError(e.Exception);
        AppDomain.CurrentDomain.UnhandledException += (s, e) => ShowUnhandledError(e.ExceptionObject as Exception);

        Application.Run(new MainForm());
        return 0;
    }

    static int RunCommandLine(string[] args)
    {
        // Attacher à la console parente si disponible
        AttachConsole(-1);

        if (args.Length == 1 && (args[0] == "-h" || args[0] == "--help" || args[0] == "/?"))
        {
            ShowHelp();
            return 0;
        }

        if (args.Length < 2)
        {
            Console.Error.WriteLine("Erreur: Il faut au moins un fichier d'entrée et un fichier de sortie.");
            Console.Error.WriteLine("Utilisez --help pour plus d'informations.");
            return 1;
        }

        string? outputFile = null;
        var inputFiles = new List<string>();
        bool verbose = false;

        // Parser les arguments
        for (int i = 0; i < args.Length; i++)
        {
            var arg = args[i];

            if (arg == "-o" || arg == "--output")
            {
                if (i + 1 < args.Length)
                {
                    outputFile = args[++i];
                }
                else
                {
                    Console.Error.WriteLine("Erreur: -o/--output nécessite un nom de fichier.");
                    return 1;
                }
            }
            else if (arg == "-v" || arg == "--verbose")
            {
                verbose = true;
            }
            else if (arg.StartsWith("-"))
            {
                Console.Error.WriteLine($"Option inconnue: {arg}");
                return 1;
            }
            else
            {
                inputFiles.Add(arg);
            }
        }

        // Si pas de -o, le dernier argument est le fichier de sortie
        bool outputInferred = false;
        if (outputFile == null && inputFiles.Count >= 2)
        {
            outputFile = inputFiles[^1];
            inputFiles.RemoveAt(inputFiles.Count - 1);
            outputInferred = true;
        }

        if (string.IsNullOrEmpty(outputFile))
        {
            Console.Error.WriteLine("Erreur: Aucun fichier de sortie spécifié.");
            return 1;
        }

        if (outputFile.Contains('*') || outputFile.Contains('?'))
        {
            Console.Error.WriteLine($"Erreur: Le fichier de sortie ne peut pas contenir de joker: '{outputFile}'.");
            return 1;
        }

        // La sortie doit être un .pdf : sans cette garde, « ImageToPdf.exe a.jpg b.jpg »
        // (oubli de la sortie) écraserait b.jpg avec le contenu du PDF.
        if (!outputFile.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            Console.Error.WriteLine($"Erreur: Le fichier de sortie doit porter l'extension .pdf (reçu: '{outputFile}').");
            Console.Error.WriteLine("Utilisez -o <sortie.pdf> pour désigner explicitement la sortie.");
            return 1;
        }

        // Sans -o explicite, refuser de réutiliser un fichier existant comme sortie :
        // un PDF d'entrée en dernière position serait sinon détruit sans confirmation.
        if (outputInferred && File.Exists(outputFile))
        {
            Console.Error.WriteLine($"Erreur: '{outputFile}' existe déjà et serait écrasé.");
            Console.Error.WriteLine("Pour écraser volontairement un fichier, utilisez -o <sortie.pdf>.");
            return 1;
        }

        string outputFullPath;
        try { outputFullPath = Path.GetFullPath(outputFile); }
        catch
        {
            Console.Error.WriteLine($"Erreur: Chemin de sortie invalide: '{outputFile}'.");
            return 1;
        }

        if (inputFiles.Count == 0)
        {
            Console.Error.WriteLine("Erreur: Aucun fichier d'entrée spécifié.");
            return 1;
        }

        // Vérifier les fichiers d'entrée
        var validFiles = new List<string>();
        foreach (var file in inputFiles)
        {
            // Supporter les wildcards
            if (file.Contains('*') || file.Contains('?'))
            {
                var dir = Path.GetDirectoryName(file);
                if (string.IsNullOrEmpty(dir)) dir = ".";
                var pattern = Path.GetFileName(file);

                try
                {
                    var matches = Directory.GetFiles(dir, pattern)
                        .Where(f => PdfMerger.IsValidExtension(f))
                        .Select(Path.GetFullPath)
                        .ToArray();
                    // Ordre naturel (page2 avant page10), comme l'Explorateur Windows —
                    // Directory.GetFiles trie « 1, 10, 2 » et mélangerait les pages scannées.
                    Array.Sort(matches, StrCmpLogicalW);
                    validFiles.AddRange(matches);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Erreur avec le pattern '{file}': {ex.Message}");
                }
            }
            else if (File.Exists(file))
            {
                if (PdfMerger.IsValidExtension(file))
                {
                    validFiles.Add(Path.GetFullPath(file));
                }
                else
                {
                    Console.Error.WriteLine($"Type de fichier non supporté: {file}");
                }
            }
            else
            {
                Console.Error.WriteLine($"Fichier introuvable: {file}");
            }
        }

        // Ne jamais reprendre la sortie comme entrée : « -o merged.pdf *.pdf » relancé
        // dans le même dossier aspirerait le merged.pdf du passage précédent.
        int selfMatches = validFiles.RemoveAll(f => string.Equals(f, outputFullPath, StringComparison.OrdinalIgnoreCase));
        if (selfMatches > 0)
        {
            Console.Error.WriteLine($"Avertissement: '{outputFile}' est à la fois entrée et sortie — ignoré comme entrée.");
        }

        if (validFiles.Count == 0)
        {
            Console.Error.WriteLine("Erreur: Aucun fichier valide à traiter.");
            return 1;
        }

        if (verbose)
        {
            Console.WriteLine($"Fichiers à fusionner: {validFiles.Count}");
            foreach (var f in validFiles)
            {
                Console.WriteLine($"  - {f}");
            }
            Console.WriteLine($"Sortie: {outputFile}");
        }

        try
        {
            Console.WriteLine($"Création du PDF avec {validFiles.Count} fichier(s)...");

            PdfMerger.CreatePdf(validFiles, outputFile, (current, total, fileName) =>
            {
                if (verbose)
                {
                    Console.WriteLine($"  [{current}/{total}] {fileName}");
                }
            });

            Console.WriteLine($"PDF créé avec succès: {outputFile}");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Erreur lors de la création du PDF: {ex.Message}");
            return 1;
        }
    }

    static void ShowHelp()
    {
        Console.WriteLine(@"
PDF Merger - Outil de fusion de fichiers en PDF

UTILISATION:
    ImageToPdf.exe [OPTIONS] <fichiers...> <sortie.pdf>
    ImageToPdf.exe -o <sortie.pdf> <fichiers...>

OPTIONS:
    -o, --output <fichier>   Spécifier le fichier PDF de sortie
    -v, --verbose            Afficher les détails du traitement
    -h, --help               Afficher cette aide

RÈGLES DE SORTIE:
    Le fichier de sortie doit porter l'extension .pdf.
    Sans -o, le dernier argument est la sortie et ne doit pas déjà exister
    (protection contre l'écrasement accidentel d'un fichier d'entrée).

FORMATS SUPPORTÉS:
    Images:    .jpg, .jpeg, .png, .bmp, .gif, .tiff, .tif
    PDF:       .pdf
    Markdown:  .md, .markdown

EXEMPLES:
    ImageToPdf.exe image1.jpg image2.png document.pdf sortie.pdf
    ImageToPdf.exe -o resultat.pdf *.jpg
    ImageToPdf.exe -v -o merged.pdf doc1.pdf doc2.pdf
    ImageToPdf.exe --output rapport.pdf images/*.png notes.md

Sans arguments, l'interface graphique est lancée.
");
    }

    static void ShowUnhandledError(Exception? ex)
    {
        var message = ex?.Message ?? "Une erreur inattendue s'est produite.";
        MessageBox.Show(
            $"Une erreur inattendue s'est produite :\n\n{message}\n\n" +
            "L'application va essayer de continuer. Si le problème persiste, redémarrez-la.",
            "PDF Merger - Erreur",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
    }

    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    static extern bool AttachConsole(int dwProcessId);

    // Tri « naturel » de l'Explorateur Windows (page2 < page10)
    [System.Runtime.InteropServices.DllImport("shlwapi.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    static extern int StrCmpLogicalW(string x, string y);
}
