namespace SPOReplaceLink.Core.Localization;

/// <summary>The language chosen in the app, kept in <c>language.txt</c> of the data folder.</summary>
public sealed class LanguageStore(string dataDir)
{
    private string FilePath => Path.Combine(dataDir, "language.txt");

    /// <summary>The saved language, or null when none was chosen yet.</summary>
    public string? Read()
    {
        try
        {
            return File.Exists(FilePath) ? L.Normalize(File.ReadAllText(FilePath)) : null;
        }
        catch (IOException)
        {
            return null;
        }
        catch (UnauthorizedAccessException)
        {
            return null;
        }
    }

    /// <summary>Saves the language and makes it current.</summary>
    public void Save(string code)
    {
        string lang = L.Normalize(code) ?? L.English;
        Directory.CreateDirectory(dataDir);
        File.WriteAllText(FilePath, lang);
        L.SetLanguage(lang);
    }

    /// <summary>Makes the saved language current (the system one when none is saved).</summary>
    public string Apply()
    {
        string lang = Read() ?? L.FromSystem();
        L.SetLanguage(lang);
        return lang;
    }
}
