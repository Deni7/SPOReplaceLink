using System.Collections.Concurrent;
using System.Globalization;
using System.Reflection;
using System.Text.Json;

namespace SPOReplaceLink.Core.Localization;

/// <summary>
/// Translations of the user-visible text. The source language of the code is Ukrainian: a Ukrainian string is the key,
/// and the catalogs <c>Localization/*.{lang}.json</c> (embedded in this assembly) map it to the other languages.
/// A string missing from a catalog is shown as is. Adding a language is adding its catalog files and a line to
/// <see cref="Languages"/>.
/// </summary>
/// <remarks>
/// The language is process-wide (<see cref="SetLanguage"/>, the desktop app and the CLI choose it at start), and a request
/// of the local API can override it for its own flow (<see cref="Use"/>). Text built once and kept (a static field,
/// a message stored in the state database) stays in the language it was built in.
/// </remarks>
public static class L
{
    public const string Ukrainian = "uk";
    public const string English = "en";

    /// <summary>Supported languages: code and the name in that language.</summary>
    public static IReadOnlyList<(string Code, string Name)> Languages { get; } = [(Ukrainian, "Українська"), (English, "English")];

    private static volatile string _language = Ukrainian;
    private static readonly AsyncLocal<string?> Scoped = new();
    private static readonly ConcurrentDictionary<string, IReadOnlyDictionary<string, string>> Catalogs = new(StringComparer.Ordinal);

    /// <summary>Raised after <see cref="SetLanguage"/> changed the language.</summary>
    public static event Action? LanguageChanged;

    /// <summary>The current language code (<c>uk</c>, <c>en</c>).</summary>
    public static string Language => Scoped.Value ?? _language;

    /// <summary>Culture for numbers and dates of the current language.</summary>
    public static CultureInfo Culture => CultureInfo.GetCultureInfo(Language == Ukrainian ? "uk-UA" : "en-US");

    /// <summary>Sets the process-wide language. Unknown codes fall back to English.</summary>
    public static void SetLanguage(string? code)
    {
        string lang = Normalize(code) ?? English;
        if (lang == _language)
            return;
        _language = lang;
        LanguageChanged?.Invoke();
    }

    /// <summary>Overrides the language for the current async flow (a request of the local API) until disposed.</summary>
    public static IDisposable Use(string? code)
    {
        string? previous = Scoped.Value;
        Scoped.Value = Normalize(code) ?? Scoped.Value;
        return new Restore(previous);
    }

    /// <summary>The supported language code for a code or culture name (<c>en-GB</c> → <c>en</c>), or null.</summary>
    public static string? Normalize(string? code)
    {
        if (string.IsNullOrWhiteSpace(code))
            return null;
        string c = code.Trim().ToLowerInvariant();
        int dash = c.IndexOfAny(['-', '_']);
        if (dash > 0)
            c = c[..dash];
        return Languages.Any(l => l.Code == c) ? c : null;
    }

    /// <summary>The language of the user interface of the system if supported, otherwise English.</summary>
    public static string FromSystem() => Normalize(CultureInfo.CurrentUICulture.Name) ?? English;

    /// <summary>The text in the current language.</summary>
    public static string T(string uk) => Translate(uk, Language);

    /// <summary><see cref="string.Format(IFormatProvider, string, object[])"/> of the translated text in the culture of the language.</summary>
    public static string F(string uk, params object?[] args) => string.Format(Culture, T(uk), args);

    /// <summary>The text in the given language.</summary>
    public static string Translate(string uk, string language)
    {
        if (language == Ukrainian || string.IsNullOrEmpty(uk))
            return uk;
        return Catalog(language).TryGetValue(uk, out var text) ? text : uk;
    }

    /// <summary>All entries of a language (the merged catalog files).</summary>
    public static IReadOnlyDictionary<string, string> Catalog(string language) => Catalogs.GetOrAdd(language, Load);

    private static IReadOnlyDictionary<string, string> Load(string language)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var assembly = typeof(L).Assembly;
        string suffix = "." + language + ".json";
        foreach (string name in assembly.GetManifestResourceNames().Where(n => n.StartsWith("i18n/", StringComparison.Ordinal)
            && n.EndsWith(suffix, StringComparison.Ordinal)).Order(StringComparer.Ordinal))
        {
            using var stream = assembly.GetManifestResourceStream(name)!;
            var entries = JsonSerializer.Deserialize<Dictionary<string, string>>(stream, new JsonSerializerOptions
            {
                ReadCommentHandling = JsonCommentHandling.Skip,
                AllowTrailingCommas = true,
            }) ?? [];
            foreach (var (key, value) in entries)
                result.TryAdd(key, value);
        }
        return result;
    }

    private sealed class Restore(string? previous) : IDisposable
    {
        public void Dispose() => Scoped.Value = previous;
    }
}
