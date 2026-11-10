using SPOReplaceLink.Core.Localization;
using SPOReplaceLink.Core.Urls;

namespace SPOReplaceLink.Cli;

internal static class RulesCommands
{
    public static int Validate(string config)
    {
        var loaded = ProfileLoader.Load(config);
        if (loaded is null)
            return ExitCodes.InvalidInput;
        Console.WriteLine(L.F("Профіль коректний. Діючих правил: {0}.", loaded.Value.Rules.Rules.Count));
        return ExitCodes.Ok;
    }

    /// <summary>RUL-9: which rule matches each link and what the result is.</summary>
    public static int Test(string config, string[] urls, string? baseUrl)
    {
        var loaded = ProfileLoader.Load(config);
        if (loaded is null)
            return ExitCodes.InvalidInput;
        var rules = loaded.Value.Rules;

        foreach (string url in urls)
        {
            Console.WriteLine();
            Console.WriteLine(L.F("Посилання:     {0}", url));
            var result = rules.Rewrite(url, baseUrl);
            if (result.ResolvedFrom is not null)
                Console.WriteLine(L.F("Розв'язано як: {0}", result.ResolvedFrom));

            string normalized = NormalizedLink.Normalize(result.ResolvedFrom ?? url);
            Console.WriteLine(L.F("Нормалізовано: {0}", normalized));
            var candidates = rules.FindAll(normalized).ToList();
            foreach (var rule in candidates)
                Console.WriteLine($"  {(rule == result.Rule ? "✔" : " ")} {rule}");

            switch (result.Outcome)
            {
                case Core.Rules.RewriteOutcome.Replaced:
                    Console.WriteLine(L.F("Результат:     {0}", result.NewValue));
                    var again = rules.Rewrite(result.NewValue!, baseUrl);
                    if (again.Changed)
                        Console.WriteLine(L.F("УВАГА: повторна заміна дасть {0}", again.NewValue));
                    break;
                case Core.Rules.RewriteOutcome.AlreadyCurrent:
                    Console.WriteLine(L.T("Результат:     без змін (вже нова адреса)"));
                    break;
                case Core.Rules.RewriteOutcome.NoRule:
                    Console.WriteLine(L.T("Результат:     жодне правило не спрацювало"));
                    break;
                case Core.Rules.RewriteOutcome.Unresolved:
                    Console.WriteLine(L.T("Результат:     відносне посилання — задайте --base-url (стару адресу папки документа) або baseHost у профілі"));
                    break;
                default:
                    Console.WriteLine(L.T("Результат:     не є посиланням на розташування (внутрішнє або порожнє)"));
                    break;
            }
        }
        return ExitCodes.Ok;
    }
}
