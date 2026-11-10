using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Localization;
using SPOReplaceLink.Core.Rules;

namespace SPOReplaceLink.Cli;

internal static class ExitCodes
{
    public const int Ok = 0;
    /// <summary>Finished, but some files need attention: failed, residuals, or wait for Excel.</summary>
    public const int Attention = 1;
    /// <summary>Invalid profile or arguments; nothing was processed.</summary>
    public const int InvalidInput = 2;
}

internal static class ProfileLoader
{
    /// <summary>Loads the profile and prints every issue (CFG-7). Returns null if it has errors.</summary>
    public static (Profile Profile, RuleSet Rules)? Load(string path)
    {
        Profile profile;
        try
        {
            profile = Profile.Load(path);
        }
        catch (ProfileException ex)
        {
            Console.Error.WriteLine(L.T("ПОМИЛКА: ") + ex.Message);
            return null;
        }

        var validation = profile.Validate();
        foreach (var issue in validation.Issues)
        {
            var writer = issue.Severity == IssueSeverity.Error ? Console.Error : Console.Out;
            writer.WriteLine(issue);
        }
        if (!validation.IsValid)
        {
            Console.Error.WriteLine(L.T("Профіль містить помилки, обробку не розпочато."));
            return null;
        }
        return (profile, validation.RuleSet!);
    }
}
