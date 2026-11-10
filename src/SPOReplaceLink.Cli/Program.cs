using System.CommandLine;
using System.Text;
using SPOReplaceLink.Cli;
using SPOReplaceLink.Core.Localization;

Console.OutputEncoding = Encoding.UTF8;

// Language of the messages and of the help: --lang, else SPO_RL_LANG, else the language of the system.
// Chosen before the commands are built, since their descriptions are translated.
L.SetLanguage(LanguageOption(args) ?? L.Normalize(Environment.GetEnvironmentVariable("SPO_RL_LANG")) ?? L.FromSystem());

// Ctrl+C while waiting asks the host for a graceful stop instead of killing the CLI at once.
using var interrupt = new CancellationTokenSource();
Console.CancelKeyPress += (_, e) =>
{
    if (!interrupt.IsCancellationRequested)
    {
        e.Cancel = true;
        interrupt.Cancel();
    }
};

var configOption = new Option<FileInfo>("--config", "-c")
{
    Description = L.T("Профіль у JSON (правила, місця заміни, фільтри)"),
    Required = true,
};
var baseUrlOption = new Option<string?>("--base-url")
{
    Description = L.T("Стара адреса папки (до міграції), де лежали файли, — для відносних посилань"),
};

// local
var inputArgument = new Argument<string>("input") { Description = L.T("Папка або файл") };
var outOption = new Option<DirectoryInfo?>("--out", "-o") { Description = L.T("Куди класти змінені копії (з --write)") };
var writeOption = new Option<bool>("--write") { Description = L.T("Записати змінені копії; без нього — лише звіт (dry-run)") };
var reportOption = new Option<DirectoryInfo?>("--report", "-r") { Description = L.T("Папка для CSV-звітів (за замовчуванням reports/<час>)") };
var parallelOption = new Option<int>("--parallel", "-p")
{
    Description = L.T("Кількість файлів, що обробляються одночасно"),
    DefaultValueFactory = _ => Environment.ProcessorCount,
};
var delimiterOption = new Option<string>("--csv-delimiter")
{
    Description = L.T("Роздільник у CSV"),
    DefaultValueFactory = _ => ",",
};
var excelOption = new Option<int?>("--excel")
{
    Description = L.T("Процесів Excel для xls і залишків xlsb (за замовчуванням — з профілю; 0 — без Excel)"),
};
var tempOption = new Option<DirectoryInfo?>("--temp")
{
    Description = L.T("Тимчасова папка (для Excel — на іншому диску, ніж локальні посилання у файлах)"),
};

var local = new Command("local", L.T("Обробити локальну папку або файл: заміна посилань, звіти CSV"));
local.Arguments.Add(inputArgument);
local.Options.Add(configOption);
local.Options.Add(outOption);
local.Options.Add(writeOption);
local.Options.Add(baseUrlOption);
local.Options.Add(reportOption);
local.Options.Add(parallelOption);
local.Options.Add(delimiterOption);
local.Options.Add(excelOption);
local.Options.Add(tempOption);
local.SetAction(parse => LocalCommand.RunAsync(new LocalOptions(
    parse.GetValue(inputArgument)!,
    parse.GetValue(configOption)!.FullName,
    parse.GetValue(outOption)?.FullName,
    parse.GetValue(writeOption),
    parse.GetValue(baseUrlOption),
    parse.GetValue(reportOption)?.FullName,
    parse.GetValue(parallelOption),
    parse.GetValue(delimiterOption)!,
    parse.GetValue(excelOption),
    parse.GetValue(tempOption)?.FullName), interrupt.Token));

// rules validate / rules test
var validate = new Command("validate", L.T("Перевірити профіль і набір правил"));
validate.Options.Add(configOption);
validate.SetAction(parse => RulesCommands.Validate(parse.GetValue(configOption)!.FullName));

var urlArgument = new Argument<string[]>("url") { Description = L.T("Одне або кілька посилань"), Arity = ArgumentArity.OneOrMore };
var test = new Command("test", L.T("Показати, яке правило спрацює для посилання і що вийде"));
test.Arguments.Add(urlArgument);
test.Options.Add(configOption);
test.Options.Add(baseUrlOption);
test.SetAction(parse => RulesCommands.Test(
    parse.GetValue(configOption)!.FullName, parse.GetValue(urlArgument)!, parse.GetValue(baseUrlOption)));

var rules = new Command("rules", L.T("Робота з правилами заміни"));
rules.Subcommands.Add(validate);
rules.Subcommands.Add(test);

var root = new RootCommand(L.T("SPOReplaceLink — заміна битих посилань в Excel-файлах після міграції в SharePoint Online"));
root.Subcommands.Add(local);
root.Subcommands.Add(rules);

// ---------------------------------------------------------------- commands that work through the host

var hostOption = new Option<string?>("--host")
{
    Description = L.T("Адреса хоста (за замовчуванням SPO_RL_HOST, інакше адреса застосунку цього користувача з api.json, інакше http://localhost:5080)"),
    Recursive = true,
};
root.Options.Add(hostOption);
var dataOption = new Option<string?>("--data")
{
    Description = L.T("Папка даних застосунку, якщо його запущено з --data (звідти CLI бере адресу й ключ API; за замовчуванням SPO_RL_DATA або %LOCALAPPDATA%\\SPOReplaceLink)"),
    Recursive = true,
};
root.Options.Add(dataOption);
var langOption = new Option<string?>("--lang")
{
    Description = L.T("Мова повідомлень: uk або en (за замовчуванням SPO_RL_LANG, інакше мова системи)"),
    Recursive = true,
};
root.Options.Add(langOption);


Command Cmd(string name, string description, Func<ParseResult, HostCommands, Task<int>> action, params Symbol[] symbols)
{
    var command = new Command(name, description);
    foreach (var symbol in symbols)
    {
        if (symbol is Argument argument)
            command.Arguments.Add(argument);
        else
            command.Options.Add((Option)symbol);
    }
    command.SetAction(async (parse, _) =>
    {
        var endpoint = SPOReplaceLink.Core.Hosting.ApiEndpoint.Read(SPOReplaceLink.Core.Hosting.ApiEndpoint.DataDirectory(
            parse.GetValue(dataOption) ?? Environment.GetEnvironmentVariable("SPO_RL_DATA")));
        string url = parse.GetValue(hostOption) ?? Environment.GetEnvironmentVariable("SPO_RL_HOST")
            ?? endpoint?.Url ?? SPOReplaceLink.Core.Hosting.ApiEndpoint.DefaultUrl;
        try
        {
            return await action(parse, new HostCommands(HostClient.Create(url, key: endpoint?.Key), Console.Out));
        }
        catch (HostException ex)
        {
            Console.Error.WriteLine(L.T("ПОМИЛКА: ") + ex.Message);
            foreach (string issue in ex.Issues)
                Console.Error.WriteLine("  " + issue);
            return ExitCodes.InvalidInput;
        }
    });
    return command;
}

var nameArgument = new Argument<string>("name") { Description = L.T("Назва") };
var fileArgument = new Argument<string>("file") { Description = L.T("JSON-файл") };
var importName = new Option<string?>("--name") { Description = L.T("Назва (за замовчуванням з поля name або імені файлу)") };
var outputFile = new Option<string?>("--output", "-o") { Description = L.T("Записати у файл") };
var profileOption = new Option<string?>("--profile") { Description = L.T("Профіль") };
foreach (var (kind, title) in new[] { ("profiles", L.T("Профілі прогону")), ("connections", L.T("Підключення до SharePoint (застосунки Entra)")), ("templates", L.T("Шаблони правил")) })
{
    var group = new Command(kind, title);
    group.Subcommands.Add(Cmd("list", L.T("Список"), (_, h) => h.ListSettings(kind)));
    group.Subcommands.Add(Cmd("get", L.T("Показати або експортувати в JSON"), (p, h) => h.GetSetting(kind, p.GetValue(nameArgument)!, p.GetValue(outputFile)), nameArgument, outputFile));
    group.Subcommands.Add(Cmd("import", L.T("Зберегти з JSON-файлу (створити або замінити)"), (p, h) => h.PutSetting(kind, p.GetValue(fileArgument)!, p.GetValue(importName)), fileArgument, importName));
    group.Subcommands.Add(Cmd("delete", L.T("Видалити"), (p, h) => h.DeleteSetting(kind, p.GetValue(nameArgument)!), nameArgument));
    if (kind == "profiles")
        group.Subcommands.Add(Cmd("validate", L.T("Перевірити профіль як перед запуском"), (p, h) => h.ValidateProfile(p.GetValue(nameArgument)!), nameArgument));
    if (kind == "connections")
        group.Subcommands.Add(Cmd("test", L.T("Отримати токен і перевірити доступ до цілей профілю"), (p, h) => h.TestConnection(p.GetValue(nameArgument)!, p.GetValue(profileOption)), nameArgument, profileOption));
    root.Subcommands.Add(group);
}

var runProfile = new Argument<string>("profile") { Description = L.T("Профіль прогону") };
var waitOption = new Option<bool>("--wait", "-w") { Description = L.T("Чекати завершення, показувати прогрес") };
var yesOption = new Option<bool>("--yes", "-y") { Description = L.T("Не питати підтвердження (apply, видалення прогону)") };
foreach (var (mode, description) in new[]
{
    ("inventory", L.T("Перелік файлів і гістограма префіксів посилань (без запису)")),
    ("scan", L.T("Dry-run: що буде замінено (без запису)")),
    ("apply", L.T("Заміна і запис у SharePoint")),
})
{
    root.Subcommands.Add(Cmd(mode, description, (p, h) => h.StartRun(mode, p.GetValue(runProfile)!, p.GetValue(waitOption), p.GetValue(yesOption), interrupt.Token),
        runProfile, waitOption, yesOption));
}

var runArgument = new Argument<long>("run") { Description = L.T("Номер прогону") };
var verifyRun = new Argument<long?>("run") { Description = L.T("Прогін apply, файли якого перевірити"), Arity = ArgumentArity.ZeroOrOne };
root.Subcommands.Add(Cmd("verify", L.T("Перевірити, що старих префіксів не лишилось: файли прогону apply або весь профіль (--profile)"), (p, h) =>
    p.GetValue(verifyRun) is long run
        ? h.VerifyRun(run, p.GetValue(waitOption), interrupt.Token)
        : p.GetValue(profileOption) is { } profile
            ? h.StartRun("verify", profile, p.GetValue(waitOption), true, interrupt.Token)
            : throw new HostException(L.T("Вкажіть номер прогону apply або --profile.")),
    verifyRun, profileOption, waitOption));

var statusesOption = new Option<string[]>("--status") { Description = L.T("Статуси (failed, done:residual, deferred…; attention — усі, що потребують уваги); за замовчуванням failed"), AllowMultipleArgumentsPerToken = true };
var classesOption = new Option<string[]>("--class") { Description = L.T("Класи помилок (transient, conflict, locked, excel, auth, file, validation, internal)"), AllowMultipleArgumentsPerToken = true };
var extOption = new Option<string[]>("--ext") { Description = L.T("Розширення"), AllowMultipleArgumentsPerToken = true };
var maskOption = new Option<string?>("--path-mask") { Description = L.T("Маска шляху (* і ?)") };
var idsOption = new Option<long[]>("--file-id") { Description = L.T("Номери файлів"), AllowMultipleArgumentsPerToken = true };
var listOption = new Option<string?>("--list") { Description = L.T("CSV зі шляхами (стовпець path або перший)") };
var modeOption = new Option<string?>("--mode") { Description = L.T("Режим повтору (за замовчуванням як у прогоні)") };
var currentOption = new Option<bool>("--current-profile") { Description = L.T("Взяти поточну версію профілю замість знімка прогону") };
root.Subcommands.Add(Cmd("retry", L.T("Повторити вибрані файли прогону без нового переліку бібліотеки"), (p, h) => h.Retry(p.GetValue(runArgument), new RetryArgs(
        p.GetValue(statusesOption) ?? [], p.GetValue(classesOption) ?? [], p.GetValue(extOption) ?? [], p.GetValue(maskOption),
        p.GetValue(idsOption) ?? [], p.GetValue(listOption), p.GetValue(modeOption), p.GetValue(currentOption)),
        p.GetValue(waitOption), p.GetValue(yesOption), interrupt.Token),
    runArgument, statusesOption, classesOption, extOption, maskOption, idsOption, listOption, modeOption, currentOption, waitOption, yesOption));

var takeOption = new Option<int>("--take") { Description = L.T("Скільки показати"), DefaultValueFactory = _ => 20 };
root.Subcommands.Add(Cmd("runs", L.T("Список прогонів"), (p, h) => h.Runs(p.GetValue(takeOption)), takeOption));
root.Subcommands.Add(Cmd("status", L.T("Стан прогону"), (p, h) => h.Status(p.GetValue(runArgument)), runArgument));
root.Subcommands.Add(Cmd("wait", L.T("Чекати завершення прогону"), (p, h) => h.Wait(p.GetValue(runArgument), interrupt.Token), runArgument));
foreach (var (action, description) in new[]
{
    ("start", L.T("Продовжити зупинений або перерваний прогін")),
    ("pause", L.T("Пауза: нові файли не беруться")),
    ("resume", L.T("Продовжити після паузи")),
    ("stop", L.T("Зупинити після завершення поточних файлів")),
})
{
    root.Subcommands.Add(Cmd(action, description, (p, h) => h.Control(p.GetValue(runArgument), action, p.GetValue(waitOption), interrupt.Token), runArgument, waitOption));
}

root.Subcommands.Add(Cmd("delete-run", L.T("Видалити прогін, що не виконується, з усіма його результатами (у SharePoint нічого не змінюється)"),
    (p, h) => h.DeleteRun(p.GetValue(runArgument), p.GetValue(yesOption), interrupt.Token), runArgument, yesOption));

var fileStatusOption = new Option<string?>("--status") { Description = L.T("Статус (через кому)") };
var fileClassOption = new Option<string?>("--class") { Description = L.T("Клас помилки") };
var searchOption = new Option<string?>("--search") { Description = L.T("Частина шляху") };
root.Subcommands.Add(Cmd("files", L.T("Файли прогону"), (p, h) => h.Files(p.GetValue(runArgument), p.GetValue(fileStatusOption), p.GetValue(fileClassOption), p.GetValue(searchOption), p.GetValue(takeOption)),
    runArgument, fileStatusOption, fileClassOption, searchOption, takeOption));

var fileIdArgument = new Argument<long>("file") { Description = L.T("Номер файлу") };
var commentOption = new Option<string>("--comment") { Description = L.T("Чому файл ігнорується"), Required = true };
root.Subcommands.Add(Cmd("ignore", L.T("Не брати файл у наступні повтори"), (p, h) => h.Ignore(p.GetValue(fileIdArgument), p.GetValue(commentOption)!), fileIdArgument, commentOption));

var exportDir = new Option<string>("--output", "-o") { Description = L.T("Папка для звітів"), DefaultValueFactory = _ => "." };
var prefixDepthOption = new Option<int?>("--prefix-depth") { Description = L.T("Скільки рівнів папок під сайтом, сервером чи диском лишати у звіті prefixes (1–10, за замовчуванням 2)") };
prefixDepthOption.Validators.Add(r =>
{
    if (r.GetValueOrDefault<int?>() is < 1 or > 10)
        r.AddError(L.T("--prefix-depth: від 1 до 10."));
});
root.Subcommands.Add(Cmd("export", L.T("Звіти прогону: files, links, residuals, attention, prefixes, sample, summary, pilot"),
    (p, h) => h.Export(p.GetValue(runArgument), p.GetValue(exportDir)!, p.GetValue(prefixDepthOption)), runArgument, exportDir, prefixDepthOption));

return await root.Parse(args).InvokeAsync(new InvocationConfiguration { ProcessTerminationTimeout = null });

static string? LanguageOption(string[] args)
{
    for (int i = 0; i < args.Length; i++)
    {
        if (args[i] == "--lang" && i + 1 < args.Length)
            return L.Normalize(args[i + 1]);
        if (args[i].StartsWith("--lang=", StringComparison.Ordinal))
            return L.Normalize(args[i]["--lang=".Length..]);
    }
    return null;
}
