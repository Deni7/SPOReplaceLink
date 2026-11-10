using Microsoft.Data.Sqlite;
using SPOReplaceLink.State;

namespace SPOReplaceLink.Tests.Shared;

/// <summary>
/// A large finished run written straight into the state database — for checking how the GUI and the API
/// behave on hundreds of thousands of files (plan, stage 3, item 12) without running SharePoint traffic.
/// </summary>
public static class SyntheticRun
{
    public static long Create(StateStore store, string dbPath, string profileName, string snapshot, int files, int seed = 1)
    {
        long runId = store.CreateRun(RunModes.Apply, null, profileName, snapshot, enumerate: true, targets: 2, DateTimeOffset.UtcNow.AddHours(-6));
        store.SetRunStatus(runId, RunStatuses.Running, DateTimeOffset.UtcNow.AddHours(-6));
        var rnd = new Random(seed);
        string[] exts = [".xlsx", ".xlsx", ".xlsx", ".xlsm", ".xlsb", ".xls"];
        string[] folders = ["Fin/2023", "Fin/2024", "Analytics", "Analytics/Q1", "Бухгалтерія/Звіти", "HR"];
        var started = DateTimeOffset.UtcNow.AddHours(-6);

        using var c = new SqliteConnection(new SqliteConnectionStringBuilder { DataSource = dbPath }.ToString());
        c.Open();
        using var tx = c.BeginTransaction();
        using var insertFile = c.CreateCommand();
        insertFile.Transaction = tx;
        insertFile.CommandText = """
            INSERT INTO files(run_id, target_idx, drive_id, item_id, name, path, ext, size, etag, status, status_code, error_class, error,
                              attempts, replacements, residuals, needs_excel, app, version_before, updated_at)
            VALUES($r,$t,'drive-1',$i,$n,$p,$e,$s,'"{x},1"',$st,$sc,$ec,$err,$a,$rep,$res,$ne,'demo-app',$vb,$u);
            SELECT last_insert_rowid();
            """;
        foreach (string p in new[] { "$r", "$t", "$i", "$n", "$p", "$e", "$s", "$st", "$sc", "$ec", "$err", "$a", "$rep", "$res", "$ne", "$vb", "$u" })
            insertFile.Parameters.Add(new SqliteParameter(p, DBNull.Value));
        using var insertLink = c.CreateCommand();
        insertLink.Transaction = tx;
        insertLink.CommandText = "INSERT INTO links(file_id, part, kind, action, old_value, new_value, rule) VALUES($f,'xl/worksheets/_rels/sheet1.xml.rels',$k,$a,$o,$nv,1)";
        foreach (string p in new[] { "$f", "$k", "$a", "$o", "$nv" })
            insertLink.Parameters.Add(new SqliteParameter(p, DBNull.Value));

        for (int n = 0; n < files; n++)
        {
            string ext = exts[rnd.Next(exts.Length)];
            string folder = folders[rnd.Next(folders.Length)];
            string name = $"Звіт {n:000000}{ext}";
            int roll = rnd.Next(1000);
            (string status, string code, string? cls, string? err) = roll switch
            {
                < 20 => ("failed:transient", "failed:transient", "transient", "503 serverBusy: fake"),
                < 30 => ("failed:locked", "failed:locked", "locked", "Файл узято на редагування (check-out)."),
                < 35 => ("failed:file", "failed:file", "file", "Пошкоджений архів або частина."),
                < 60 => ("done", "done:residual", null, null),
                < 75 => ("needs_excel", "needs_excel", null, "Формат xls обробляється через Excel (ExcelWorker)."),
                < 500 => ("done", "done", null, null),
                < 980 => ("unchanged", "unchanged", null, null),
                _ => ("pending", "pending", null, null),
            };
            int replacements = status == "done" ? 1 + rnd.Next(4) : 0;
            Set(insertFile, ("$r", runId), ("$t", rnd.Next(2)), ("$i", $"i{n}"), ("$n", name), ("$p", $"{folder}/{name}"), ("$e", ext),
                ("$s", 20_000 + rnd.Next(5_000_000)), ("$st", status), ("$sc", code), ("$ec", cls), ("$err", err), ("$a", cls is null ? 1 : 3),
                ("$rep", replacements), ("$res", code.EndsWith(":residual", StringComparison.Ordinal) ? 1 : 0), ("$ne", status == "needs_excel" ? 1 : 0),
                ("$vb", status == "done" ? "3.0" : null), ("$u", StateStore.Now(DateTimeOffset.UtcNow)));
            long id = (long)insertFile.ExecuteScalar()!;
            for (int k = 0; k < Math.Max(1, replacements); k++)
            {
                bool replaced = k < replacements;
                Set(insertLink, ("$f", id), ("$k", k % 3 == 0 ? "Hyperlink" : k % 3 == 1 ? "LinkSource" : "Formula"), ("$a", replaced ? "Replaced" : "NoRule"),
                    ("$o", $"https://contoso.sharepoint.com/Document/{folder}/src{k}.xlsx"),
                    ("$nv", replaced ? $"https://contoso.sharepoint.com/sites/finance/{folder}/src{k}.xlsx" : null));
                insertLink.ExecuteNonQuery();
            }
        }

        using var metrics = c.CreateCommand();
        metrics.Transaction = tx;
        metrics.CommandText = "INSERT INTO run_metrics(run_id, at, json) VALUES($r,$at,$j)";
        metrics.Parameters.Add(new SqliteParameter("$r", runId));
        metrics.Parameters.Add(new SqliteParameter("$at", DBNull.Value));
        metrics.Parameters.Add(new SqliteParameter("$j", DBNull.Value));
        for (int m = 0; m <= 360; m += 5)
        {
            long done = (long)(files * Math.Min(1.0, m / 360.0));
            metrics.Parameters["$at"].Value = StateStore.Now(started.AddMinutes(m));
            metrics.Parameters["$j"].Value = string.Create(System.Globalization.CultureInfo.InvariantCulture,
                $$"""{"at":"{{started.AddMinutes(m):o}}","found":{{files}},"finished":{{done}},"filesPerMinute":{{(m == 0 ? 0 : files / 360.0 * (0.8 + (rnd.NextDouble() * 0.4))):0.0}},"megabytesPerSecond":1.5,"activeWorkers":12,"replacements":{},"errors":{"transient":{{m / 10}}},"apps":[{"name":"demo-app","ruSpent":{{done * 4}},"ruLast24h":{{done * 4}},"minuteUtilization":0.7,"minuteCapacity":1000,"throttled429":{{m / 60}},"throttled503":0,"requests":{{done * 3}},"pausedUntil":null,"concurrencyLimit":12,"activeFiles":8,"blocked":false}],"tenantSpent":{{done * 4}},"tenantUtilization":0.3}""");
            metrics.ExecuteNonQuery();
        }
        tx.Commit();
        store.SetRunStatus(runId, RunStatuses.CompletedWithErrors, DateTimeOffset.UtcNow);
        return runId;
    }

    private static void Set(SqliteCommand cmd, params (string Name, object? Value)[] values)
    {
        foreach (var (name, value) in values)
            cmd.Parameters[name].Value = value ?? DBNull.Value;
    }
}
