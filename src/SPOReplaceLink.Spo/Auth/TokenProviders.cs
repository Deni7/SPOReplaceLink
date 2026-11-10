using System.Security.Cryptography.X509Certificates;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using SPOReplaceLink.Core.Config;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Core.Localization;

namespace SPOReplaceLink.Spo.Auth;

/// <summary>Access tokens for Graph. MSAL caches them and renews before they expire (RSL-1).</summary>
public interface ITokenProvider
{
    /// <param name="forceRefresh">After a 401: skip the cache and get a new token.</param>
    Task<string> GetTokenAsync(bool forceRefresh, CancellationToken ct);
}

public static class TokenProviderFactory
{
    public static ITokenProvider Create(Connection connection, ILogger logger) => connection.AuthMode switch
    {
        "deviceCode" => new DeviceCodeTokenProvider(connection, logger),
        _ => new CertificateTokenProvider(connection),
    };
}

/// <summary>SEC-2: app-only with a certificate from the Windows store (CFG-6: only the thumbprint is configured).</summary>
public sealed class CertificateTokenProvider : ITokenProvider
{
    private static readonly string[] Scopes = ["https://graph.microsoft.com/.default"];
    private readonly IConfidentialClientApplication _app;

    public CertificateTokenProvider(Connection connection)
    {
        var certificate = FindCertificate(connection.CertificateThumbprint!, connection.CertificateStore);
        _app = ConfidentialClientApplicationBuilder.Create(connection.ClientId)
            .WithCertificate(certificate, sendX5C: true)
            .WithAuthority(AzureCloudInstance.AzurePublic, connection.TenantId)
            .Build();
    }

    public async Task<string> GetTokenAsync(bool forceRefresh, CancellationToken ct)
    {
        try
        {
            var result = await _app.AcquireTokenForClient(Scopes).WithForceRefresh(forceRefresh).ExecuteAsync(ct).ConfigureAwait(false);
            return result.AccessToken;
        }
        catch (MsalServiceException ex) when (ex.StatusCode is 400 or 401 or 403)
        {
            throw new SpoException(ErrorClass.Auth, L.F("Не вдалося отримати токен: {0} {1}", ex.ErrorCode, ex.Message), inner: ex);
        }
        catch (MsalException ex)
        {
            throw new SpoException(ErrorClass.Transient, L.F("Не вдалося отримати токен: {0} {1}", ex.ErrorCode, ex.Message), inner: ex);
        }
    }

    public static X509Certificate2 FindCertificate(string thumbprint, string storeName)
    {
        var location = storeName == "LocalMachine" ? StoreLocation.LocalMachine : StoreLocation.CurrentUser;
        using var store = new X509Store(StoreName.My, location);
        store.Open(OpenFlags.ReadOnly);
        var found = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint.Replace(" ", string.Empty), validOnly: false);
        var cert = found.Count > 0 ? found[0] : null;
        if (cert is null)
            throw new SpoException(ErrorClass.Auth, L.F("Сертифікат {0} не знайдено у сховищі {1}\\My.", thumbprint, storeName));
        if (!cert.HasPrivateKey)
            throw new SpoException(ErrorClass.Auth, L.F("Сертифікат {0} у сховищі {1}\\My без закритого ключа (імпортовано лише .cer). Виберіть сховище, де сертифікат створено, або створіть новий (довідка, 1.5, крок 1).", thumbprint, storeName));
        try
        {
            using var key = cert.GetRSAPrivateKey();
        }
        catch (System.Security.Cryptography.CryptographicException ex)
        {
            throw new SpoException(ErrorClass.Auth,
                L.F("Немає доступу до закритого ключа сертифіката {0} у {1}\\My для {2}\\{3}: {4} Адміністратор: certlm.msc → сертифікат → Manage Private Keys → право Read для цього запису.",
                    thumbprint, storeName, Environment.UserDomainName, Environment.UserName, ex.Message), inner: ex);
        }
        if (cert.NotAfter < DateTime.Now)
            throw new SpoException(ErrorClass.Auth, L.F("Сертифікат {0} прострочений ({1:d}).", thumbprint, cert.NotAfter));
        return cert;
    }
}

/// <summary>SEC-2 fallback: delegated sign-in with a device code; the code is written to the host log.</summary>
public sealed class DeviceCodeTokenProvider(Connection connection, ILogger logger) : ITokenProvider
{
    private static readonly string[] Scopes = ["https://graph.microsoft.com/Sites.ReadWrite.All"];
    private readonly IPublicClientApplication _app = PublicClientApplicationBuilder.Create(connection.ClientId)
        .WithAuthority(AzureCloudInstance.AzurePublic, connection.TenantId)
        .WithDefaultRedirectUri()
        .Build();
    private readonly SemaphoreSlim _interactive = new(1, 1);

    public async Task<string> GetTokenAsync(bool forceRefresh, CancellationToken ct)
    {
        var account = (await _app.GetAccountsAsync().ConfigureAwait(false)).FirstOrDefault();
        if (account is not null)
        {
            try
            {
                var silent = await _app.AcquireTokenSilent(Scopes, account).WithForceRefresh(forceRefresh).ExecuteAsync(ct).ConfigureAwait(false);
                return silent.AccessToken;
            }
            catch (MsalUiRequiredException)
            {
            }
        }

        await _interactive.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            var result = await _app.AcquireTokenWithDeviceCode(Scopes, code =>
            {
                logger.LogWarning(L.T("Вхід для підключення {Connection}: {Message}"), connection.Name, code.Message);
                return Task.CompletedTask;
            }).ExecuteAsync(ct).ConfigureAwait(false);
            return result.AccessToken;
        }
        catch (MsalException ex)
        {
            throw new SpoException(ErrorClass.Auth, L.F("Вхід за кодом не виконано: {0} {1}", ex.ErrorCode, ex.Message), inner: ex);
        }
        finally
        {
            _interactive.Release();
        }
    }
}
