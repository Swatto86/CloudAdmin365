namespace CloudAdmin365.Services.Implementations;

using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using CloudAdmin365.Services;
using Microsoft.Identity.Client;

/// <summary>
/// Azure AD authentication service using MSAL with interactive browser login.
/// Uses Microsoft Graph PowerShell's public client ID for consent-free authentication.
/// </summary>
public class AzureIdentityAuthService : IAuthService
{
    // Microsoft Graph PowerShell public client (pre-consented for common Graph scopes)
    private const string PUBLIC_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
    private const string TENANT_ID = "organizations"; // Multi-tenant: any Azure AD
    private const string REDIRECT_URI = "http://localhost";

    private const string GraphScopePrefix = "https://graph.microsoft.com/";

    // Only request User.Read at login (no admin consent needed).
    // Other scopes are requested on-demand via GetAccessTokenAsync() when audit features are used.
    // This allows non-admin service desk users to authenticate successfully.
    // Admin-only scopes (AuditLog.Read.All, Calendars.Read.All, Group.Read.All) will be requested
    // interactively if the user tries to use those features—they'll succeed for admins, fail for non-admins.
    private readonly string[] _defaultScopes = new[]
    {
        "User.Read"
    };

    private readonly IPublicClientApplication _clientApp;
    private IAccount? _currentAccount;
    private AuthUser? _currentUser;

    public AuthUser? CurrentUser => _currentUser;
    public bool IsAuthenticated => _currentAccount != null && _currentUser != null;

    /// <summary>
    /// Initialize MSAL public client application.
    /// </summary>
    public AzureIdentityAuthService()
    {
        _clientApp = PublicClientApplicationBuilder
            .Create(PUBLIC_CLIENT_ID)
            .WithAuthority(AzureCloudInstance.AzurePublic, TENANT_ID)
            .WithRedirectUri(REDIRECT_URI)
            .WithDefaultRedirectUri()
            .Build();
    }

    /// <summary>
    /// Authenticate user with interactive browser login (like PowerShell Graph modules).
    /// Opens system browser for Azure AD authentication.
    /// </summary>
    public async Task<AuthResult> LoginAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            // Try silent authentication first (cached token)
            var accounts = await _clientApp.GetAccountsAsync();
            var firstAccount = accounts.FirstOrDefault();

            AuthenticationResult authResult;

            if (firstAccount != null)
            {
                try
                {
                    // Attempt silent token acquisition
                    authResult = await _clientApp
                        .AcquireTokenSilent(NormalizeScopes(_defaultScopes), firstAccount)
                        .ExecuteAsync(cancellationToken);
                }
                catch (MsalUiRequiredException)
                {
                    // Silent failed, need interactive login
                    authResult = await AcquireTokenInteractiveAsync(_defaultScopes, cancellationToken);
                }
            }
            else
            {
                // No cached account, interactive login required
                authResult = await AcquireTokenInteractiveAsync(_defaultScopes, cancellationToken);
            }

            // Store account and extract user info
            _currentAccount = authResult.Account;
            _currentUser = new AuthUser
            {
                UserPrincipalName = authResult.Account.Username,
                DisplayName = authResult.Account.Username.Split('@')[0], // Simplified
                Email = authResult.Account.Username,
                TenantId = authResult.TenantId
            };

            return new AuthResult
            {
                Success = true,
                User = _currentUser,
                GrantedScopes = authResult.Scopes.ToArray()
            };
        }
        catch (MsalException msalEx)
        {
            return new AuthResult
            {
                Success = false,
                Error = $"MSAL authentication failed: {msalEx.Message}\nError Code: {msalEx.ErrorCode}"
            };
        }
        catch (OperationCanceledException)
        {
            return new AuthResult
            {
                Success = false,
                Error = "Authentication cancelled by user."
            };
        }
        catch (Exception ex)
        {
            return new AuthResult
            {
                Success = false,
                Error = $"Authentication failed: {ex.Message}"
            };
        }
    }

    /// <summary>
    /// Perform interactive browser authentication.
    /// </summary>
    private async Task<AuthenticationResult> AcquireTokenInteractiveAsync(
        string[] scopes,
        CancellationToken cancellationToken)
    {
        return await _clientApp
            .AcquireTokenInteractive(NormalizeScopes(scopes))
            .WithPrompt(Prompt.SelectAccount) // Allow account selection
            .WithUseEmbeddedWebView(false) // Use system browser (better UX)
            .ExecuteAsync(cancellationToken);
    }

    /// <summary>
    /// Logout and remove cached tokens.
    /// </summary>
    public async Task LogoutAsync()
    {
        if (_currentAccount != null)
        {
            await _clientApp.RemoveAsync(_currentAccount);
            _currentAccount = null;
            _currentUser = null;
        }
    }

    /// <summary>
    /// Get access token for Microsoft Graph API calls.
    /// Automatically handles token refresh via MSAL cache.
    /// </summary>
    public async Task<string> GetAccessTokenAsync(string scope, CancellationToken cancellationToken = default)
    {
        if (!IsAuthenticated || _currentAccount == null)
            throw new InvalidOperationException("Not authenticated. Call LoginAsync first.");

        if (string.IsNullOrWhiteSpace(scope))
            throw new ArgumentException("Scope is required.", nameof(scope));

        try
        {
            // MSAL handles token caching and refresh automatically
            var result = await _clientApp
                .AcquireTokenSilent(NormalizeScopes(new[] { scope }), _currentAccount)
                .ExecuteAsync(cancellationToken);

            return result.AccessToken;
        }
        catch (MsalUiRequiredException)
        {
            // Token expired or missing consent; prompt interactively
            var result = await AcquireTokenInteractiveAsync(new[] { scope }, cancellationToken);
            return result.AccessToken;
        }
    }

    private static string[] NormalizeScopes(string[] scopes)
    {
        return scopes
            .Where(scope => !string.IsNullOrWhiteSpace(scope))
            .Select(scope =>
            {
                // If scope is already a full URL (e.g., https://outlook.office365.com/.default), don't modify it
                if (scope.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                    return scope;
                
                // If it's already prefixed with Graph scope, use as-is
                if (scope.StartsWith(GraphScopePrefix, StringComparison.OrdinalIgnoreCase))
                    return scope;
                
                // Otherwise, prepend Graph scope prefix for short-form scopes like "User.Read"
                return GraphScopePrefix + scope;
            })
            .ToArray();
    }
}

