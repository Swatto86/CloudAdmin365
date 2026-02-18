namespace CloudAdmin365.Services;

/// <summary>
/// Authentication service for Azure/Office 365 services.
/// Handles token acquisition, caching, and refresh.
/// Designed to support multiple services (Exchange, Teams, SharePoint, etc).
/// </summary>
public interface IAuthService
{
    /// <summary>
    /// Login interactively or with cached tokens.
    /// </summary>
    Task<AuthResult> LoginAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Logout and clear all cached tokens.
    /// </summary>
    Task LogoutAsync();

    /// <summary>
    /// Get current authenticated user info.
    /// </summary>
    AuthUser? CurrentUser { get; }

    /// <summary>
    /// Check if authenticated.
    /// </summary>
    bool IsAuthenticated { get; }

    /// <summary>
    /// Get access token for a specific service scope.
    /// </summary>
    Task<string> GetAccessTokenAsync(string scope, CancellationToken cancellationToken = default);
}

/// <summary>
/// Result of authentication attempt.
/// </summary>
public class AuthResult
{
    /// <summary>
    /// True if login succeeded.
    /// </summary>
    public required bool Success { get; set; }

    /// <summary>
    /// Error message if login failed.
    /// </summary>
    public string? Error { get; set; }

    /// <summary>
    /// Authenticated user info.
    /// </summary>
    public AuthUser? User { get; set; }

    /// <summary>
    /// Scopes that were successfully authorized.
    /// </summary>
    public string[] GrantedScopes { get; set; } = [];
}

/// <summary>
/// Authenticated user information.
/// </summary>
public class AuthUser
{
    /// <summary>
    /// User principal name (e.g., user@company.onmicrosoft.com).
    /// </summary>
    public required string UserPrincipalName { get; set; }

    /// <summary>
    /// Display name.
    /// </summary>
    public required string DisplayName { get; set; }

    /// <summary>
    /// Email address.
    /// </summary>
    public required string Email { get; set; }

    /// <summary>
    /// Tenant ID.
    /// </summary>
    public string? TenantId { get; set; }
}
