using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;

namespace MsGraphApiDemo.GraphProviders;

public class TokenProviderWithConfidentialApp : IAccessTokenProvider
{
	private readonly IConfidentialClientApplication _clientApplication;
	private readonly string[] _scopes;

	public TokenProviderWithConfidentialApp(IConfidentialClientApplication clientApplication, string[] scopes)
	{
		_clientApplication = clientApplication;
		_scopes = scopes;
	}

	public async Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object>? additionalAuthenticationContext = null, CancellationToken cancellationToken = new())
	{
		var (token, _) = await GetTokenAsync().ConfigureAwait(false);

		return token ?? throw new ArgumentException("token is required for graph client", nameof(token));
	}

	public AllowedHostsValidator AllowedHostsValidator { get; }

	/// <summary>
	/// Acquire Token
	/// </summary>
	private async Task<(string? accessToken, DateTimeOffset expiresOn)> GetTokenAsync()
	{
		var authResult = await _clientApplication.AcquireTokenForClient(_scopes).ExecuteAsync();

		return (authResult.AccessToken, authResult.ExpiresOn);
	}
}