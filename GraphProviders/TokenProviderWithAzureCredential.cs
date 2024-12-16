using Azure.Identity;
using Microsoft.Kiota.Abstractions.Authentication;

namespace MsGraphApiDemo.GraphProviders;

public class TokenProviderWithAzureCredential : IAccessTokenProvider
{
	private readonly ChainedTokenCredential _azureCredential;
	private readonly string[] _scopes;

	public TokenProviderWithAzureCredential(ChainedTokenCredential azureCredential, string[] scopes)
	{
		_azureCredential = azureCredential;
		_scopes = scopes;
	}

	public async Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object>? additionalAuthenticationContext = null, CancellationToken cancellationToken = new CancellationToken())
	{
		var (token, _) = await GetTokenAsync().ConfigureAwait(false);

		return token ?? throw new ArgumentNullException(nameof(token), "token is required");
	}

	public AllowedHostsValidator AllowedHostsValidator { get; }

	/// <summary>
	/// Acquire Token
	/// </summary>
	private async Task<(string? accessToken, DateTimeOffset expiresOn)> GetTokenAsync()
	{
		var context = new Azure.Core.TokenRequestContext(_scopes);
		var authResult = await _azureCredential.GetTokenAsync(context);

		return (authResult.Token, authResult.ExpiresOn);
	}
}