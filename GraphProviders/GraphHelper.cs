﻿using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;
using MsGraphApiDemo.Models;
using MsGraphApiDemo.Settings;
using Spectre.Console;

namespace MsGraphApiDemo.GraphProviders;
internal static class GraphHelper
{
	internal const int UsersTop = 20;
	private static AppSettings? _settings;

	// Client configured with user authentication
	private static GraphServiceClient? _graphClient;

	// User auth token provider
	private static IAccessTokenProvider? _tokenProvider;

	public static string AccessToken { get; private set; } = string.Empty;
	public static DateTimeOffset AccessTokenExpiresOn { get; private set; } = DateTimeOffset.MinValue;

	public static void InitializeGraphForUserAuth(AppSettings appSettings)
	{
		_settings = appSettings;

		// initialiaze graph client with client secret auth
		InitializeClientSecretAuth(appSettings);

		// initialiaze graph client with token provider and IAuthenticationProvider
		//InitializeTokenProvider(settings);

		AnsiConsole.MarkupLine("### Token that is used in the GraphServiceClient ###");
		AnsiConsole.WriteLine();
		AnsiConsole.MarkupLine($"[green]{_tokenProvider?.GetAuthorizationTokenAsync(null!).GetAwaiter().GetResult()}[/]");
		AnsiConsole.Write(new Rule("[yellow] ### [/]").RuleStyle(Style.Parse("silver")).Centered());
		AnsiConsole.WriteLine();

		var authenticationProvider = new BaseBearerTokenAuthenticationProvider(_tokenProvider!);
		_graphClient = new GraphServiceClient(authenticationProvider);
	}

	/// <summary>
	/// Retrieves the current user from the Microsoft Graph API.
	/// </summary>
	/// <returns>The user object representing the current user.</returns>
	public static async Task<User?> GetMeAsync()
	{
		_ = _graphClient ?? throw new ArgumentNullException(nameof(_graphClient), "Graph has not been initialized for user auth");

		User? result = await _graphClient.Me.GetAsync();

		return result;
	}

	/// <summary>
	/// Retrieves a user from the Microsoft Graph API based on the user ID.
	/// </summary>
	/// <param name="userId">The ID of the user to retrieve.</param>
	/// <returns>The user object.</returns>
	public static async Task<User?> GetUserAsync(string userId)
	{
		_ = _graphClient ?? throw new ArgumentNullException(nameof(_graphClient), "Graph has not been initialized for user auth");

		User? result = await _graphClient.Users[userId].GetAsync();

		return result;
	}

	/// <summary>
	/// Retrieves a list of users from the Microsoft Graph API using query parameters option.
	/// </summary>
	/// <returns>The list of users.</returns>
	public static async Task<List<User>?> GetUsersAsync()
	{
		_ = _graphClient ?? throw new ArgumentNullException(nameof(_graphClient), "Graph has not been initialized for user auth");

		var usersResponse = await _graphClient.Users.GetAsync(requestConfiguration =>
		{
			requestConfiguration.QueryParameters.Select = ["id", "displayName", "userPrincipalName", "mail"];
			requestConfiguration.QueryParameters.Top = UsersTop;
			requestConfiguration.QueryParameters.Orderby = ["userPrincipalName"];
			requestConfiguration.QueryParameters.Count = true;
			requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
		});

		return usersResponse?.Value;
	}

	/// <summary>
	/// Retrieves a list of users from the Microsoft Graph API based on the user's email.
	/// </summary>
	/// <param name="userMail">The email of the user to retrieve.</param>
	/// <exception cref="ArgumentNullException"></exception>
	/// <returns>The list of users.</returns>
	public static async Task<List<User>?> GetUsersByMailAsync(string userMail)
	{
		_ = _graphClient ?? throw new ArgumentNullException(nameof(_graphClient), "Graph has not been initialized for user auth");

		var result = await _graphClient.Users.GetAsync(requestConfiguration =>
		{
			requestConfiguration.QueryParameters.Filter = $"mail eq '{userMail}'";
			requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
		});

		return result is null
			? Array.Empty<User>().ToList()
			: result.Value;
	}

	public static async Task<List<GroupDetails>?> ListGroupsAsync(string userId)
	{
		_ = _graphClient ?? throw new ArgumentNullException(nameof(_graphClient), "Graph has not been initialized for user auth");

		List<GroupDetails> groupIds = new();

		GroupCollectionResponse? response = await _graphClient.Users[userId].MemberOf.GraphGroup.GetAsync(requestConfiguration =>
		{
			requestConfiguration.QueryParameters.Select = ["id", "displayName", "description"];
			requestConfiguration.QueryParameters.Top = 100;
		});

		var pageIterator = PageIterator<Group, GroupCollectionResponse?>
			.CreatePageIterator(_graphClient, response, (group) =>
		{
			groupIds.Add(new GroupDetails() { Id = group.Id!, DisplayName = group.DisplayName!, Description = group.Description! });

			return true;
		});

		await pageIterator.IterateAsync();

		return groupIds;
	}

	public static async Task GetEventsCalendarAsync(string userId)
	{
		_ = _graphClient ?? throw new ArgumentNullException(nameof(_graphClient), "Graph has not been initialized for user auth");

		var result = await _graphClient!.Users[userId].Events.GetAsync((requestConfiguration) =>
		{
			requestConfiguration.QueryParameters.Select = ["subject", "body", "bodyPreview", "organizer", "attendees", "start", "end", "location"];
			requestConfiguration.QueryParameters.Top = 10;
		});

		result?.Value?.ForEach(x => Console.WriteLine(x.Subject));
	}

	private static void InitializeClientSecretAuth(AppSettings appSettings)
	{
		ArgumentNullException.ThrowIfNull(appSettings);

		// Validate required settings
		ValidateSettings(appSettings);

		ClientSecretCredentialOptions clientSecretCredentialOptions = new()
		{
			AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
			Diagnostics =
			{
				LoggedHeaderNames = { "x-ms-request-id" },
				LoggedQueryParameters = { "api-version" },
				IsAccountIdentifierLoggingEnabled = true,
			}
		};
		var clientSecretCredential = new ClientSecretCredential(appSettings.TenantId, appSettings.ClientId, appSettings.ClientSecret, clientSecretCredentialOptions);
		var chainedCredential = new ChainedTokenCredential(clientSecretCredential);

		_tokenProvider = new TokenProviderWithAzureCredential(chainedCredential, appSettings.GraphUserScopes!);
	}

	private static void InitializeTokenProvider(AppSettings appSettings)
	{
		// Validate required settings
		ValidateSettings(appSettings);

		var confidentialClientApplication = ConfidentialClientApplicationBuilder
						.Create(appSettings.ClientId)
						.WithClientSecret(appSettings.ClientSecret)
						.WithTenantId(appSettings.TenantId)
						.Build();

		_tokenProvider = new TokenProviderWithConfidentialApp(confidentialClientApplication, appSettings.GraphUserScopes!);
	}

	private static void ValidateSettings(AppSettings appSettings)
	{
		if (string.IsNullOrEmpty(appSettings.TenantId) ||
			string.IsNullOrEmpty(appSettings.ClientId) ||
			string.IsNullOrEmpty(appSettings.ClientSecret) ||
			appSettings.GraphUserScopes is null ||
			appSettings.GraphUserScopes.Length == 0)
		{
			throw new ArgumentException("Invalid settings. Please check the appsettings.json file.", nameof(appSettings));
		}
	}
}