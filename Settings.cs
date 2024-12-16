using Microsoft.Extensions.Configuration;

namespace MsGraphApiDemo;
public record Settings
{
	public const string SectionName = "Settings";

	public string ClientId { get; init; } = null!;
	public string TenantId { get; init; } = null!;
	public string ClientSecret { get; init; } = null!;
	public string[]? GraphUserScopes { get; init; }

	public static Settings LoadSettings()
	{
		IConfiguration config = new ConfigurationBuilder()
			.AddJsonFile("appsettings.json", optional: false)
			.AddUserSecrets<Program>() // Read clientSecret from usersecrets
			.Build();

		Settings? settings = config.GetRequiredSection(SectionName).Get<Settings>();
		return settings ?? throw new ArgumentNullException(nameof(config));
	}
}
