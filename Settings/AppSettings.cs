using Microsoft.Extensions.Configuration;

namespace MsGraphApiDemo.Settings;
public record AppSettings
{
	public const string SectionName = "Settings";

	public string ClientId { get; init; } = null!;
	public string TenantId { get; init; } = null!;
	public string ClientSecret { get; init; } = null!;
	public string[]? GraphUserScopes { get; init; }

	public static AppSettings LoadSettings()
	{
		IConfiguration config = new ConfigurationBuilder()
			.AddJsonFile("appsettings.json", optional: false)
			.AddUserSecrets<Program>() // Read clientSecret from usersecrets
			.Build();

		AppSettings? settings = config.GetRequiredSection(SectionName).Get<AppSettings>();
		return settings ?? throw new ArgumentNullException(nameof(config));
	}
}
