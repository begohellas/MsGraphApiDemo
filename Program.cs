using Microsoft.Graph.Models.ODataErrors;
using MsGraphApiDemo;
using MsGraphApiDemo.GraphProviders;
using MsGraphApiDemo.Settings;
using Spectre.Console;

var loadedSettings = AppSettings.LoadSettings();
InitializeGraph(loadedSettings);

bool exit = false;
while (!exit)
{
	Console.WriteLine();
	var menuChoice = AnsiConsole.Prompt(
		new SelectionPrompt<MenuChoice>()
			.Title("\n\nPlease select an option: ")
			.AddChoices(
				MenuChoice.Exit,
				MenuChoice.DisplayMeUser,
				MenuChoice.DisplayUser,
				MenuChoice.DisplayUsers,
				MenuChoice.DisplayUsersByMail,
				MenuChoice.DisplayListGroups,
				MenuChoice.DisplayEventsCalendar));

	switch (menuChoice)
	{
		case MenuChoice.Exit:
			AnsiConsole.Write(new Rule("[yellow]Exit from the demo...[/]").RuleStyle(Style.Parse("silver")).LeftJustified());
			AnsiConsole.WriteLine();
			exit = true;
			break;

		case MenuChoice.DisplayMeUser:
			await DisplayMeUserAsync();
			AnsiConsole.Write(new Rule("[yellow] ### [/]").RuleStyle(Style.Parse("silver")).Centered());
			AnsiConsole.WriteLine();
			break;

		case MenuChoice.DisplayUser:
			await DisplayUserAsync(loadedSettings);
			AnsiConsole.Write(new Rule("[yellow] ### [/]").RuleStyle(Style.Parse("silver")).Centered());
			AnsiConsole.WriteLine();
			break;

		case MenuChoice.DisplayUsers:
			await DisplayUsersAsync(loadedSettings);
			AnsiConsole.Write(new Rule("[yellow] ### [/]").RuleStyle(Style.Parse("silver")).Centered());
			AnsiConsole.WriteLine();
			break;

		case MenuChoice.DisplayUsersByMail:
			await DisplayUsersByMailAsync(loadedSettings);
			AnsiConsole.Write(new Rule("[yellow] ### [/]").RuleStyle(Style.Parse("silver")).Centered());
			AnsiConsole.WriteLine();
			break;

		case MenuChoice.DisplayListGroups:
			await DisplayListGroupsAsync(loadedSettings);
			AnsiConsole.Write(new Rule("[yellow] ### [/]").RuleStyle(Style.Parse("silver")).Centered());
			AnsiConsole.WriteLine();
			break;

		case MenuChoice.DisplayEventsCalendar:
			AnsiConsole.Write(new Rule("[yellow] ### [/]").RuleStyle(Style.Parse("silver")).LeftJustified());
			AnsiConsole.WriteLine();
			await GraphHelper.GetEventsCalendarAsync("658b9ac7-bd21-4eaa-a045-15a5adaa2455");
			break;

		default:
			AnsiConsole.Write(new Rule("[yellow]Invalid choice. Please try again.[/]").RuleStyle(Style.Parse("silver")).LeftJustified());
			AnsiConsole.WriteLine();
			break;
	}
}

return;

void InitializeGraph(AppSettings settings)
{
	GraphHelper.InitializeGraphForUserAuth(settings);
}

async Task DisplayMeUserAsync()
{
	// will go to error because me is not supported in app-only auth
	try
	{
		var user = await GraphHelper.GetMeAsync();
		if (user is null)
		{
			AnsiConsole.MarkupLine("[red]User Me not found[/]");
			return;
		}

		var table = new Table()
			.Border(TableBorder.Rounded);
		table.AddColumn("Key");
		table.AddColumn(new TableColumn("Value").Centered());

		table.AddRow("ObjectId", $"[green]{user.Id ?? string.Empty}[/]");
		table.AddRow("DisplayName", $"[green]{user.DisplayName ?? string.Empty}[/]");
		table.AddRow("UserPrincipalName", $"[green]{user.UserPrincipalName ?? string.Empty}[/]");
		table.AddRow("Mail", $"[green]{user.Mail ?? string.Empty}[/]");
		table.AddRow("GivenName", $"[green]{user.GivenName ?? string.Empty}[/]");
		table.AddRow("Surname", $"[green]{user.Surname ?? string.Empty}[/]");
		table.AddRow("CompanyName", $"[green]{user.CompanyName ?? string.Empty}[/]");
		table.AddRow("JobTitle", $"[green]{user.JobTitle ?? string.Empty}[/]");

		AnsiConsole.Write(table);
	}
	catch (ODataError ex)
	{
		Console.ForegroundColor = ConsoleColor.Red;
		Console.WriteLine(ex.Message);
		Console.ResetColor();
	}
}

async Task DisplayUserAsync(AppSettings settings)
{
	var userId = AnsiConsole.Ask<string>($"Enter user's objectID or UPN to search in TenantID [silver]{settings.TenantId}[/]: ");
	userId ??= string.Empty;

	try
	{
		var user = await GraphHelper.GetUserAsync(userId);
		if (user is null)
		{
			AnsiConsole.MarkupLine($"[red]User with id {userId} not found[/]");
			return;
		}

		var table = new Table().RoundedBorder().Alignment(Justify.Left).BorderColor(Color.LightSlateGrey).Title($"[LightGreen]User {userId}[/]");
		table.AddColumn("Key");
		table.AddColumn(new TableColumn("Value").Centered());

		table.AddRow("ObjectId", $"[green]{user.Id ?? string.Empty}[/]");
		table.AddRow("DisplayName", $"[green]{user.DisplayName ?? string.Empty}[/]");
		table.AddRow("UserPrincipalName", $"[green]{user.UserPrincipalName ?? string.Empty}[/]");
		table.AddRow("Mail", $"[green]{user.Mail ?? string.Empty}[/]");
		table.AddRow("GivenName", $"[green]{user.GivenName ?? string.Empty}[/]");
		table.AddRow("Surname", $"[green]{user.Surname ?? string.Empty}[/]");
		table.AddRow("CompanyName", $"[green]{user.CompanyName ?? string.Empty}[/]");
		table.AddRow("JobTitle", $"[green]{user.JobTitle ?? string.Empty}[/]");

		AnsiConsole.Write(table);
	}
	catch (ODataError ex)
	{
		Console.ForegroundColor = ConsoleColor.Red;
		Console.WriteLine(ex.Message);
		Console.ResetColor();
	}
}

async Task DisplayUsersAsync(AppSettings settings)
{
	try
	{
		var users = await GraphHelper.GetUsersAsync();
		if (users is null)
		{
			AnsiConsole.MarkupLine("[red]Users not found[/]");
			return;
		}

		AnsiConsole.MarkupLine("========================================================================");
		var table = new Table().RoundedBorder().Alignment(Justify.Left).BorderColor(Color.LightSlateGrey).Title($"[LightGreen]View {GraphHelper.UsersTop} users on TenantID {settings.TenantId}[/]");
		table.AddColumn("ObjectId");
		table.AddColumn("UPN");
		table.AddColumn("DisplayName");
		table.AddColumn("Mail");

		foreach (var user in users)
		{
			table.AddRow($"[green]{user.Id ?? string.Empty}[/]", $"[green]{user.UserPrincipalName ?? string.Empty}[/]", $"[green]{user.DisplayName ?? string.Empty}[/]", $"[green]{user.Mail ?? string.Empty}[/]");
		}

		AnsiConsole.Write(table);
	}
	catch (ODataError ex)
	{
		Console.ForegroundColor = ConsoleColor.Red;
		Console.WriteLine(ex.Message);
		Console.ResetColor();
	}
}

async Task DisplayUsersByMailAsync(AppSettings settings)
{
	Console.Write($"Enter user's mail to search in TenantID {settings.TenantId}: ");
	var userMail = Console.ReadLine() ?? string.Empty;

	try
	{
		var users = await GraphHelper.GetUsersByMailAsync(userMail);
		if (users is null || users.Count == 0)
		{
			Console.ForegroundColor = ConsoleColor.Red;
			Console.WriteLine("Users not found");
			Console.ResetColor();
			return;
		}

		Console.ForegroundColor = ConsoleColor.Green;
		Console.WriteLine($"Find {users.Count} users on TenantID {settings.TenantId}");
		Console.WriteLine("========================================================================");
		foreach (var user in users)
		{
			Console.WriteLine($"ObjectId:\t {user.Id}");
			Console.WriteLine($"UPN:\t {user.UserPrincipalName}");
			Console.WriteLine($"DisplayName:\t {user.DisplayName}");
			Console.WriteLine($"Mail:\t {user.Mail}");
			Console.WriteLine("========================================================================");
		}

		Console.ResetColor();
	}
	catch (ODataError ex)
	{
		Console.ForegroundColor = ConsoleColor.Red;
		Console.WriteLine(ex.Message);
		Console.ResetColor();
	}
}

async Task DisplayListGroupsAsync(AppSettings settings)
{
	try
	{
		var userId = AnsiConsole.Ask<string>($"Enter user's objectID to search in TenantID [silver]{settings.TenantId}[/]: ");
		userId ??= string.Empty;

		var groups = await GraphHelper.ListGroupsAsync(userId);
		if (groups is null)
		{
			Console.ForegroundColor = ConsoleColor.Red;
			Console.WriteLine("Groups not found");
			Console.ResetColor();
			return;
		}

		AnsiConsole.MarkupLine("========================================================================");
		var table = new Table().RoundedBorder().Alignment(Justify.Left).BorderColor(Color.LightSlateGrey).Title("[LightGreen]Groups[/]");
		table.AddColumn("Id");
		table.AddColumn("DisplayName");
		table.AddColumn("Description");

		foreach (var group in groups)
		{
			table.AddRow($"[green]{group.Id ?? string.Empty}[/]", $"[green]{group.DisplayName ?? string.Empty}[/]", $"[green]{group.Description ?? string.Empty}[/]");
		}

		AnsiConsole.Write(table);
	}
	catch (ODataError ex)
	{
		Console.ForegroundColor = ConsoleColor.Red;
		Console.WriteLine(ex.Message);
		Console.ResetColor();
	}
}