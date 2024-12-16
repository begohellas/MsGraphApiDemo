using Spectre.Console;
using System.Runtime.CompilerServices;

namespace MsGraphApiDemo;
internal static class InitializeConsole
{
	[ModuleInitializer]
	public static void Init()
	{
		Console.Title = "Examples Ms Graph Api";

		AnsiConsole.Write(
			new FigletText("Ms Graph DEMO")
				.LeftJustified()
				.Color(Color.DarkOrange));

		AnsiConsole.MarkupLine("");
	}
}
