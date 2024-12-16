namespace MsGraphApiDemo.Models;

public sealed record GroupDetails
{
	public string Id { get; init; } = default!;

	public string DisplayName { get; init; } = default!;

	public string Description { get; init; } = default!;
}