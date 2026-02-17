namespace WinOptApp.Models;

public sealed class OptimizationState
{
    public DateTime SavedAtUtc { get; set; }
    public string? PreviousPowerSchemeGuid { get; set; }
    public int? PreviousAutoGameModeEnabled { get; set; }
}
