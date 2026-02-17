using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace WinOptApp.Models;

public sealed class TweakCardItem : INotifyPropertyChanged
{
    private bool _isEnabled;
    private string _liveStateText = "Live: Unknown";

    public required string Key { get; init; }
    public required string Title { get; init; }
    public required string Description { get; init; }
    public string? WarningText { get; init; }

    public bool IsEnabled
    {
        get => _isEnabled;
        set
        {
            if (_isEnabled == value)
            {
                return;
            }

            _isEnabled = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsEnabled)));
        }
    }

    public string LiveStateText
    {
        get => _liveStateText;
        set
        {
            if (_liveStateText == value)
            {
                return;
            }

            _liveStateText = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LiveStateText)));
        }
    }

    public bool HasWarning => !string.IsNullOrWhiteSpace(WarningText);

    public event PropertyChangedEventHandler? PropertyChanged;
}
