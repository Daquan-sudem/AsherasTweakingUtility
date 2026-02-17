using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Management;
using System.Net.NetworkInformation;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;
using Microsoft.Win32;
using WinOptApp.Models;
using WinOptApp.Services;

namespace WinOptApp;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private readonly OptimizerService _optimizer = new();
    private readonly DispatcherTimer _telemetryTimer = new() { Interval = TimeSpan.FromMilliseconds(1400) };
    private readonly PerformanceCounter? _cpuCounter;
    private readonly PerformanceCounter? _cpuFrequencyCounter;
    private readonly PerformanceCounter? _cpuFrequencyPercentCounter;
    private readonly string _systemDriveRoot;
    private readonly double _cpuMaxMhz;
    private readonly string _cpuName;
    private readonly string _gpuName;

    private readonly ObservableCollection<TweakCardItem> _tweakCards = [];

    private string _statusText = "Ready";
    private double _cpuUsage;
    private double _memoryUsage;
    private double _diskUsage;
    private double _cpuSpeedMhz;
    private string _cpuUsageText = "0%";
    private string _cpuSpeedText = "CPU speed: N/A";
    private string _memoryUsageText = "0%";
    private string _memoryDetailsText = "0.0 / 0.0 GB";
    private string _diskUsageText = "0%";
    private string _diskDetailsText = "C: 0.0 / 0.0 GB";
    private string _diskUsedText = "Used: 0.0 GB";
    private string _diskFreeText = "Free: 0.0 GB";
    private double _gpuUsage;
    private string _gpuUsageText = "N/A";
    private string _gpuDetailsText = "Usage: N/A";
    private string _gpuVramText = "VRAM: N/A";
    private string _networkUploadText = "Upload: 0.00 Mbps";
    private string _networkDownloadText = "Download: 0.00 Mbps";
    private string _networkAdapterText = "Adapter: N/A";
    private string _networkTotalsText = "Sent/Recv: 0.00 / 0.00 GB";
    private string _temperatureValueText = "N/A";
    private string _temperatureDetailText = "Source: CPU | Sensor unavailable";
    private string _temperatureTrendText = "Trend: --";
    private string _controllerPollingStatusText = "Select a controller/USB device, choose polling rate, then apply.";
    private string _fortniteRegionStatusText = "Fortnite region: waiting for config...";
    private string _systemSpecsText = "Loading specs...";
    private bool _isDashboardVisible = true;
    private bool _isGeneralTweaksVisible;
    private bool _isGameProfilesVisible;
    private DateTime _lastTweakStateRefreshUtc = DateTime.MinValue;
    private DateTime _lastGpuRefreshUtc = DateTime.MinValue;
    private DateTime _lastTempRefreshUtc = DateTime.MinValue;
    private DateTime _lastBackgroundTelemetryUtc = DateTime.MinValue;
    private long _lastNetworkBytesSent;
    private long _lastNetworkBytesReceived;
    private DateTime _lastNetworkSampleUtc = DateTime.MinValue;
    private string _tempSource = "CPU";
    private readonly List<double> _tempHistory = [];
    private readonly ObservableCollection<ControllerDeviceItem> _controllerDevices = [];
    private readonly ObservableCollection<FortniteRegionItem> _fortniteRegions = [];
    private readonly ObservableCollection<FortniteServerPingItem> _fortniteServerPings = [];
    private double? _latestCpuTempC;
    private double? _latestGpuTempC;
    private bool _isRefreshingTweakStates;
    private bool _suppressToggleEvents;
    private bool _isPingingFortniteServers;
    private DateTime _lastFortnitePingUtc = DateTime.MinValue;
    private DateTime _lastCoreTempProbeUtc = DateTime.MinValue;
    private bool _coreTempAvailable;
    private float _coreTempCpuSpeedMhz;
    private float _coreTempPackageTempC;

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;
        _systemDriveRoot = Path.GetPathRoot(Environment.SystemDirectory) ?? "C:\\";

        try
        {
            _cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
            _ = _cpuCounter.NextValue();
        }
        catch
        {
            _cpuCounter = null;
        }

        try
        {
            _cpuFrequencyCounter = new PerformanceCounter("Processor Information", "Processor Frequency", "_Total");
            _ = _cpuFrequencyCounter.NextValue();
        }
        catch
        {
            _cpuFrequencyCounter = null;
        }

        try
        {
            _cpuFrequencyPercentCounter = new PerformanceCounter("Processor Information", "% of Maximum Frequency", "_Total");
            _ = _cpuFrequencyPercentCounter.NextValue();
        }
        catch
        {
            _cpuFrequencyPercentCounter = null;
        }

        (_cpuName, _cpuMaxMhz, _gpuName) = LoadCpuAndGpuSpecs();
        GpuVramText = ReadGpuVramText();
        _systemSpecsText = BuildSystemSpecsText();
        SeedTweakCards();

        _telemetryTimer.Tick += (_, _) => RefreshTelemetry();
        Loaded += (_, _) =>
        {
            RefreshTelemetry();
            UpdateTempSourceButtons();
            LoadControllerDevices();
            LoadFortniteRegionOptions();
            ReloadFortniteRegionState();
            _telemetryTimer.Start();
            _ = RefreshTweakStatesAsync(showPopup: false);
        };
        Closing += MainWindow_Closing;

        OutputTextBox.Text = "Ashera's Tweaking Utility ready. Click Analyze to inspect current settings.";
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public string StatusText
    {
        get => _statusText;
        set
        {
            if (_statusText == value)
            {
                return;
            }

            _statusText = value;
            OnPropertyChanged();
        }
    }

    public ObservableCollection<TweakCardItem> TweakCards => _tweakCards;

    public double CpuUsage
    {
        get => _cpuUsage;
        private set
        {
            if (Math.Abs(_cpuUsage - value) < 0.1)
            {
                return;
            }

            _cpuUsage = value;
            OnPropertyChanged();
        }
    }

    public double MemoryUsage
    {
        get => _memoryUsage;
        private set
        {
            if (Math.Abs(_memoryUsage - value) < 0.1)
            {
                return;
            }

            _memoryUsage = value;
            OnPropertyChanged();
        }
    }

    public double DiskUsage
    {
        get => _diskUsage;
        private set
        {
            if (Math.Abs(_diskUsage - value) < 0.1)
            {
                return;
            }

            _diskUsage = value;
            OnPropertyChanged();
        }
    }

    public string CpuUsageText
    {
        get => _cpuUsageText;
        private set
        {
            if (_cpuUsageText == value)
            {
                return;
            }

            _cpuUsageText = value;
            OnPropertyChanged();
        }
    }

    public double CpuSpeedMhz
    {
        get => _cpuSpeedMhz;
        private set
        {
            if (Math.Abs(_cpuSpeedMhz - value) < 1)
            {
                return;
            }

            _cpuSpeedMhz = value;
            OnPropertyChanged();
        }
    }

    public double CpuSpeedMaxMhz => _cpuMaxMhz > 0 ? _cpuMaxMhz : 6000;

    public string CpuSpeedText
    {
        get => _cpuSpeedText;
        private set
        {
            if (_cpuSpeedText == value)
            {
                return;
            }

            _cpuSpeedText = value;
            OnPropertyChanged();
        }
    }

    public string MemoryUsageText
    {
        get => _memoryUsageText;
        private set
        {
            if (_memoryUsageText == value)
            {
                return;
            }

            _memoryUsageText = value;
            OnPropertyChanged();
        }
    }

    public string MemoryDetailsText
    {
        get => _memoryDetailsText;
        private set
        {
            if (_memoryDetailsText == value)
            {
                return;
            }

            _memoryDetailsText = value;
            OnPropertyChanged();
        }
    }

    public string DiskUsageText
    {
        get => _diskUsageText;
        private set
        {
            if (_diskUsageText == value)
            {
                return;
            }

            _diskUsageText = value;
            OnPropertyChanged();
        }
    }

    public string DiskDetailsText
    {
        get => _diskDetailsText;
        private set
        {
            if (_diskDetailsText == value)
            {
                return;
            }

            _diskDetailsText = value;
            OnPropertyChanged();
        }
    }

    public string DiskUsedText
    {
        get => _diskUsedText;
        private set
        {
            if (_diskUsedText == value)
            {
                return;
            }

            _diskUsedText = value;
            OnPropertyChanged();
        }
    }

    public string DiskFreeText
    {
        get => _diskFreeText;
        private set
        {
            if (_diskFreeText == value)
            {
                return;
            }

            _diskFreeText = value;
            OnPropertyChanged();
        }
    }

    public double GpuUsage
    {
        get => _gpuUsage;
        private set
        {
            if (Math.Abs(_gpuUsage - value) < 0.1)
            {
                return;
            }

            _gpuUsage = value;
            OnPropertyChanged();
        }
    }

    public string GpuUsageText
    {
        get => _gpuUsageText;
        private set
        {
            if (_gpuUsageText == value)
            {
                return;
            }

            _gpuUsageText = value;
            OnPropertyChanged();
        }
    }

    public string GpuDetailsText
    {
        get => _gpuDetailsText;
        private set
        {
            if (_gpuDetailsText == value)
            {
                return;
            }

            _gpuDetailsText = value;
            OnPropertyChanged();
        }
    }

    public string GpuVramText
    {
        get => _gpuVramText;
        private set
        {
            if (_gpuVramText == value)
            {
                return;
            }

            _gpuVramText = value;
            OnPropertyChanged();
        }
    }

    public string NetworkUploadText
    {
        get => _networkUploadText;
        private set
        {
            if (_networkUploadText == value)
            {
                return;
            }

            _networkUploadText = value;
            OnPropertyChanged();
        }
    }

    public string NetworkDownloadText
    {
        get => _networkDownloadText;
        private set
        {
            if (_networkDownloadText == value)
            {
                return;
            }

            _networkDownloadText = value;
            OnPropertyChanged();
        }
    }

    public string NetworkAdapterText
    {
        get => _networkAdapterText;
        private set
        {
            if (_networkAdapterText == value)
            {
                return;
            }

            _networkAdapterText = value;
            OnPropertyChanged();
        }
    }

    public string NetworkTotalsText
    {
        get => _networkTotalsText;
        private set
        {
            if (_networkTotalsText == value)
            {
                return;
            }

            _networkTotalsText = value;
            OnPropertyChanged();
        }
    }

    public string TemperatureValueText
    {
        get => _temperatureValueText;
        private set
        {
            if (_temperatureValueText == value)
            {
                return;
            }

            _temperatureValueText = value;
            OnPropertyChanged();
        }
    }

    public string TemperatureDetailText
    {
        get => _temperatureDetailText;
        private set
        {
            if (_temperatureDetailText == value)
            {
                return;
            }

            _temperatureDetailText = value;
            OnPropertyChanged();
        }
    }

    public string TemperatureTrendText
    {
        get => _temperatureTrendText;
        private set
        {
            if (_temperatureTrendText == value)
            {
                return;
            }

            _temperatureTrendText = value;
            OnPropertyChanged();
        }
    }

    public string ControllerPollingStatusText
    {
        get => _controllerPollingStatusText;
        private set
        {
            if (_controllerPollingStatusText == value)
            {
                return;
            }

            _controllerPollingStatusText = value;
            OnPropertyChanged();
        }
    }

    public string FortniteRegionStatusText
    {
        get => _fortniteRegionStatusText;
        private set
        {
            if (_fortniteRegionStatusText == value)
            {
                return;
            }

            _fortniteRegionStatusText = value;
            OnPropertyChanged();
        }
    }

    public ObservableCollection<FortniteServerPingItem> FortniteServerPings => _fortniteServerPings;

    public string SystemSpecsText
    {
        get => _systemSpecsText;
        private set
        {
            if (_systemSpecsText == value)
            {
                return;
            }

            _systemSpecsText = value;
            OnPropertyChanged();
        }
    }

    public bool IsDashboardVisible
    {
        get => _isDashboardVisible;
        private set
        {
            if (_isDashboardVisible == value)
            {
                return;
            }

            _isDashboardVisible = value;
            OnPropertyChanged();
        }
    }

    public bool IsGeneralTweaksVisible
    {
        get => _isGeneralTweaksVisible;
        private set
        {
            if (_isGeneralTweaksVisible == value)
            {
                return;
            }

            _isGeneralTweaksVisible = value;
            OnPropertyChanged();
        }
    }

    public bool IsGameProfilesVisible
    {
        get => _isGameProfilesVisible;
        private set
        {
            if (_isGameProfilesVisible == value)
            {
                return;
            }

            _isGameProfilesVisible = value;
            OnPropertyChanged();
        }
    }

    private async void AnalyzeButton_Click(object sender, RoutedEventArgs e)
    {
        await ExecuteActionAsync("Analyzing...", _optimizer.AnalyzeAsync, "Analysis complete", "Analyze failed");
    }

    private async void ApplyProfileButton_Click(object sender, RoutedEventArgs e)
    {
        await ExecuteActionAsync("Applying profile...", _optimizer.ApplyGamingProfileAsync, "Gaming profile applied", "Apply failed");
    }

    private async void RollbackButton_Click(object sender, RoutedEventArgs e)
    {
        await ExecuteActionAsync("Rolling back...", _optimizer.RollbackAsync, "Rollback complete", "Rollback failed");
    }

    private async void CheckUpdatesButton_Click(object sender, RoutedEventArgs e)
    {
        await ExecuteActionAsync("Checking updates...", _optimizer.CheckForUpdateAsync, "Update check complete", "Update check failed");
    }

    private async void DownloadUpdateButton_Click(object sender, RoutedEventArgs e)
    {
        SetBusy("Preparing in-place update...");
        try
        {
            var result = await _optimizer.DownloadLatestUpdateAsync();
            OutputTextBox.Text = result;

            if (!result.Contains("STATUS: READY_FOR_RESTART", StringComparison.Ordinal))
            {
                StatusText = "Update download complete";
                return;
            }

            var updaterPath = ExtractUpdaterPath(result);
            if (string.IsNullOrWhiteSpace(updaterPath) || !File.Exists(updaterPath))
            {
                StatusText = "Update prepared, but updater script missing";
                return;
            }

            var choice = MessageBox.Show(
                "Update is ready. The app will close, update itself, and reopen. Continue now?",
                "Install Update",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (choice != MessageBoxResult.Yes)
            {
                StatusText = "Update ready (pending install)";
                return;
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = updaterPath,
                UseShellExecute = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            });
            Application.Current.Shutdown();
        }
        catch (Exception ex)
        {
            OutputTextBox.Text = $"Update download failed:\n{ex.Message}";
            StatusText = "Update download failed";
        }
    }

    private void HomeButton_Click(object sender, RoutedEventArgs e)
    {
        SetDashboardView();
        RefreshTelemetry();
        OutputTextBox.Text =
            "Ashera's Tweaking Utility\n" +
            "=========================\n" +
            $"CPU: {CpuUsageText}\n" +
            $"GPU: {GpuUsageText}\n" +
            $"Memory: {MemoryUsageText} ({MemoryDetailsText})\n" +
            $"Disk: {DiskUsageText} ({DiskDetailsText})\n" +
            $"{NetworkUploadText} | {NetworkDownloadText}\n" +
            $"{TemperatureDetailText}\n" +
            $"\nUpdated: {DateTime.Now}";
        StatusText = "Home";
    }

    private void GeneralTweaksButton_Click(object sender, RoutedEventArgs e)
    {
        SetGeneralTweaksView();
        StatusText = "General Tweaks";
        OutputTextBox.Text = "General Tweaks view active. Toggle cards to apply or revert each tweak.";
        _ = RefreshTweakStatesAsync(showPopup: true);
    }

    private void GameProfilesButton_Click(object sender, RoutedEventArgs e)
    {
        SetGameProfilesView();
        ReloadFortniteRegionState();
        LoadFortniteServerTargets();
        _ = PingFortniteServersAsync(force: false);
        StatusText = "Game Profiles";
        OutputTextBox.Text = "Game Profiles view active. Configure per-game settings below.";
    }

    private async void SystemRestoreButton_Click(object sender, RoutedEventArgs e)
    {
        await ExecuteActionAsync("Creating restore point...", _optimizer.CreateRestorePointAsync, "Restore point task finished", "Restore point failed");
    }

    private async void ResourcesButton_Click(object sender, RoutedEventArgs e)
    {
        await ExecuteActionAsync("Collecting resource data...", _optimizer.GetResourceReportAsync, "Resource report ready", "Resource report failed");
    }

    private async void FixesButton_Click(object sender, RoutedEventArgs e)
    {
        await ExecuteActionAsync("Running quick fixes...", _optimizer.RunQuickFixesAsync, "Quick fixes completed", "Quick fixes failed");
    }

    private async void GamingProfileButton_Click(object sender, RoutedEventArgs e)
    {
        await ExecuteActionAsync("Applying profile...", _optimizer.ApplyGamingProfileAsync, "Gaming profile applied", "Apply failed");
    }

    private async void StartupOptimizerButton_Click(object sender, RoutedEventArgs e)
    {
        await ExecuteActionAsync("Reviewing startup apps...", _optimizer.GetStartupOptimizationReportAsync, "Startup report ready", "Startup report failed");
    }

    private async void AppOptimizerButton_Click(object sender, RoutedEventArgs e)
    {
        await ExecuteActionAsync("Reviewing active apps...", _optimizer.GetAppOptimizationReportAsync, "App report ready", "App report failed");
    }

    private async void TweakToggle_Changed(object sender, RoutedEventArgs e)
    {
        if (_suppressToggleEvents)
        {
            return;
        }

        if (sender is not ToggleButton { Tag: TweakCardItem item })
        {
            return;
        }

        var enabled = item.IsEnabled;
        await ExecuteActionAsync(
            $"{(enabled ? "Enabling" : "Disabling")} {item.Title}...",
            () => _optimizer.ApplyTweakToggleAsync(item.Key, enabled),
            $"{item.Title} {(enabled ? "enabled" : "disabled")}",
            $"{item.Title} toggle failed");

        if (IsRestartRequiredTweak(item.Key))
        {
            var msg = "Restart required for this tweak to fully apply.";
            StatusText = $"{item.Title} updated (restart required)";
            OutputTextBox.Text = $"{OutputTextBox.Text}\n\n{msg}";
            MessageBox.Show(msg, "Restart Required", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        _ = RefreshTweakStatesAsync(showPopup: false);
    }

    private void SetBusy(string message)
    {
        StatusText = message;
        OutputTextBox.Text = message;
    }

    private async Task ExecuteActionAsync(
        string busyMessage,
        Func<Task<string>> action,
        string successStatus,
        string errorStatus)
    {
        SetBusy(busyMessage);

        try
        {
            OutputTextBox.Text = await action();
            StatusText = successStatus;
        }
        catch (Exception ex)
        {
            OutputTextBox.Text = $"{errorStatus}:\n{ex.Message}";
            StatusText = errorStatus;
        }
    }

    private void RefreshTelemetry()
    {
        // Keep background sections responsive by skipping heavy WMI/network work when dashboard is hidden.
        if (!IsDashboardVisible)
        {
            if ((DateTime.UtcNow - _lastBackgroundTelemetryUtc).TotalSeconds >= 5)
            {
                CpuUsage = ClampPercent(ReadCpuUsage());
                CpuUsageText = $"{CpuUsage:0}%";
                _lastBackgroundTelemetryUtc = DateTime.UtcNow;
            }

            if ((DateTime.UtcNow - _lastTweakStateRefreshUtc).TotalSeconds >= 6)
            {
                _ = RefreshTweakStatesAsync(showPopup: false);
            }

            return;
        }

        CpuUsage = ClampPercent(ReadCpuUsage());
        CpuUsageText = $"{CpuUsage:0}%";

        var speedMhz = ReadCpuSpeedMhz();
        if (speedMhz > 0)
        {
            CpuSpeedMhz = Math.Min(CpuSpeedMaxMhz, speedMhz);
            CpuSpeedText = $"CPU speed: {speedMhz:0} MHz ({speedMhz / 1000.0:0.00} GHz)";
        }
        else
        {
            CpuSpeedMhz = 0;
            CpuSpeedText = "CPU speed: unavailable";
        }

        var memoryStatus = new MemoryStatusEx();
        if (GlobalMemoryStatusEx(memoryStatus))
        {
            var totalMem = (long)memoryStatus.ullTotalPhys;
            var availMem = (long)memoryStatus.ullAvailPhys;
            var usedMem = totalMem - availMem;
            var memPercent = totalMem == 0 ? 0 : (double)usedMem / totalMem * 100.0;

            MemoryUsage = ClampPercent(memPercent);
            MemoryUsageText = $"{MemoryUsage:0}%";
            MemoryDetailsText = $"{ToGiB(usedMem):0.0} / {ToGiB(totalMem):0.0} GB";
            SystemSpecsText = BuildSystemSpecsText((ulong)totalMem);
        }

        try
        {
            var drive = new DriveInfo(_systemDriveRoot);
            if (drive.IsReady)
            {
                var usedDisk = drive.TotalSize - drive.TotalFreeSpace;
                var diskPercent = drive.TotalSize == 0 ? 0 : (double)usedDisk / drive.TotalSize * 100.0;
                DiskUsage = ClampPercent(diskPercent);
                DiskUsageText = $"{DiskUsage:0}%";
                DiskDetailsText = $"{drive.Name} {ToGiB(usedDisk):0.0} / {ToGiB(drive.TotalSize):0.0} GB";
                DiskUsedText = $"Used: {ToGiB(usedDisk):0.0} GB";
                DiskFreeText = $"Free: {ToGiB((long)drive.TotalFreeSpace):0.0} GB";
            }
        }
        catch
        {
            DiskUsage = 0;
            DiskUsageText = "N/A";
            DiskDetailsText = "Drive metrics unavailable";
            DiskUsedText = "Used: N/A";
            DiskFreeText = "Free: N/A";
        }

        RefreshNetworkTelemetry();
        RefreshGpuTelemetry();
        RefreshTemperatureTelemetry();

        if ((DateTime.UtcNow - _lastTweakStateRefreshUtc).TotalSeconds >= 6)
        {
            _ = RefreshTweakStatesAsync(showPopup: false);
        }
    }

    private double ReadCpuUsage()
    {
        if (_cpuCounter is null)
        {
            return 0;
        }

        try
        {
            return _cpuCounter.NextValue();
        }
        catch
        {
            return 0;
        }
    }

    private double ReadCpuSpeedMhz()
    {
        if (TryReadCoreTempSharedMemory())
        {
            if (_coreTempCpuSpeedMhz > 0)
            {
                return _coreTempCpuSpeedMhz;
            }
        }

        try
        {
            using var currentClockSearcher = new ManagementObjectSearcher("SELECT CurrentClockSpeed FROM Win32_Processor");
            foreach (var obj in currentClockSearcher.Get().OfType<ManagementObject>())
            {
                if (double.TryParse(obj["CurrentClockSpeed"]?.ToString(), out var currentClock) && currentClock > 0)
                {
                    return currentClock;
                }
            }

            if (_cpuFrequencyPercentCounter is not null && _cpuMaxMhz > 0)
            {
                var percentOfMax = _cpuFrequencyPercentCounter.NextValue();
                if (percentOfMax > 0)
                {
                    return (_cpuMaxMhz * percentOfMax) / 100.0;
                }
            }

            if (_cpuFrequencyCounter is not null)
            {
                var direct = _cpuFrequencyCounter.NextValue();
                if (direct > 0)
                {
                    return direct;
                }
            }
        }
        catch
        {
            // Ignore and fallback below.
        }

        return 0;
    }

    private (string CpuName, double MaxMhz, string GpuName) LoadCpuAndGpuSpecs()
    {
        var cpuName = "Unknown CPU";
        var maxMhz = 0d;
        var gpuName = "Unknown GPU";

        try
        {
            using var cpuSearcher = new ManagementObjectSearcher("SELECT Name, MaxClockSpeed FROM Win32_Processor");
            foreach (var obj in cpuSearcher.Get().OfType<ManagementObject>())
            {
                cpuName = obj["Name"]?.ToString()?.Trim() ?? cpuName;
                if (double.TryParse(obj["MaxClockSpeed"]?.ToString(), out var parsed))
                {
                    maxMhz = parsed;
                }

                break;
            }
        }
        catch
        {
            // Keep defaults.
        }

        try
        {
            using var gpuSearcher = new ManagementObjectSearcher("SELECT Name FROM Win32_VideoController");
            foreach (var obj in gpuSearcher.Get().OfType<ManagementObject>())
            {
                var name = obj["Name"]?.ToString()?.Trim();
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                if (!name.Contains("Microsoft Basic", StringComparison.OrdinalIgnoreCase))
                {
                    gpuName = name;
                    break;
                }

                gpuName = name;
            }
        }
        catch
        {
            // Keep default.
        }

        return (cpuName, maxMhz, gpuName);
    }

    private string ReadGpuVramText()
    {
        try
        {
            using var gpuSearcher = new ManagementObjectSearcher("SELECT AdapterRAM, Name FROM Win32_VideoController");
            foreach (var obj in gpuSearcher.Get().OfType<ManagementObject>())
            {
                var name = obj["Name"]?.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                if (!name.Equals(_gpuName, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(_gpuName, "Unknown GPU", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (double.TryParse(obj["AdapterRAM"]?.ToString(), out var bytes) && bytes > 0)
                {
                    return $"VRAM: {bytes / 1024d / 1024d / 1024d:0.0} GB";
                }
            }
        }
        catch
        {
            // ignored
        }

        return "VRAM: N/A";
    }

    private string BuildSystemSpecsText(ulong? totalMemBytes = null)
    {
        var totalRamText = "Unknown RAM";
        if (totalMemBytes.HasValue)
        {
            totalRamText = $"{ToGiB((long)totalMemBytes.Value):0.0} GB RAM";
        }
        else
        {
            var ms = new MemoryStatusEx();
            if (GlobalMemoryStatusEx(ms))
            {
                totalRamText = $"{ToGiB((long)ms.ullTotalPhys):0.0} GB RAM";
            }
        }

        var cpuMaxText = _cpuMaxMhz > 0 ? $"Max {_cpuMaxMhz:0} MHz" : "Max clock unknown";
        return $"CPU: {_cpuName} ({cpuMaxText})  |  GPU: {_gpuName}  |  {totalRamText}  |  OS: {Environment.OSVersion.VersionString}";
    }

    private void SeedTweakCards()
    {
        if (_tweakCards.Count > 0)
        {
            return;
        }

        _tweakCards.Add(new TweakCardItem
        {
            Key = "sensor_suite_off",
            Title = "Sensor Suite Off",
            Description = "Disables sensor/location related services and reduces unnecessary polling during gaming."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "sticky_keys_guard",
            Title = "StickyKeys Guard",
            Description = "Prevents accidental StickyKeys popups and focus steal while gaming."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "telemetry_zero",
            Title = "Telemetry Zero",
            Description = "Forces diagnostic telemetry to minimum policy when allowed.",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "voice_activation_off",
            Title = "Voice Activation Off",
            Description = "Disables voice activation listeners to reduce idle background activity."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "webdav_scan_off",
            Title = "WebDAV Scan Off",
            Description = "Stops WebClient polling to reduce remote scan wakeups."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "whql_only",
            Title = "WHQL Only",
            Description = "Reports driver status and enforces safer certified-driver workflow."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "wu_control",
            Title = "WU Control",
            Description = "Shows Windows Update state and allows safer scheduling strategy.",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "wu_core_off",
            Title = "WU Core Off",
            Description = "Attempts to pause core Windows Update services (admin required).",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "amd_chill_off",
            Title = "AMD Chill Off",
            Description = "Disables Radeon Chill and related frame pacing power-saver behavior."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "amd_power_hold",
            Title = "AMD Power Hold",
            Description = "Keeps steadier GPU clocks under sustained load.",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "amd_service_trim",
            Title = "AMD Service Trim",
            Description = "Cuts non-essential AMD background services."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "cloud_sync_off",
            Title = "Cloud Sync Off",
            Description = "Disables settings sync across devices to reduce background churn."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "do_solo_mode",
            Title = "DO Solo Mode",
            Description = "Forces Delivery Optimization to Microsoft-only download mode."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "process_count_reduction",
            Title = "Process Count Reduction",
            Description = "Consolidates service host process count to reduce background process load.",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "nudge_blocker",
            Title = "Nudge Blocker",
            Description = "Reduces notifications and feedback prompts during sessions."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "low_latency_mode",
            Title = "Low Latency Mode",
            Description = "Applies scheduler, network throttling, and DVR-related settings to reduce system-side input and frame delay.",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "network_driver_optimize",
            Title = "Network Driver Optimize",
            Description = "Disables common NIC power-saving/latency features and applies gaming-friendly network adapter settings.",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "ultimate_power_plan",
            Title = "Ultimate Power Plan",
            Description = "Enables Ultimate Performance, sets minimum CPU state to 100%, and disables USB selective suspend.",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "memory_integrity_off",
            Title = "Memory Integrity Off",
            Description = "Turns off Core Isolation Memory Integrity (security tradeoff).",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "gpu_msi_mode",
            Title = "GPU MSI Mode",
            Description = "Enables MSI interrupt mode for detected PCI GPUs (advanced).",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "hpet_tune_off",
            Title = "HPET Tune Off",
            Description = "Disables forced platform clock usage via boot config (hardware-dependent).",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "startup_cleanup_assist",
            Title = "Startup Cleanup Assist",
            Description = "Opens startup app controls and records managed state for clean boot workflow."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "nvidia_latency_profile",
            Title = "NVIDIA Latency Profile",
            Description = "Applies available NVIDIA CLI settings and opens guidance for low-latency control panel values.",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "hibernation_off",
            Title = "Hibernation Off",
            Description = "Disables hibernation to reduce disk footprint and hiberfile writes."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "gamebar_overlay_off",
            Title = "GameBar Overlay Off",
            Description = "Disables Xbox Game Bar overlays and capture background activity."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "widgets_feed_off",
            Title = "Widgets Feed Off",
            Description = "Turns off widgets feed and taskbar widget integration."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "ads_id_off",
            Title = "Ads ID Off",
            Description = "Disables advertising ID and personalized ad tracking.",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "search_highlights_off",
            Title = "Search Highlights Off",
            Description = "Disables search highlights and web suggestions in search."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "remote_assistance_off",
            Title = "Remote Assistance Off",
            Description = "Disables Remote Assistance to reduce remote exposure.",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "fast_startup_off",
            Title = "Fast Startup Off",
            Description = "Disables fast startup for cleaner full shutdown behavior."
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "print_spooler_off",
            Title = "Print Spooler Off",
            Description = "Disables Print Spooler service on systems without printers.",
            WarningText = "1 WARNING"
        });
        _tweakCards.Add(new TweakCardItem
        {
            Key = "smb1_off",
            Title = "SMB1 Off",
            Description = "Disables legacy SMB1 protocol support for stronger security.",
            WarningText = "1 WARNING"
        });
    }

    private async Task RefreshTweakStatesAsync(bool showPopup)
    {
        if (_isRefreshingTweakStates || _tweakCards.Count == 0)
        {
            return;
        }

        _isRefreshingTweakStates = true;
        _lastTweakStateRefreshUtc = DateTime.UtcNow;

        try
        {
            var states = new Dictionary<string, bool?>(StringComparer.OrdinalIgnoreCase);
            foreach (var card in _tweakCards)
            {
                states[card.Key] = await _optimizer.GetTweakStateAsync(card.Key);
            }

            var drifted = new List<string>();
            _suppressToggleEvents = true;
            try
            {
                foreach (var card in _tweakCards)
                {
                    var state = states[card.Key];
                    if (state is true)
                    {
                        card.LiveStateText = "Live: ON";
                        card.IsEnabled = true;
                    }
                    else if (state is false)
                    {
                        card.LiveStateText = "Live: OFF";
                        card.IsEnabled = false;
                        drifted.Add(card.Title);
                    }
                    else
                    {
                        card.LiveStateText = "Live: Unknown";
                    }
                }
            }
            finally
            {
                _suppressToggleEvents = false;
            }

            if (showPopup && drifted.Count > 0)
            {
                var names = string.Join(", ", drifted.Take(6));
                var suffix = drifted.Count > 6 ? $" (+{drifted.Count - 6} more)" : string.Empty;
                OutputTextBox.Text = $"Re-apply available for: {names}{suffix}";
            }
        }
        finally
        {
            _isRefreshingTweakStates = false;
        }
    }

    private void RefreshGpuTelemetry()
    {
        if ((DateTime.UtcNow - _lastGpuRefreshUtc).TotalSeconds < 2)
        {
            return;
        }

        _lastGpuRefreshUtc = DateTime.UtcNow;
        try
        {
            using var searcher = new ManagementObjectSearcher("SELECT Name, UtilizationPercentage FROM Win32_PerfFormattedData_GPUPerformanceCounters_GPUEngine");
            var total = 0d;
            foreach (var obj in searcher.Get().OfType<ManagementObject>())
            {
                var name = obj["Name"]?.ToString() ?? string.Empty;
                if (!name.Contains("engtype_3D", StringComparison.OrdinalIgnoreCase) &&
                    !name.Contains("engtype_Compute", StringComparison.OrdinalIgnoreCase) &&
                    !name.Contains("engtype_Cuda", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (double.TryParse(obj["UtilizationPercentage"]?.ToString(), out var util))
                {
                    total += util;
                }
            }

            if (total > 0)
            {
                var usage = ClampPercent(total);
                GpuUsage = usage;
                GpuUsageText = $"{usage:0}%";
                GpuDetailsText = $"Usage: {usage:0.0}%";
            }
            else
            {
                GpuUsage = 0;
                GpuUsageText = "N/A";
                GpuDetailsText = "Usage: N/A";
            }
        }
        catch
        {
            GpuUsage = 0;
            GpuUsageText = "N/A";
            GpuDetailsText = "Usage: N/A";
        }
    }

    private void RefreshNetworkTelemetry()
    {
        try
        {
            var nics = NetworkInterface.GetAllNetworkInterfaces()
                .Where(n => n.OperationalStatus == OperationalStatus.Up &&
                            n.NetworkInterfaceType != NetworkInterfaceType.Loopback &&
                            n.NetworkInterfaceType != NetworkInterfaceType.Tunnel)
                .ToList();
            if (nics.Count == 0)
            {
                NetworkAdapterText = "Adapter: N/A";
                return;
            }

            var selected = nics
                .Select(n => new { Nic = n, Stats = n.GetIPv4Statistics() })
                .OrderByDescending(x => (long)x.Stats.BytesSent + x.Stats.BytesReceived)
                .First();

            var sent = selected.Stats.BytesSent;
            var recv = selected.Stats.BytesReceived;
            var now = DateTime.UtcNow;
            if (_lastNetworkSampleUtc != DateTime.MinValue)
            {
                var dt = (now - _lastNetworkSampleUtc).TotalSeconds;
                if (dt > 0.01)
                {
                    var upMbps = Math.Max(0, (sent - _lastNetworkBytesSent) * 8.0 / dt / 1_000_000.0);
                    var downMbps = Math.Max(0, (recv - _lastNetworkBytesReceived) * 8.0 / dt / 1_000_000.0);
                    NetworkUploadText = $"Upload: {upMbps:0.00} Mbps";
                    NetworkDownloadText = $"Download: {downMbps:0.00} Mbps";
                }
            }

            _lastNetworkBytesSent = sent;
            _lastNetworkBytesReceived = recv;
            _lastNetworkSampleUtc = now;
            NetworkAdapterText = $"Adapter: {selected.Nic.Name}";
            NetworkTotalsText = $"Sent/Recv: {sent / 1024d / 1024d / 1024d:0.00} / {recv / 1024d / 1024d / 1024d:0.00} GB";
        }
        catch
        {
            NetworkAdapterText = "Adapter: N/A";
        }
    }

    private void RefreshTemperatureTelemetry()
    {
        if ((DateTime.UtcNow - _lastTempRefreshUtc).TotalSeconds < 2)
        {
            return;
        }

        _lastTempRefreshUtc = DateTime.UtcNow;
        _latestCpuTempC = ReadCpuTempC();
        _latestGpuTempC = ReadGpuTempC();
        var current = _tempSource == "GPU" ? _latestGpuTempC : _latestCpuTempC;
        if (!current.HasValue)
        {
            TemperatureValueText = "N/A";
            TemperatureDetailText = $"Source: {_tempSource} | Sensor unavailable";
            TemperatureTrendText = "Trend: --";
            return;
        }

        _tempHistory.Add(current.Value);
        while (_tempHistory.Count > 25)
        {
            _tempHistory.RemoveAt(0);
        }

        TemperatureValueText = $"{current.Value:0.0}°C";
        TemperatureDetailText = $"Source: {_tempSource} | Min {_tempHistory.Min():0.0}°C  Max {_tempHistory.Max():0.0}°C";
        if (_tempHistory.Count >= 2)
        {
            var delta = _tempHistory[^1] - _tempHistory[^2];
            TemperatureTrendText = $"Trend: {delta:+0.0;-0.0;0.0}°C / sample";
        }
        else
        {
            TemperatureTrendText = "Trend: collecting...";
        }
    }

    private double? ReadCpuTempC()
    {
        if (TryReadCoreTempSharedMemory())
        {
            if (_coreTempPackageTempC > 0)
            {
                return _coreTempPackageTempC;
            }
        }

        try
        {
            using var searcher = new ManagementObjectSearcher(@"root\wmi", "SELECT CurrentTemperature FROM MSAcpi_ThermalZoneTemperature");
            var vals = new List<double>();
            foreach (var obj in searcher.Get().OfType<ManagementObject>())
            {
                if (double.TryParse(obj["CurrentTemperature"]?.ToString(), out var raw) && raw > 0)
                {
                    vals.Add((raw / 10.0) - 273.15);
                }
            }

            if (vals.Count > 0)
            {
                return vals.Average();
            }
        }
        catch
        {
            // ignored
        }

        return null;
    }

    private bool TryReadCoreTempSharedMemory()
    {
        if ((DateTime.UtcNow - _lastCoreTempProbeUtc).TotalMilliseconds < 700)
        {
            return _coreTempAvailable;
        }

        _lastCoreTempProbeUtc = DateTime.UtcNow;
        try
        {
            using var mmf = MemoryMappedFile.OpenExisting("CoreTempMappingObjectEx", MemoryMappedFileRights.Read);
            using var view = mmf.CreateViewAccessor(0, 4096, MemoryMappedFileAccess.Read);

            // Core Temp SharedDataEx layout (4-byte aligned):
            // uiCoreCnt @ 1536, fTemp[0] @ 1544, fCPUSpeed @ 2572
            var coreCount = view.ReadUInt32(1536);
            var packageTemp = view.ReadSingle(1544);
            var cpuSpeed = view.ReadSingle(2572);

            if (coreCount == 0 || packageTemp <= 0)
            {
                _coreTempAvailable = false;
                return false;
            }

            _coreTempPackageTempC = packageTemp;
            _coreTempCpuSpeedMhz = cpuSpeed;
            _coreTempAvailable = true;
            return true;
        }
        catch
        {
            _coreTempAvailable = false;
            return false;
        }
    }

    private double? ReadGpuTempC()
    {
        try
        {
            using var proc = new Process();
            proc.StartInfo = new ProcessStartInfo
            {
                FileName = "nvidia-smi",
                Arguments = "--query-gpu=temperature.gpu --format=csv,noheader,nounits",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            proc.Start();
            var stdout = proc.StandardOutput.ReadToEnd().Trim();
            if (!proc.WaitForExit(600))
            {
                return null;
            }

            if (double.TryParse(stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault(), out var temp))
            {
                return temp;
            }
        }
        catch
        {
            // ignored
        }

        return null;
    }

    private void TempCpuButton_Click(object sender, RoutedEventArgs e)
    {
        _tempSource = "CPU";
        _tempHistory.Clear();
        UpdateTempSourceButtons();
        RefreshTemperatureTelemetry();
    }

    private void TempGpuButton_Click(object sender, RoutedEventArgs e)
    {
        _tempSource = "GPU";
        _tempHistory.Clear();
        UpdateTempSourceButtons();
        RefreshTemperatureTelemetry();
    }

    private void UpdateTempSourceButtons()
    {
        if (TempCpuButton is null || TempGpuButton is null)
        {
            return;
        }

        TempCpuButton.Background = _tempSource == "CPU"
            ? new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#1E4F8C")!)
            : new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#112240")!);
        TempGpuButton.Background = _tempSource == "GPU"
            ? new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#1E4F8C")!)
            : new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#112240")!);
    }

    private void LoadControllerDevices()
    {
        _controllerDevices.Clear();
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT Name, DeviceID, PNPClass FROM Win32_PnPEntity WHERE DeviceID IS NOT NULL");
            foreach (var obj in searcher.Get().OfType<ManagementObject>())
            {
                var name = obj["Name"]?.ToString()?.Trim();
                var id = obj["DeviceID"]?.ToString()?.Trim();
                var pnpClass = obj["PNPClass"]?.ToString()?.Trim();
                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(id))
                {
                    continue;
                }

                var lowerName = name.ToLowerInvariant();
                var lowerId = id.ToLowerInvariant();
                var looksController =
                    lowerName.Contains("controller") ||
                    lowerName.Contains("gamepad") ||
                    lowerName.Contains("joystick") ||
                    lowerName.Contains("xinput") ||
                    lowerName.Contains("xbox") ||
                    lowerName.Contains("dualshock") ||
                    lowerName.Contains("dualsense");
                var looksUsbHid =
                    string.Equals(pnpClass, "HIDClass", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(pnpClass, "USB", StringComparison.OrdinalIgnoreCase) ||
                    lowerId.StartsWith("usb\\", StringComparison.OrdinalIgnoreCase) ||
                    lowerId.StartsWith("hid\\", StringComparison.OrdinalIgnoreCase);

                if (!looksUsbHid || !looksController)
                {
                    continue;
                }

                _controllerDevices.Add(new ControllerDeviceItem
                {
                    DisplayName = name,
                    InstanceId = id
                });
            }
        }
        catch
        {
            // ignored
        }

        ControllerDeviceComboBox.ItemsSource = _controllerDevices;
        if (_controllerDevices.Count > 0)
        {
            ControllerDeviceComboBox.SelectedIndex = 0;
        }
        else
        {
            ControllerPollingStatusText = "No controller/USB gaming device detected.";
        }

        if (PollingRateComboBox.Items.Count == 0)
        {
            PollingRateComboBox.ItemsSource = new[] { "125 Hz", "250 Hz", "500 Hz", "1000 Hz", "2000 Hz", "4000 Hz", "8000 Hz" };
            PollingRateComboBox.SelectedIndex = 3;
        }
    }

    private void RefreshControllerDevicesButton_Click(object sender, RoutedEventArgs e)
    {
        LoadControllerDevices();
        StatusText = "Controller device list refreshed";
    }

    private void ControllerDeviceComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ControllerDeviceComboBox.SelectedItem is not ControllerDeviceItem item)
        {
            return;
        }

        ControllerPollingStatusText = $"Selected: {item.DisplayName}";
    }

    private void ApplyControllerPollingButton_Click(object sender, RoutedEventArgs e)
    {
        if (ControllerDeviceComboBox.SelectedItem is not ControllerDeviceItem item)
        {
            ControllerPollingStatusText = "Select a controller/USB device first.";
            return;
        }

        var text = PollingRateComboBox.SelectedItem?.ToString() ?? "1000 Hz";
        if (!int.TryParse(text.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault(), out var hz))
        {
            hz = 1000;
        }

        var profileKey = item.InstanceId.Replace("\\", "#");
        try
        {
            using var appKey = Registry.CurrentUser.CreateSubKey($@"Software\AsherasTweakingUtility\ControllerPolling\{profileKey}", true);
            appKey?.SetValue("PollingHz", hz, RegistryValueKind.DWord);
            appKey?.SetValue("DisplayName", item.DisplayName, RegistryValueKind.String);
        }
        catch
        {
            // ignore profile write failures
        }

        // Apply to HIDUSBF if installed. This keeps the app original and uses local system capabilities.
        var applied = false;
        try
        {
            using var devKey = Registry.LocalMachine.CreateSubKey($@"SYSTEM\CurrentControlSet\Services\HIDUSBF\Parameters\Devices\{profileKey}", true);
            if (devKey is not null)
            {
                devKey.SetValue("Rate", hz, RegistryValueKind.DWord);
                devKey.SetValue("FilterOn", 1, RegistryValueKind.DWord);
                applied = true;
            }
        }
        catch
        {
            applied = false;
        }

        StatusText = "Controller polling profile updated";
        ControllerPollingStatusText = applied
            ? $"Applied {hz} Hz to {item.DisplayName}. Replug/restart the device to finalize."
            : $"Saved {hz} Hz profile for {item.DisplayName}. Install/enable HIDUSBF filter and re-apply.";
        OutputTextBox.Text =
            "Controller Polling\n" +
            "======================================================================\n" +
            $"Device: {item.DisplayName}\n" +
            $"Instance: {item.InstanceId}\n" +
            $"Polling target: {hz} Hz\n" +
            (applied
                ? "Driver-level filter settings were written. Restart/replug controller."
                : "Only profile saved. HIDUSBF filter driver keys were not writable (run as admin / install filter).");
    }

    private void LoadFortniteRegionOptions()
    {
        if (_fortniteRegions.Count > 0)
        {
            return;
        }

        _fortniteRegions.Add(new FortniteRegionItem { Code = "AUTO", DisplayName = "Auto (closest)", Hostname = "__AUTO__" });
        foreach (var target in FortniteServerTargets)
        {
            _fortniteRegions.Add(new FortniteRegionItem
            {
                Code = target.Code,
                DisplayName = target.DisplayName,
                Hostname = target.Hostname
            });
        }

        FortniteRegionComboBox.ItemsSource = _fortniteRegions;
        FortniteRegionComboBox.SelectedValue = "__AUTO__";
    }

    private void LoadFortniteServerTargets()
    {
        if (_fortniteServerPings.Count > 0)
        {
            return;
        }

        foreach (var target in FortniteServerTargets)
        {
            _fortniteServerPings.Add(new FortniteServerPingItem
            {
                Region = target.Region,
                Code = target.Code,
                Hostname = target.Hostname,
                PingText = "--"
            });
        }
    }

    private void ReloadFortniteRegionState()
    {
        LoadFortniteRegionOptions();

        var path = GetFortniteSettingsPath();
        if (!File.Exists(path))
        {
            FortniteRegionStatusText = "Fortnite config not found yet. Launch Fortnite once, then reload.";
            FortniteRegionComboBox.SelectedValue = "__AUTO__";
            return;
        }

        var current = ReadFortnitePreferredRegion(path);
        if (string.IsNullOrWhiteSpace(current))
        {
            FortniteRegionStatusText = $"Config found at {path}. Region not set; game will use Auto.";
            FortniteRegionComboBox.SelectedValue = "__AUTO__";
            return;
        }

        var normalized = current.Trim().ToUpperInvariant();
        var selected = _fortniteRegions.FirstOrDefault(r => r.Code == normalized);
        FortniteRegionComboBox.SelectedValue = selected?.Hostname ?? "__AUTO__";
        var known = selected is not null;
        FortniteRegionStatusText = known
            ? $"Current Fortnite region: {normalized}"
            : $"Current Fortnite region in config: {normalized} (unknown code)";
    }

    private void ReloadFortniteRegionButton_Click(object sender, RoutedEventArgs e)
    {
        ReloadFortniteRegionState();
        StatusText = "Fortnite region reloaded";
    }

    private void SuggestFortniteRegionButton_Click(object sender, RoutedEventArgs e)
    {
        LoadFortniteRegionOptions();
        var bestMeasured = _fortniteServerPings
            .Where(s => s.PingMs.HasValue)
            .OrderBy(s => s.PingMs!.Value)
            .FirstOrDefault();
        var suggestedCode = bestMeasured?.Code ?? SuggestClosestFortniteRegion();
        var selected = bestMeasured is not null
            ? _fortniteRegions.FirstOrDefault(r => string.Equals(r.Hostname, bestMeasured.Hostname, StringComparison.OrdinalIgnoreCase))
            : _fortniteRegions.FirstOrDefault(r => r.Code == suggestedCode);
        FortniteRegionComboBox.SelectedValue = selected?.Hostname ?? "__AUTO__";
        FortniteRegionStatusText = bestMeasured is null
            ? $"Suggested region based on local timezone: {suggestedCode}"
            : $"Suggested best measured server: {selected?.DisplayName ?? suggestedCode} ({bestMeasured.PingMs:0} ms)";
        StatusText = "Fortnite region suggested";
    }

    private void ApplyFortniteRegionButton_Click(object sender, RoutedEventArgs e)
    {
        LoadFortniteRegionOptions();
        var selected = FortniteRegionComboBox.SelectedItem as FortniteRegionItem;
        var code = (selected?.Code ?? "AUTO").Trim().ToUpperInvariant();
        if (!_fortniteRegions.Any(r => r.Code == code))
        {
            code = "AUTO";
        }

        var path = GetFortniteSettingsPath();
        var folder = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(folder))
        {
            Directory.CreateDirectory(folder);
        }

        var lines = File.Exists(path)
            ? File.ReadAllLines(path).ToList()
            : new List<string>();

        var keyIndex = lines.FindIndex(l => l.TrimStart().StartsWith("PreferredRegion=", StringComparison.OrdinalIgnoreCase));
        if (keyIndex >= 0)
        {
            lines[keyIndex] = $"PreferredRegion={code}";
        }
        else
        {
            if (lines.Count > 0 && !string.IsNullOrWhiteSpace(lines[^1]))
            {
                lines.Add(string.Empty);
            }

            lines.Add("[/Script/FortniteGame.FortGameUserSettings]");
            lines.Add($"PreferredRegion={code}");
        }

        File.WriteAllLines(path, lines);
        FortniteRegionStatusText = selected is null
            ? $"Applied Fortnite region: {code}. Restart Fortnite to take effect."
            : $"Applied server selection: {selected.DisplayName} -> region {code}. Restart Fortnite to take effect.";
        StatusText = "Fortnite region updated";
        OutputTextBox.Text =
            "Fortnite Region\n" +
            "======================================================================\n" +
            $"Config: {path}\n" +
            $"PreferredRegion: {code}\n" +
            (selected is null ? string.Empty : $"Selected Server Node: {selected.DisplayName} ({selected.Hostname})\n") +
            "Restart Fortnite to apply the new server region.";
    }

    private static string GetFortniteSettingsPath()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "FortniteGame",
            "Saved",
            "Config",
            "WindowsClient",
            "GameUserSettings.ini");
    }

    private static string? ReadFortnitePreferredRegion(string path)
    {
        foreach (var line in File.ReadLines(path))
        {
            var trimmed = line.Trim();
            if (!trimmed.StartsWith("PreferredRegion=", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var idx = trimmed.IndexOf('=');
            if (idx < 0 || idx == trimmed.Length - 1)
            {
                return null;
            }

            return trimmed[(idx + 1)..].Trim();
        }

        return null;
    }

    private static string SuggestClosestFortniteRegion()
    {
        var tz = TimeZoneInfo.Local.Id.ToLowerInvariant();
        if (tz.Contains("eastern"))
        {
            return "NAE";
        }

        if (tz.Contains("central"))
        {
            return "NAC";
        }

        if (tz.Contains("mountain") || tz.Contains("pacific") || tz.Contains("alaska") || tz.Contains("hawai"))
        {
            return "NAW";
        }

        if (tz.Contains("europe") || tz.Contains("gmt"))
        {
            return "EU";
        }

        if (tz.Contains("australia") || tz.Contains("new zealand"))
        {
            return "OCE";
        }

        if (tz.Contains("tokyo") || tz.Contains("korea") || tz.Contains("china") || tz.Contains("singapore"))
        {
            return "ASIA";
        }

        if (tz.Contains("brazil") || tz.Contains("sa eastern"))
        {
            return "BR";
        }

        if (tz.Contains("arab") || tz.Contains("middle east"))
        {
            return "ME";
        }

        return "AUTO";
    }

    private async void PingFortniteServersButton_Click(object sender, RoutedEventArgs e)
    {
        await PingFortniteServersAsync(force: true);
    }

    private async Task PingFortniteServersAsync(bool force)
    {
        LoadFortniteServerTargets();
        if (_isPingingFortniteServers)
        {
            return;
        }

        if (!force && (DateTime.UtcNow - _lastFortnitePingUtc).TotalSeconds < 30)
        {
            return;
        }

        _isPingingFortniteServers = true;
        StatusText = "Pinging Fortnite servers...";
        FortniteRegionStatusText = "Measuring ping for all Fortnite regions...";
        foreach (var server in _fortniteServerPings)
        {
            server.PingMs = null;
            server.PingText = "Testing...";
        }

        try
        {
            var tasks = _fortniteServerPings.Select(async server =>
            {
                var ms = await MeasureAveragePingAsync(server.Hostname);
                return (server, ms);
            });
            var results = await Task.WhenAll(tasks);

            foreach (var (server, ms) in results)
            {
                server.PingMs = ms;
                server.PingText = ms.HasValue ? $"{ms.Value:0}" : "Timeout";
            }

            var best = _fortniteServerPings.Where(s => s.PingMs.HasValue).OrderBy(s => s.PingMs!.Value).FirstOrDefault();
            FortniteRegionStatusText = best is null
                ? "Ping test complete. No datacenter endpoint responded."
                : $"Ping test complete. Best measured region: {best.Code} ({best.PingMs:0} ms)";
            StatusText = "Fortnite ping test complete";
            _lastFortnitePingUtc = DateTime.UtcNow;

            OutputTextBox.Text =
                "Fortnite Datacenter Ping\n" +
                "======================================================================\n" +
                string.Join("\n", _fortniteServerPings.Select(s => $"{s.Region} ({s.Code}) -> {s.PingText} ms [{s.Hostname}]"));
        }
        finally
        {
            _isPingingFortniteServers = false;
        }
    }

    private static async Task<double?> MeasureAveragePingAsync(string host)
    {
        var samples = new List<long>(3);
        for (var i = 0; i < 3; i++)
        {
            try
            {
                using var ping = new Ping();
                var reply = await ping.SendPingAsync(host, 1400);
                if (reply.Status == IPStatus.Success)
                {
                    samples.Add(reply.RoundtripTime);
                }
            }
            catch
            {
                // ignored
            }

            await Task.Delay(60);
        }

        if (samples.Count == 0)
        {
            return null;
        }

        return samples.Average();
    }

    private void SetDashboardView()
    {
        IsDashboardVisible = true;
        IsGeneralTweaksVisible = false;
        IsGameProfilesVisible = false;
    }

    private void SetGeneralTweaksView()
    {
        IsDashboardVisible = false;
        IsGeneralTweaksVisible = true;
        IsGameProfilesVisible = false;
    }

    private void SetGameProfilesView()
    {
        IsDashboardVisible = false;
        IsGeneralTweaksVisible = false;
        IsGameProfilesVisible = true;
    }

    private static bool IsRestartRequiredTweak(string tweakKey)
    {
        return tweakKey is
            "fast_startup_off" or
            "smb1_off" or
            "network_driver_optimize" or
            "memory_integrity_off" or
            "gpu_msi_mode" or
            "hpet_tune_off";
    }

    private static string? ExtractUpdaterPath(string text)
    {
        foreach (var line in text.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            var trimmed = line.Trim();
            if (!trimmed.StartsWith("Updater script:", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var idx = trimmed.IndexOf(':');
            if (idx < 0 || idx == trimmed.Length - 1)
            {
                return null;
            }

            return trimmed[(idx + 1)..].Trim();
        }

        return null;
    }

    private static double ClampPercent(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            return 0;
        }

        return Math.Min(100, Math.Max(0, value));
    }

    private static double ToGiB(long bytes)
    {
        return bytes / 1024d / 1024d / 1024d;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern bool GlobalMemoryStatusEx([In, Out] MemoryStatusEx lpBuffer);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    private sealed class MemoryStatusEx
    {
        public uint dwLength = (uint)Marshal.SizeOf<MemoryStatusEx>();
        public uint dwMemoryLoad;
        public ulong ullTotalPhys;
        public ulong ullAvailPhys;
        public ulong ullTotalPageFile;
        public ulong ullAvailPageFile;
        public ulong ullTotalVirtual;
        public ulong ullAvailVirtual;
        public ulong ullAvailExtendedVirtual;
    }

    private sealed class ControllerDeviceItem
    {
        public required string DisplayName { get; init; }
        public required string InstanceId { get; init; }
    }

    private sealed class FortniteRegionItem
    {
        public required string Code { get; init; }
        public required string DisplayName { get; init; }
        public required string Hostname { get; init; }
    }

    public sealed class FortniteServerPingItem : INotifyPropertyChanged
    {
        private string _pingText = "--";
        private double? _pingMs;

        public required string Region { get; init; }
        public required string Code { get; init; }
        public required string Hostname { get; init; }

        public double? PingMs
        {
            get => _pingMs;
            set
            {
                if (_pingMs == value)
                {
                    return;
                }

                _pingMs = value;
                OnPropertyChanged();
            }
        }

        public string PingText
        {
            get => _pingText;
            set
            {
                if (_pingText == value)
                {
                    return;
                }

                _pingText = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    private static readonly FortniteServerTarget[] FortniteServerTargets =
    [
        new FortniteServerTarget { Region = "Ashburn, US-East", Code = "NAE", Hostname = "ec2.us-east-1.amazonaws.com", DisplayName = "Ashburn (NA-East)" },
        new FortniteServerTarget { Region = "Columbus, US-East", Code = "NAE", Hostname = "ec2.us-east-2.amazonaws.com", DisplayName = "Columbus (NA-East)" },
        new FortniteServerTarget { Region = "Montreal, NA-Central", Code = "NAC", Hostname = "ec2.ca-central-1.amazonaws.com", DisplayName = "Montreal (NA-Central)" },
        new FortniteServerTarget { Region = "Portland, US-West", Code = "NAW", Hostname = "ec2.us-west-2.amazonaws.com", DisplayName = "Portland (NA-West)" },
        new FortniteServerTarget { Region = "Frankfurt, EU", Code = "EU", Hostname = "ec2.eu-central-1.amazonaws.com", DisplayName = "Frankfurt (Europe)" },
        new FortniteServerTarget { Region = "London, EU", Code = "EU", Hostname = "ec2.eu-west-2.amazonaws.com", DisplayName = "London (Europe)" },
        new FortniteServerTarget { Region = "Paris, EU", Code = "EU", Hostname = "ec2.eu-west-3.amazonaws.com", DisplayName = "Paris (Europe)" },
        new FortniteServerTarget { Region = "Sao Paulo, BR", Code = "BR", Hostname = "ec2.sa-east-1.amazonaws.com", DisplayName = "Sao Paulo (Brazil)" },
        new FortniteServerTarget { Region = "Tokyo, Asia", Code = "ASIA", Hostname = "ec2.ap-northeast-1.amazonaws.com", DisplayName = "Tokyo (Asia)" },
        new FortniteServerTarget { Region = "Seoul, Asia", Code = "ASIA", Hostname = "ec2.ap-northeast-2.amazonaws.com", DisplayName = "Seoul (Asia)" },
        new FortniteServerTarget { Region = "Singapore, Asia", Code = "ASIA", Hostname = "ec2.ap-southeast-1.amazonaws.com", DisplayName = "Singapore (Asia)" },
        new FortniteServerTarget { Region = "Sydney, Oceania", Code = "OCE", Hostname = "ec2.ap-southeast-2.amazonaws.com", DisplayName = "Sydney (Oceania)" },
        new FortniteServerTarget { Region = "Bahrain, Middle East", Code = "ME", Hostname = "ec2.me-south-1.amazonaws.com", DisplayName = "Bahrain (Middle East)" },
        new FortniteServerTarget { Region = "Dubai, Middle East", Code = "ME", Hostname = "ec2.me-central-1.amazonaws.com", DisplayName = "Dubai (Middle East)" }
    ];

    private sealed class FortniteServerTarget
    {
        public required string Region { get; init; }
        public required string Code { get; init; }
        public required string Hostname { get; init; }
        public required string DisplayName { get; init; }
    }

    private void MainWindow_Closing(object? sender, CancelEventArgs e)
    {
        _telemetryTimer.Stop();
        _cpuCounter?.Dispose();
        _cpuFrequencyCounter?.Dispose();
        _cpuFrequencyPercentCounter?.Dispose();
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

