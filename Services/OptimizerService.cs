using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using WinOptApp.Models;

namespace WinOptApp.Services;

public sealed class OptimizerService
{
    private const string GameBarKey = @"Software\Microsoft\GameBar";
    private const string GameBarValueName = "AutoGameModeEnabled";
    private const string LatestExeApiUrl = "https://api.github.com/repos/Daquan-sudem/AsherasTweakingUtility/contents/downloads/AsherasTweakingUtility-win-x64-latest.exe";
    private static readonly HttpClient UpdateClient = CreateUpdateClient();

    private static readonly string StateFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
        "WinOptApp",
        "rollback-state.json");

    private static readonly string ManagedStateFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
        "WinOptApp",
        "managed-tweaks.json");

    static OptimizerService()
    {
        AppDomain.CurrentDomain.ProcessExit += (_, _) =>
        {
            try
            {
                _ = NtSetTimerResolution(5000, false, out _);
            }
            catch
            {
                // ignored
            }
        };
    }

    public async Task<string> CheckForUpdateAsync()
    {
        var sb = new StringBuilder();
        sb.AppendLine("Update Check");
        sb.AppendLine(new string('=', 70));
        sb.AppendLine($"Timestamp: {DateTime.Now}");
        sb.AppendLine();

        var remote = await GetLatestExecutableAsync();
        if (remote is null || string.IsNullOrWhiteSpace(remote.download_url))
        {
            sb.AppendLine("Unable to read update metadata from GitHub.");
            sb.AppendLine("STATUS: UPDATE_CHECK_FAILED");
            return sb.ToString();
        }

        sb.AppendLine($"Latest file: {remote.name}");
        sb.AppendLine($"Remote size: {remote.size / 1024d / 1024d:0.00} MB");

        var currentExe = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
        if (string.IsNullOrWhiteSpace(currentExe) || !File.Exists(currentExe))
        {
            sb.AppendLine("Local executable path not found. Run packaged EXE for direct version comparison.");
            sb.AppendLine($"Download: {remote.download_url}");
            sb.AppendLine("STATUS: UPDATE_UNKNOWN");
            return sb.ToString();
        }

        if (!currentExe.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
        {
            sb.AppendLine("You are running from a development host. Packaged EXE comparison is unavailable.");
            sb.AppendLine($"Download: {remote.download_url}");
            sb.AppendLine("STATUS: UPDATE_UNKNOWN");
            return sb.ToString();
        }

        var localSize = new FileInfo(currentExe).Length;
        sb.AppendLine($"Local file: {Path.GetFileName(currentExe)}");
        sb.AppendLine($"Local size: {localSize / 1024d / 1024d:0.00} MB");

        if (localSize == remote.size)
        {
            sb.AppendLine();
            sb.AppendLine("Result: Up to date.");
            sb.AppendLine("STATUS: UP_TO_DATE");
        }
        else
        {
            sb.AppendLine();
            sb.AppendLine("Result: Update available.");
            sb.AppendLine($"Download: {remote.download_url}");
            sb.AppendLine("STATUS: UPDATE_AVAILABLE");
        }

        return sb.ToString();
    }

    public async Task<string> DownloadLatestUpdateAsync()
    {
        var sb = new StringBuilder();
        sb.AppendLine("Download Update");
        sb.AppendLine(new string('=', 70));
        sb.AppendLine($"Timestamp: {DateTime.Now}");
        sb.AppendLine();

        var remote = await GetLatestExecutableAsync();
        if (remote is null || string.IsNullOrWhiteSpace(remote.download_url))
        {
            sb.AppendLine("Unable to locate update binary on GitHub.");
            return sb.ToString();
        }

        var currentExe = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
        if (string.IsNullOrWhiteSpace(currentExe) || !currentExe.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
        {
            sb.AppendLine("In-place update requires running from the packaged EXE.");
            return sb.ToString();
        }

        var appDir = Path.GetDirectoryName(currentExe)!;
        var stagingPath = Path.Combine(appDir, $"AsherasTweakingUtility-update-{DateTime.Now:yyyyMMdd-HHmmss}.new");
        var updaterPath = Path.Combine(Path.GetTempPath(), $"AsherasTU-updater-{DateTime.Now:yyyyMMdd-HHmmss}.cmd");

        using var response = await UpdateClient.GetAsync(remote.download_url, HttpCompletionOption.ResponseHeadersRead);
        response.EnsureSuccessStatusCode();

        await using (var source = await response.Content.ReadAsStreamAsync())
        await using (var destination = new FileStream(stagingPath, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            await source.CopyToAsync(destination);
        }

        var script = string.Join(Environment.NewLine,
            "@echo off",
            "setlocal",
            "timeout /t 2 /nobreak >nul",
            ":retry",
            $"copy /y \"{stagingPath}\" \"{currentExe}\" >nul 2>&1",
            "if errorlevel 1 (",
            "  timeout /t 1 /nobreak >nul",
            "  goto retry",
            ")",
            $"start \"\" \"{currentExe}\"",
            $"del /f /q \"{stagingPath}\" >nul 2>&1",
            "del /f /q \"%~f0\" >nul 2>&1");
        File.WriteAllText(updaterPath, script, Encoding.ASCII);

        sb.AppendLine("In-place update prepared.");
        sb.AppendLine($"Staged binary: {stagingPath}");
        sb.AppendLine($"Updater script: {updaterPath}");
        sb.AppendLine("STATUS: READY_FOR_RESTART");
        return sb.ToString();
    }
    public Task<string> AnalyzeAsync()
    {
        return Task.Run(() =>
        {
            var sb = new StringBuilder();
            sb.AppendLine("Ashera's Tweaking Utility Analysis");
            sb.AppendLine(new string('=', 70));
            sb.AppendLine($"Time: {DateTime.Now}");
            sb.AppendLine($"Running as admin: {IsRunningAsAdmin()}");
            sb.AppendLine();

            var activeScheme = GetActivePowerSchemeGuid();
            sb.AppendLine($"Active power scheme GUID: {activeScheme ?? "Unknown"}");

            var gameMode = GetRegistryDword(Registry.CurrentUser, GameBarKey, GameBarValueName);
            sb.AppendLine($"Game Mode (AutoGameModeEnabled): {(gameMode is null ? "Not set" : gameMode)}");
            sb.AppendLine();

            sb.AppendLine("Startup apps (quick count):");
            sb.AppendLine($"- HKCU Run: {GetRunKeyCount(Registry.CurrentUser)}");
            sb.AppendLine($"- HKLM Run: {GetRunKeyCount(Registry.LocalMachine)}");
            sb.AppendLine($"- Startup folder items: {GetStartupFolderItemCount()}");
            sb.AppendLine();

            sb.AppendLine("Top memory processes:");
            foreach (var line in GetTopMemoryProcesses(10))
            {
                sb.AppendLine($"- {line}");
            }

            sb.AppendLine();
            sb.AppendLine("Saved rollback state:");
            sb.AppendLine(File.Exists(StateFilePath)
                ? $"- Found: {StateFilePath}"
                : "- None found");

            return sb.ToString();
        });
    }

    public Task<string> ApplyGamingProfileAsync()
    {
        return Task.Run(() =>
        {
            var sb = new StringBuilder();
            sb.AppendLine("Applying Safe Gaming Profile");
            sb.AppendLine(new string('=', 70));

            var state = new OptimizationState
            {
                SavedAtUtc = DateTime.UtcNow,
                PreviousPowerSchemeGuid = GetActivePowerSchemeGuid(),
                PreviousAutoGameModeEnabled = GetRegistryDword(Registry.CurrentUser, GameBarKey, GameBarValueName)
            };

            SaveState(state);
            sb.AppendLine($"Saved rollback state to: {StateFilePath}");

            var highPerformanceGuid = GetHighPerformancePowerSchemeGuid();
            if (!string.IsNullOrWhiteSpace(highPerformanceGuid))
            {
                var powerResult = RunProcess("powercfg", $"/SETACTIVE {highPerformanceGuid}");
                sb.AppendLine($"Set power plan: {powerResult}");
            }
            else
            {
                sb.AppendLine("Set power plan: skipped (high performance plan not found)");
            }

            SetRegistryDword(Registry.CurrentUser, GameBarKey, GameBarValueName, 1);
            sb.AppendLine("Set Game Mode: AutoGameModeEnabled=1");

            sb.AppendLine();
            sb.AppendLine("Done. If a setting fails, run this app as Administrator.");

            return sb.ToString();
        });
    }

    public Task<string> RollbackAsync()
    {
        return Task.Run(() =>
        {
            var sb = new StringBuilder();
            sb.AppendLine("Rollback Safe Gaming Profile");
            sb.AppendLine(new string('=', 70));

            var state = LoadState();
            if (state is null)
            {
                sb.AppendLine("No rollback state found.");
                sb.AppendLine($"Expected: {StateFilePath}");
                return sb.ToString();
            }

            sb.AppendLine($"Loaded state from: {StateFilePath}");
            sb.AppendLine($"State timestamp (UTC): {state.SavedAtUtc:u}");

            if (!string.IsNullOrWhiteSpace(state.PreviousPowerSchemeGuid))
            {
                var powerResult = RunProcess("powercfg", $"/SETACTIVE {state.PreviousPowerSchemeGuid}");
                sb.AppendLine($"Restore power plan: {powerResult}");
            }
            else
            {
                sb.AppendLine("Restore power plan: skipped (no saved value)");
            }

            if (state.PreviousAutoGameModeEnabled.HasValue)
            {
                SetRegistryDword(Registry.CurrentUser, GameBarKey, GameBarValueName, state.PreviousAutoGameModeEnabled.Value);
                sb.AppendLine($"Restore Game Mode: AutoGameModeEnabled={state.PreviousAutoGameModeEnabled.Value}");
            }
            else
            {
                sb.AppendLine("Restore Game Mode: skipped (no saved value)");
            }

            sb.AppendLine();
            sb.AppendLine("Rollback finished.");
            return sb.ToString();
        });
    }

    public Task<string> GetResourceReportAsync()
    {
        return Task.Run(() =>
        {
            var sb = new StringBuilder();
            sb.AppendLine("System Resource Report");
            sb.AppendLine(new string('=', 70));
            sb.AppendLine($"Timestamp: {DateTime.Now}");
            sb.AppendLine($"Machine: {Environment.MachineName}");
            sb.AppendLine($"OS: {Environment.OSVersion}");
            sb.AppendLine($"64-bit OS: {Environment.Is64BitOperatingSystem}");
            sb.AppendLine($"Logical CPU cores: {Environment.ProcessorCount}");
            sb.AppendLine();

            using var uptimeCounter = new PerformanceCounter("System", "System Up Time");
            _ = uptimeCounter.NextValue();
            Thread.Sleep(100);
            var uptime = TimeSpan.FromSeconds(uptimeCounter.NextValue());
            sb.AppendLine($"Uptime: {uptime:dd\\.hh\\:mm\\:ss}");
            sb.AppendLine();

            sb.AppendLine("Top CPU processes:");
            foreach (var p in Process.GetProcesses()
                         .OrderByDescending(x => SafeGet(() => x.TotalProcessorTime.TotalSeconds))
                         .Take(10))
            {
                sb.AppendLine($"- {p.ProcessName} (PID {p.Id})");
            }

            sb.AppendLine();
            sb.AppendLine("Top memory processes:");
            foreach (var line in GetTopMemoryProcesses(10))
            {
                sb.AppendLine($"- {line}");
            }

            return sb.ToString();
        });
    }

    public Task<string> CreateRestorePointAsync()
    {
        return Task.Run(() =>
        {
            var sb = new StringBuilder();
            sb.AppendLine("System Restore");
            sb.AppendLine(new string('=', 70));
            sb.AppendLine($"Timestamp: {DateTime.Now}");
            sb.AppendLine($"Running as admin: {IsRunningAsAdmin()}");
            sb.AppendLine();

            if (!IsRunningAsAdmin())
            {
                sb.AppendLine("Restore point requires Administrator privileges.");
                sb.AppendLine("Run this app as Administrator and retry.");
                return sb.ToString();
            }

            var name = $"AsherasTweakingUtility-{DateTime.Now:yyyyMMdd-HHmmss}";
            var cmd =
                $"-NoProfile -ExecutionPolicy Bypass -Command \"Checkpoint-Computer -Description '{name}' -RestorePointType 'MODIFY_SETTINGS'\"";
            var result = RunProcess("powershell.exe", cmd);
            sb.AppendLine("Checkpoint-Computer result:");
            sb.AppendLine(result);
            sb.AppendLine();
            sb.AppendLine("If restore points are disabled, enable System Protection for your drive and retry.");
            return sb.ToString();
        });
    }

    public Task<string> RunQuickFixesAsync()
    {
        return Task.Run(() =>
        {
            var sb = new StringBuilder();
            sb.AppendLine("Quick Fixes");
            sb.AppendLine(new string('=', 70));
            sb.AppendLine($"Timestamp: {DateTime.Now}");
            sb.AppendLine();

            var flush = RunProcess("ipconfig", "/flushdns");
            sb.AppendLine("DNS cache flush:");
            sb.AppendLine(flush);
            sb.AppendLine();

            var tempPath = Path.GetTempPath();
            var deletedFiles = 0;
            var deletedDirs = 0;
            var skipped = 0;

            foreach (var file in Directory.EnumerateFiles(tempPath, "*", SearchOption.TopDirectoryOnly))
            {
                try
                {
                    File.Delete(file);
                    deletedFiles++;
                }
                catch
                {
                    skipped++;
                }
            }

            foreach (var dir in Directory.EnumerateDirectories(tempPath, "*", SearchOption.TopDirectoryOnly))
            {
                try
                {
                    Directory.Delete(dir, true);
                    deletedDirs++;
                }
                catch
                {
                    skipped++;
                }
            }

            sb.AppendLine("Temp cleanup:");
            sb.AppendLine($"- Path: {tempPath}");
            sb.AppendLine($"- Deleted files: {deletedFiles}");
            sb.AppendLine($"- Deleted folders: {deletedDirs}");
            sb.AppendLine($"- Skipped (in use/protected): {skipped}");
            sb.AppendLine();

            sb.AppendLine("Quick fixes finished.");
            return sb.ToString();
        });
    }

    public Task<string> GetStartupOptimizationReportAsync()
    {
        return Task.Run(() =>
        {
            var sb = new StringBuilder();
            sb.AppendLine("Startup Optimizer");
            sb.AppendLine(new string('=', 70));
            sb.AppendLine("Review high-impact startup apps and disable only non-essential entries.");
            sb.AppendLine();

            sb.AppendLine("HKCU Run entries:");
            foreach (var line in EnumerateRunEntries(Registry.CurrentUser))
            {
                sb.AppendLine($"- {line}");
            }

            sb.AppendLine();
            sb.AppendLine("HKLM Run entries:");
            foreach (var line in EnumerateRunEntries(Registry.LocalMachine))
            {
                sb.AppendLine($"- {line}");
            }

            sb.AppendLine();
            var startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            sb.AppendLine($"Startup folder: {startupPath}");
            if (Directory.Exists(startupPath))
            {
                foreach (var file in Directory.EnumerateFiles(startupPath))
                {
                    sb.AppendLine($"- {Path.GetFileName(file)}");
                }
            }

            sb.AppendLine();
            sb.AppendLine("Tip: Use Task Manager > Startup Apps to disable entries safely.");
            return sb.ToString();
        });
    }

    public Task<string> GetAppOptimizationReportAsync()
    {
        return Task.Run(() =>
        {
            var sb = new StringBuilder();
            sb.AppendLine("App Optimizer");
            sb.AppendLine(new string('=', 70));
            sb.AppendLine("Suggestions are informational and non-destructive.");
            sb.AppendLine();

            var processes = Process.GetProcesses().ToList();
            sb.AppendLine($"Running process count: {processes.Count}");
            sb.AppendLine();

            sb.AppendLine("Top memory usage:");
            foreach (var p in processes
                         .OrderByDescending(x => SafeGet(() => x.WorkingSet64))
                         .Take(12))
            {
                var mb = SafeGet(() => p.WorkingSet64) / (1024 * 1024);
                sb.AppendLine($"- {p.ProcessName} (PID {p.Id}) - {mb} MB");
            }

            sb.AppendLine();
            sb.AppendLine("Quick actions:");
            sb.AppendLine("- Close browser tabs/windows you are not using.");
            sb.AppendLine("- Disable overlays you do not need (Discord, GeForce, Xbox Game Bar).");
            sb.AppendLine("- Keep antivirus and drivers enabled.");

            return sb.ToString();
        });
    }

    public Task<string> ApplyTweakToggleAsync(string tweakKey, bool enabled)
    {
        return Task.Run(() =>
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Tweak Toggle: {tweakKey}");
            sb.AppendLine(new string('=', 70));
            sb.AppendLine($"State: {(enabled ? "ON" : "OFF")}");
            sb.AppendLine($"Time: {DateTime.Now}");
            sb.AppendLine();

            try
            {
                switch (tweakKey)
                {
                    case "sensor_suite_off":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required.");
                        }
                        else
                        {
                            SetServiceStartMode("SensorService", enabled ? 4 : 3);
                            SetServiceStartMode("lfsvc", enabled ? 4 : 3);
                            sb.AppendLine("Sensor services updated.");
                        }
                        break;

                    case "sticky_keys_guard":
                        SetRegistryString(
                            Registry.CurrentUser,
                            @"Control Panel\Accessibility\StickyKeys",
                            "Flags",
                            enabled ? "506" : "510");
                        sb.AppendLine("StickyKeys flags updated.");
                        break;

                    case "telemetry_zero":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required to edit telemetry policy under HKLM.");
                        }
                        else
                        {
                            SetRegistryDword(
                                Registry.LocalMachine,
                                @"SOFTWARE\Policies\Microsoft\Windows\DataCollection",
                                "AllowTelemetry",
                                enabled ? 0 : 1);
                            sb.AppendLine("Telemetry policy updated.");
                        }
                        break;

                    case "voice_activation_off":
                        SetRegistryDword(
                            Registry.CurrentUser,
                            @"Software\Microsoft\Speech_OneCore\Settings\VoiceActivation",
                            "AgentActivationEnabled",
                            enabled ? 0 : 1);
                        SetRegistryDword(
                            Registry.CurrentUser,
                            @"Software\Microsoft\Speech_OneCore\Settings\VoiceActivation\UserPreferenceForAllApps",
                            "AgentActivationOnLockScreenEnabled",
                            enabled ? 0 : 1);
                        sb.AppendLine("Voice activation settings updated.");
                        break;

                    case "webdav_scan_off":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required to change WebClient service startup.");
                        }
                        else
                        {
                            var startType = enabled ? "disabled" : "demand";
                            sb.AppendLine(RunProcess("sc.exe", $"config WebClient start= {startType}"));
                        }
                        break;

                    case "wu_core_off":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required to control Windows Update services.");
                        }
                        else
                        {
                            var startType = enabled ? "disabled" : "demand";
                            sb.AppendLine(RunProcess("sc.exe", $"config wuauserv start= {startType}"));
                            sb.AppendLine(RunProcess("sc.exe", $"config UsoSvc start= {startType}"));
                            sb.AppendLine(RunProcess("sc.exe", $"config WaaSMedicSvc start= {startType}"));
                        }
                        break;

                    case "whql_only":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required.");
                        }
                        else
                        {
                            SetRegistryDword(
                                Registry.LocalMachine,
                                @"SOFTWARE\Microsoft\Windows\CurrentVersion\DriverSearching",
                                "SearchOrderConfig",
                                enabled ? 0 : 1);
                            sb.AppendLine("Driver search policy updated.");
                        }
                        break;

                    case "wu_control":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required.");
                        }
                        else
                        {
                            SetRegistryDword(
                                Registry.LocalMachine,
                                @"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU",
                                "NoAutoUpdate",
                                enabled ? 1 : 0);
                            sb.AppendLine("Windows Update auto-update policy updated.");
                        }
                        break;

                    case "cloud_sync_off":
                        SetRegistryDword(
                            Registry.CurrentUser,
                            @"Software\Microsoft\Windows\CurrentVersion\SettingSync",
                            "SyncPolicy",
                            enabled ? 5 : 0);
                        sb.AppendLine("Cloud sync policy updated.");
                        break;

                    case "do_solo_mode":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required.");
                        }
                        else
                        {
                            SetRegistryDword(
                                Registry.LocalMachine,
                                @"SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization",
                                "DODownloadMode",
                                enabled ? 0 : 1);
                            sb.AppendLine("Delivery Optimization mode updated.");
                        }
                        break;

                    case "process_count_reduction":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required.");
                        }
                        else
                        {
                            SetRegistryDword(
                                Registry.LocalMachine,
                                @"SYSTEM\CurrentControlSet\Control",
                                "SvcHostSplitThresholdInKB",
                                enabled ? 3670016 : 0);
                            sb.AppendLine("Process count reduction policy updated.");
                        }
                        break;

                    case "amd_chill_off":
                    case "amd_power_hold":
                    case "amd_service_trim":
                        sb.AppendLine(ToggleAmdServices(enabled));
                        SetManagedState(tweakKey, enabled);
                        break;

                    case "nudge_blocker":
                        SetRegistryDword(
                            Registry.CurrentUser,
                            @"Software\Microsoft\Windows\CurrentVersion\UserProfileEngagement",
                            "ScoobeSystemSettingEnabled",
                            enabled ? 0 : 1);
                        SetRegistryDword(
                            Registry.CurrentUser,
                            @"Software\Microsoft\Siuf\Rules",
                            "NumberOfSIUFInPeriod",
                            enabled ? 0 : 1);
                        SetRegistryDword(
                            Registry.CurrentUser,
                            @"Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager",
                            "SubscribedContent-338389Enabled",
                            enabled ? 0 : 1);
                        sb.AppendLine("Nudges and prompts policy updated.");
                        break;

                    case "network_driver_optimize":
                        sb.AppendLine(ToggleNetworkDriverOptimization(enabled));
                        SetManagedState(tweakKey, enabled);
                        sb.AppendLine("Restart required to fully apply NIC driver-level changes.");
                        break;

                    case "ultimate_power_plan":
                        sb.AppendLine(ApplyUltimatePowerPlan(enabled));
                        break;

                    case "memory_integrity_off":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required.");
                        }
                        else
                        {
                            SetRegistryDword(
                                Registry.LocalMachine,
                                @"SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity",
                                "Enabled",
                                enabled ? 0 : 1);
                            sb.AppendLine("Memory Integrity setting updated. Restart required.");
                        }
                        break;

                    case "gpu_msi_mode":
                        sb.AppendLine(SetGpuMsiMode(enabled));
                        break;

                    case "hpet_tune_off":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required.");
                        }
                        else
                        {
                            var cmd = enabled
                                ? "/deletevalue useplatformclock"
                                : "/set useplatformclock true";
                            sb.AppendLine(RunProcess("bcdedit", cmd));
                            sb.AppendLine("HPET/clock policy updated. Restart required.");
                        }
                        break;

                    case "startup_cleanup_assist":
                        _ = RunProcess("explorer.exe", "ms-settings:startupapps");
                        SetManagedState(tweakKey, enabled);
                        sb.AppendLine("Opened Startup Apps settings for manual cleanup.");
                        break;

                    case "nvidia_latency_profile":
                        sb.AppendLine(ApplyNvidiaLatencyProfile(enabled));
                        SetManagedState(tweakKey, enabled);
                        break;

                    case "wu_tournament_mode":
                        sb.AppendLine(RunProcess("net", enabled ? "stop wuauserv" : "start wuauserv"));
                        sb.AppendLine("Windows Update service toggled temporarily.");
                        break;

                    case "sysmain_off":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required.");
                        }
                        else
                        {
                            sb.AppendLine(RunProcess("sc.exe", $"config SysMain start= {(enabled ? "disabled" : "auto")}"));
                            sb.AppendLine(RunProcess("net", enabled ? "stop SysMain" : "start SysMain"));
                        }
                        break;

                    case "search_indexing_off":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required.");
                        }
                        else
                        {
                            sb.AppendLine(RunProcess("sc.exe", $"config WSearch start= {(enabled ? "disabled" : "delayed-auto")}"));
                            sb.AppendLine(RunProcess("net", enabled ? "stop WSearch" : "start WSearch"));
                        }
                        break;

                    case "selective_background_services_off":
                        sb.AppendLine(ToggleSelectiveBackgroundServices(enabled));
                        SetManagedState(tweakKey, enabled);
                        break;

                    case "competitive_service_trim":
                        sb.AppendLine(ToggleCompetitiveServiceTrim(enabled, hardcore: false));
                        SetManagedState(tweakKey, enabled);
                        break;

                    case "hardcore_service_trim":
                        sb.AppendLine(ToggleCompetitiveServiceTrim(enabled, hardcore: true));
                        SetManagedState(tweakKey, enabled);
                        break;

                    case "network_hardcore_mode":
                        sb.AppendLine(ToggleNetworkDriverOptimization(enabled));
                        if (enabled)
                        {
                            sb.AppendLine(RunProcess("netsh", "winsock reset"));
                            sb.AppendLine(RunProcess("netsh", "int ip reset"));
                            sb.AppendLine(RunProcess("ipconfig", "/flushdns"));
                            sb.AppendLine(RunProcess("ipconfig", "/release"));
                            sb.AppendLine(RunProcess("ipconfig", "/renew"));
                            sb.AppendLine("Network hardcore sequence applied. Restart recommended.");
                        }
                        else
                        {
                            sb.AppendLine("Network hardcore disabled. Adapter settings returned to safer defaults where supported.");
                        }
                        SetManagedState(tweakKey, enabled);
                        break;

                    case "power_hardcore_mode":
                        sb.AppendLine(ApplyPowerHardcoreMode(enabled));
                        break;

                    case "timer_resolution_mode":
                        sb.AppendLine(ToggleTimerResolutionMode(enabled));
                        SetManagedState(tweakKey, enabled);
                        break;

                    case "background_apps_off":
                        SetRegistryDword(
                            Registry.CurrentUser,
                            @"Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications",
                            "GlobalUserDisabled",
                            enabled ? 1 : 0);
                        sb.AppendLine("Background apps permission policy updated.");
                        break;

                    case "low_latency_mode":
                        SetRegistryDword(
                            Registry.CurrentUser,
                            @"System\GameConfigStore",
                            "GameDVR_Enabled",
                            enabled ? 0 : 1);
                        SetRegistryDword(
                            Registry.CurrentUser,
                            @"SOFTWARE\Microsoft\Windows\CurrentVersion\GameDVR",
                            "AppCaptureEnabled",
                            enabled ? 0 : 1);

                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("User-level latency settings applied.");
                            sb.AppendLine("Admin required for full scheduler/network latency profile.");
                        }
                        else
                        {
                            SetRegistryDword(
                                Registry.LocalMachine,
                                @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile",
                                "NetworkThrottlingIndex",
                                enabled ? unchecked((int)0xFFFFFFFF) : 10);
                            SetRegistryDword(
                                Registry.LocalMachine,
                                @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile",
                                "SystemResponsiveness",
                                enabled ? 0 : 20);
                            SetRegistryString(
                                Registry.LocalMachine,
                                @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games",
                                "Scheduling Category",
                                enabled ? "High" : "Medium");
                            SetRegistryString(
                                Registry.LocalMachine,
                                @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games",
                                "Priority",
                                enabled ? "6" : "2");
                            SetRegistryString(
                                Registry.LocalMachine,
                                @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games",
                                "SFIO Priority",
                                enabled ? "High" : "Normal");
                            sb.AppendLine("Full low-latency profile applied.");
                        }

                        sb.AppendLine("Restart game(s) to apply runtime-side effects.");
                        break;

                    case "hibernation_off":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required.");
                        }
                        else
                        {
                            sb.AppendLine(RunProcess("powercfg", enabled ? "/hibernate off" : "/hibernate on"));
                        }
                        break;

                    case "gamebar_overlay_off":
                        SetRegistryDword(Registry.CurrentUser, @"Software\Microsoft\GameBar", "ShowStartupPanel", enabled ? 0 : 1);
                        SetRegistryDword(Registry.CurrentUser, @"SOFTWARE\Microsoft\Windows\CurrentVersion\GameDVR", "AppCaptureEnabled", enabled ? 0 : 1);
                        SetRegistryDword(Registry.CurrentUser, @"System\GameConfigStore", "GameDVR_Enabled", enabled ? 0 : 1);
                        sb.AppendLine("Game Bar / DVR overlay settings updated.");
                        break;

                    case "widgets_feed_off":
                        SetRegistryDword(Registry.CurrentUser, @"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "TaskbarDa", enabled ? 0 : 1);
                        sb.AppendLine("Widgets feed setting updated.");
                        break;

                    case "ads_id_off":
                        SetRegistryDword(Registry.CurrentUser, @"Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo", "Enabled", enabled ? 0 : 1);
                        sb.AppendLine("Advertising ID setting updated.");
                        break;

                    case "search_highlights_off":
                        SetRegistryDword(Registry.CurrentUser, @"Software\Microsoft\Windows\CurrentVersion\SearchSettings", "IsDynamicSearchBoxEnabled", enabled ? 0 : 1);
                        sb.AppendLine("Search highlights setting updated.");
                        break;

                    case "remote_assistance_off":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required.");
                        }
                        else
                        {
                            SetRegistryDword(Registry.LocalMachine, @"SYSTEM\CurrentControlSet\Control\Remote Assistance", "fAllowToGetHelp", enabled ? 0 : 1);
                            sb.AppendLine("Remote Assistance policy updated.");
                        }
                        break;

                    case "fast_startup_off":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required.");
                        }
                        else
                        {
                            SetRegistryDword(Registry.LocalMachine, @"SYSTEM\CurrentControlSet\Control\Session Manager\Power", "HiberbootEnabled", enabled ? 0 : 1);
                            sb.AppendLine("Fast Startup setting updated.");
                        }
                        break;

                    case "print_spooler_off":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required.");
                        }
                        else
                        {
                            SetServiceStartMode("Spooler", enabled ? 4 : 2);
                            sb.AppendLine("Print Spooler startup updated.");
                        }
                        break;

                    case "smb1_off":
                        if (!IsRunningAsAdmin())
                        {
                            sb.AppendLine("Admin required.");
                        }
                        else
                        {
                            sb.AppendLine(RunProcess("dism.exe", enabled
                                ? "/online /disable-feature /featurename:SMB1Protocol /NoRestart"
                                : "/online /enable-feature /featurename:SMB1Protocol /NoRestart"));
                        }
                        break;

                    default:
                        sb.AppendLine("This toggle is currently mapped to status tracking only.");
                        break;
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"Toggle action failed: {ex.Message}");
            }

            return sb.ToString();
        });
    }

    public Task<bool?> GetTweakStateAsync(string tweakKey)
    {
        return Task.Run(() =>
        {
            return tweakKey switch
            {
                "sensor_suite_off" => GetAnyServiceStates(["SensorService", "lfsvc"], 4),
                "sticky_keys_guard" => GetRegistryString(Registry.CurrentUser, @"Control Panel\Accessibility\StickyKeys", "Flags") == "506",
                "telemetry_zero" => GetRegistryDword(Registry.LocalMachine, @"SOFTWARE\Policies\Microsoft\Windows\DataCollection", "AllowTelemetry") == 0,
                "voice_activation_off" => GetRegistryDword(Registry.CurrentUser, @"Software\Microsoft\Speech_OneCore\Settings\VoiceActivation", "AgentActivationEnabled") == 0,
                "webdav_scan_off" => GetServiceStartMode("WebClient") == 4,
                "whql_only" => GetRegistryDword(Registry.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\DriverSearching", "SearchOrderConfig") == 0,
                "wu_control" => GetRegistryDword(Registry.LocalMachine, @"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", "NoAutoUpdate") == 1,
                "wu_core_off" => GetAllServiceStates(["wuauserv", "UsoSvc", "WaaSMedicSvc"], 4),
                "cloud_sync_off" => GetRegistryDword(Registry.CurrentUser, @"Software\Microsoft\Windows\CurrentVersion\SettingSync", "SyncPolicy") == 5,
                "do_solo_mode" => GetRegistryDword(Registry.LocalMachine, @"SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization", "DODownloadMode") == 0,
                "process_count_reduction" => GetRegistryDword(Registry.LocalMachine, @"SYSTEM\CurrentControlSet\Control", "SvcHostSplitThresholdInKB") == 3670016,
                "amd_chill_off" => GetManagedState(tweakKey),
                "amd_power_hold" => GetManagedState(tweakKey),
                "amd_service_trim" => GetManagedState(tweakKey),
                "nudge_blocker" => GetRegistryDword(Registry.CurrentUser, @"Software\Microsoft\Windows\CurrentVersion\UserProfileEngagement", "ScoobeSystemSettingEnabled") == 0,
                "network_driver_optimize" => GetManagedState(tweakKey),
                "ultimate_power_plan" => IsUltimatePlanActive(),
                "memory_integrity_off" => GetRegistryDword(Registry.LocalMachine, @"SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity", "Enabled") == 0,
                "gpu_msi_mode" => IsGpuMsiEnabled(),
                "hpet_tune_off" => IsHpetTuneOff(),
                "startup_cleanup_assist" => GetManagedState(tweakKey),
                "nvidia_latency_profile" => GetManagedState(tweakKey),
                "wu_tournament_mode" => IsServiceStopped("wuauserv"),
                "sysmain_off" => GetServiceStartMode("SysMain") == 4,
                "search_indexing_off" => GetServiceStartMode("WSearch") == 4,
                "selective_background_services_off" => GetManagedState(tweakKey),
                "competitive_service_trim" => GetManagedState(tweakKey),
                "hardcore_service_trim" => GetManagedState(tweakKey),
                "network_hardcore_mode" => GetManagedState(tweakKey),
                "power_hardcore_mode" => IsPowerHardcoreModeEnabled(),
                "timer_resolution_mode" => GetManagedState(tweakKey),
                "background_apps_off" => GetRegistryDword(Registry.CurrentUser, @"Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications", "GlobalUserDisabled") == 1,
                "low_latency_mode" =>
                    GetRegistryDword(Registry.CurrentUser, @"System\GameConfigStore", "GameDVR_Enabled") == 0 &&
                    GetRegistryDword(Registry.CurrentUser, @"SOFTWARE\Microsoft\Windows\CurrentVersion\GameDVR", "AppCaptureEnabled") == 0 &&
                    GetRegistryDword(Registry.LocalMachine, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile", "SystemResponsiveness") == 0 &&
                    GetRegistryDword(Registry.LocalMachine, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile", "NetworkThrottlingIndex") == unchecked((int)0xFFFFFFFF),
                "hibernation_off" => IsHibernateDisabled(),
                "gamebar_overlay_off" => GetRegistryDword(Registry.CurrentUser, @"SOFTWARE\Microsoft\Windows\CurrentVersion\GameDVR", "AppCaptureEnabled") == 0,
                "widgets_feed_off" => GetRegistryDword(Registry.CurrentUser, @"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "TaskbarDa") == 0,
                "ads_id_off" => GetRegistryDword(Registry.CurrentUser, @"Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo", "Enabled") == 0,
                "search_highlights_off" => GetRegistryDword(Registry.CurrentUser, @"Software\Microsoft\Windows\CurrentVersion\SearchSettings", "IsDynamicSearchBoxEnabled") == 0,
                "remote_assistance_off" => GetRegistryDword(Registry.LocalMachine, @"SYSTEM\CurrentControlSet\Control\Remote Assistance", "fAllowToGetHelp") == 0,
                "fast_startup_off" => GetRegistryDword(Registry.LocalMachine, @"SYSTEM\CurrentControlSet\Control\Session Manager\Power", "HiberbootEnabled") == 0,
                "print_spooler_off" => GetServiceStartMode("Spooler") == 4,
                "smb1_off" => GetRegistryDword(Registry.LocalMachine, @"SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters", "SMB1") == 0,
                _ => (bool?)null
            };
        });
    }

    private static bool IsRunningAsAdmin()
    {
        using var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    private static IEnumerable<string> GetTopMemoryProcesses(int count)
    {
        return Process.GetProcesses()
            .OrderByDescending(p =>
            {
                try
                {
                    return p.WorkingSet64;
                }
                catch
                {
                    return 0;
                }
            })
            .Take(count)
            .Select(p =>
            {
                var mb = 0L;
                try
                {
                    mb = p.WorkingSet64 / (1024 * 1024);
                }
                catch
                {
                    // Ignore inaccessible process memory details.
                }

                return $"{p.ProcessName} (PID {p.Id}) - {mb} MB";
            });
    }

    private static int GetRunKeyCount(RegistryKey hive)
    {
        using var runKey = hive.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", false);
        return runKey?.ValueCount ?? 0;
    }

    private static IEnumerable<string> EnumerateRunEntries(RegistryKey hive)
    {
        using var runKey = hive.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", false);
        if (runKey is null)
        {
            yield break;
        }

        foreach (var name in runKey.GetValueNames())
        {
            var value = runKey.GetValue(name)?.ToString() ?? "<null>";
            yield return $"{name} => {value}";
        }
    }

    private static int GetStartupFolderItemCount()
    {
        var startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
        if (!Directory.Exists(startupPath))
        {
            return 0;
        }

        return Directory.GetFiles(startupPath).Length;
    }

    private static string? GetActivePowerSchemeGuid()
    {
        var output = RunProcessRaw("powercfg", "/GETACTIVESCHEME");
        var match = Regex.Match(output, "([A-Fa-f0-9-]{36})");
        return match.Success ? match.Value : null;
    }

    private static string? GetHighPerformancePowerSchemeGuid()
    {
        var output = RunProcessRaw("powercfg", "/LIST");

        foreach (var line in output.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries))
        {
            if (!line.Contains("High performance", StringComparison.OrdinalIgnoreCase) &&
                !line.Contains("Ultimate Performance", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var match = Regex.Match(line, "([A-Fa-f0-9-]{36})");
            if (match.Success)
            {
                return match.Value;
            }
        }

        return null;
    }

    private static int? GetRegistryDword(RegistryKey hive, string subKey, string valueName)
    {
        using var key = hive.OpenSubKey(subKey, false);
        return key?.GetValue(valueName) as int?;
    }

    private static string? GetRegistryString(RegistryKey hive, string subKey, string valueName)
    {
        using var key = hive.OpenSubKey(subKey, false);
        return key?.GetValue(valueName)?.ToString();
    }

    private static void SetRegistryDword(RegistryKey hive, string subKey, string valueName, int value)
    {
        using var key = hive.CreateSubKey(subKey, true);
        key?.SetValue(valueName, value, RegistryValueKind.DWord);
    }

    private static void SetRegistryString(RegistryKey hive, string subKey, string valueName, string value)
    {
        using var key = hive.CreateSubKey(subKey, true);
        key?.SetValue(valueName, value, RegistryValueKind.String);
    }

    private static int? GetServiceStartMode(string serviceName)
    {
        using var key = Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Services\{serviceName}", false);
        return key?.GetValue("Start") as int?;
    }

    private static void SetServiceStartMode(string serviceName, int mode)
    {
        using var key = Registry.LocalMachine.CreateSubKey($@"SYSTEM\CurrentControlSet\Services\{serviceName}", true);
        key?.SetValue("Start", mode, RegistryValueKind.DWord);
    }

    private static bool? GetAllServiceStates(string[] serviceNames, int expectedStartMode)
    {
        foreach (var service in serviceNames)
        {
            var mode = GetServiceStartMode(service);
            if (mode is null)
            {
                return null;
            }

            if (mode.Value != expectedStartMode)
            {
                return false;
            }
        }

        return true;
    }

    private static bool? GetAnyServiceStates(string[] serviceNames, int expectedStartMode)
    {
        var found = false;
        foreach (var service in serviceNames)
        {
            var mode = GetServiceStartMode(service);
            if (mode is null)
            {
                continue;
            }

            found = true;
            if (mode.Value != expectedStartMode)
            {
                return false;
            }
        }

        return found ? true : null;
    }

    private static bool? IsHibernateDisabled()
    {
        try
        {
            var outText = RunProcessRaw("powercfg", "/a");
            if (outText.Contains("Hibernate has not been enabled", StringComparison.OrdinalIgnoreCase) ||
                outText.Contains("Hibernation has not been enabled", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (outText.Contains("Hibernate", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }
        catch
        {
            // ignored
        }

        return null;
    }

    private static string ToggleAmdServices(bool disable)
    {
        if (!IsRunningAsAdmin())
        {
            return "Admin required for AMD service tuning.";
        }

        using var services = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services", false);
        if (services is null)
        {
            return "Services registry not available.";
        }

        var target = disable ? 4 : 3;
        var changed = 0;
        foreach (var name in services.GetSubKeyNames())
        {
            if (!name.Contains("amd", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            try
            {
                SetServiceStartMode(name, target);
                changed++;
            }
            catch
            {
                // Ignore individual service failures.
            }
        }

        return changed == 0 ? "No AMD services found." : $"AMD service start mode updated for {changed} services.";
    }

    private static string ToggleNetworkDriverOptimization(bool enable)
    {
        if (!IsRunningAsAdmin())
        {
            return "Admin required for network driver optimization.";
        }

        var interruptModeration = enable ? "Disabled" : "Enabled";
        var eee = enable ? "Off" : "On";
        var powerToggle = enable ? "Disabled" : "Enabled";
        var script =
            "$adapters = Get-NetAdapter -Physical -ErrorAction SilentlyContinue;" +
            "foreach ($a in $adapters) {" +
            $"try {{ Set-NetAdapterPowerManagement -Name $a.Name -AllowComputerToTurnOffDevice {powerToggle} -ErrorAction Stop }} catch {{ }};" +
            $"try {{ Set-NetAdapterAdvancedProperty -Name $a.Name -DisplayName 'Interrupt Moderation' -DisplayValue '{interruptModeration}' -NoRestart -ErrorAction Stop }} catch {{ }};" +
            $"try {{ Set-NetAdapterAdvancedProperty -Name $a.Name -DisplayName 'Energy-Efficient Ethernet' -DisplayValue '{eee}' -NoRestart -ErrorAction Stop }} catch {{ }};" +
            $"try {{ Set-NetAdapterAdvancedProperty -Name $a.Name -DisplayName 'Speed & Duplex' -DisplayValue '{(enable ? "1.0 Gbps Full Duplex" : "Auto Negotiation")}' -NoRestart -ErrorAction Stop }} catch {{ }};" +
            "}";

        var result = RunProcess("powershell.exe", $"-NoProfile -ExecutionPolicy Bypass -Command \"{script}\"");
        return string.IsNullOrWhiteSpace(result) ? "Network driver optimization command completed." : result;
    }

    private static string ApplyUltimatePowerPlan(bool enable)
    {
        if (!IsRunningAsAdmin())
        {
            return "Admin required.";
        }

        var sb = new StringBuilder();
        if (enable)
        {
            sb.AppendLine(RunProcess("powercfg", "/duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61"));
            var list = RunProcessRaw("powercfg", "/list");
            var ultimateGuid = list
                .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)
                .Where(l => l.Contains("Ultimate Performance", StringComparison.OrdinalIgnoreCase))
                .Select(l => Regex.Match(l, "([A-Fa-f0-9-]{36})"))
                .Where(m => m.Success)
                .Select(m => m.Value)
                .FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(ultimateGuid))
            {
                sb.AppendLine(RunProcess("powercfg", $"/setactive {ultimateGuid}"));
            }
        }

        sb.AppendLine(RunProcess("powercfg", "/setacvalueindex scheme_current 54533251-82be-4824-96c1-47b60b740d00 893dee8e-2bef-41e0-89c6-b55d0929964c 100"));
        sb.AppendLine(RunProcess("powercfg", "/setdcvalueindex scheme_current 54533251-82be-4824-96c1-47b60b740d00 893dee8e-2bef-41e0-89c6-b55d0929964c 100"));
        sb.AppendLine(RunProcess("powercfg", $"/setacvalueindex scheme_current 2a737441-1930-4402-8d77-b2bebba308a3 4f971e89-eebd-4455-a8de-9e59040e7347 {(enable ? 0 : 1)}"));
        sb.AppendLine(RunProcess("powercfg", "/setactive scheme_current"));
        return sb.ToString();
    }

    private static bool? IsUltimatePlanActive()
    {
        try
        {
            var text = RunProcessRaw("powercfg", "/getactivescheme");
            if (text.Contains("Ultimate Performance", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }
        catch
        {
            return null;
        }

        return false;
    }

    private static string SetGpuMsiMode(bool enable)
    {
        if (!IsRunningAsAdmin())
        {
            return "Admin required.";
        }

        var changed = 0;
        try
        {
            using var gpuSearcher = new System.Management.ManagementObjectSearcher("SELECT PNPDeviceID FROM Win32_VideoController");
            foreach (var obj in gpuSearcher.Get().OfType<System.Management.ManagementObject>())
            {
                var pnp = obj["PNPDeviceID"]?.ToString();
                if (string.IsNullOrWhiteSpace(pnp) || !pnp.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var path = $@"SYSTEM\CurrentControlSet\Enum\{pnp}\Device Parameters\Interrupt Management\MessageSignaledInterruptProperties";
                using var key = Registry.LocalMachine.CreateSubKey(path, true);
                if (key is null)
                {
                    continue;
                }

                key.SetValue("MSISupported", enable ? 1 : 0, RegistryValueKind.DWord);
                changed++;
            }
        }
        catch (Exception ex)
        {
            return $"GPU MSI operation failed: {ex.Message}";
        }

        return changed == 0
            ? "No compatible PCI GPU registry path found."
            : $"GPU MSI mode updated for {changed} device(s). Restart required.";
    }

    private static bool? IsGpuMsiEnabled()
    {
        try
        {
            using var gpuSearcher = new System.Management.ManagementObjectSearcher("SELECT PNPDeviceID FROM Win32_VideoController");
            var found = false;
            foreach (var obj in gpuSearcher.Get().OfType<System.Management.ManagementObject>())
            {
                var pnp = obj["PNPDeviceID"]?.ToString();
                if (string.IsNullOrWhiteSpace(pnp) || !pnp.StartsWith("PCI\\", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                found = true;
                using var key = Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Enum\{pnp}\Device Parameters\Interrupt Management\MessageSignaledInterruptProperties", false);
                var val = key?.GetValue("MSISupported") as int?;
                if (val != 1)
                {
                    return false;
                }
            }

            return found ? true : null;
        }
        catch
        {
            return null;
        }
    }

    private static bool? IsHpetTuneOff()
    {
        try
        {
            var outText = RunProcessRaw("bcdedit", "/enum");
            return !Regex.IsMatch(outText, @"useplatformclock\s+Yes", RegexOptions.IgnoreCase);
        }
        catch
        {
            return null;
        }
    }

    private static string ApplyNvidiaLatencyProfile(bool enable)
    {
        var sb = new StringBuilder();
        if (File.Exists(Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\NVIDIA Corporation\NVSMI\nvidia-smi.exe")))
        {
            var exe = Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\NVIDIA Corporation\NVSMI\nvidia-smi.exe");
            sb.AppendLine(RunProcess(exe, $"-pm {(enable ? 1 : 0)}"));
        }
        else
        {
            sb.AppendLine("nvidia-smi not found; applied guidance only.");
        }

        _ = RunProcess("nvcplui.exe", "");
        sb.AppendLine("Open NVIDIA Control Panel and set: Low Latency Mode=On/Ultra, Power Management=Prefer Maximum Performance, Shader Cache=On, V-Sync=Off.");
        return sb.ToString();
    }

    private static string ToggleSelectiveBackgroundServices(bool disable)
    {
        if (!IsRunningAsAdmin())
        {
            return "Admin required.";
        }

        var sb = new StringBuilder();
        var startType = disable ? "disabled" : "demand";
        foreach (var service in new[] { "XboxGipSvc", "XblAuthManager", "XblGameSave", "Fax" })
        {
            sb.AppendLine(RunProcess("sc.exe", $"config {service} start= {startType}"));
            sb.AppendLine(RunProcess("sc.exe", disable ? $"stop {service}" : $"start {service}"));
        }

        sb.AppendLine(RunProcess("sc.exe", $"config Spooler start= {(disable ? "disabled" : "auto")}"));
        if (disable)
        {
            sb.AppendLine(RunProcess("sc.exe", "stop Spooler"));
        }

        sb.AppendLine(RunProcess("taskkill", disable ? "/IM OneDrive.exe /F" : "/IM OneDrive.exe"));
        sb.AppendLine(RunProcess("taskkill", disable ? "/IM PhoneExperienceHost.exe /F" : "/IM PhoneExperienceHost.exe"));
        return sb.ToString();
    }

    private static string ToggleCompetitiveServiceTrim(bool enableTrim, bool hardcore)
    {
        if (!IsRunningAsAdmin())
        {
            return "Admin required.";
        }

        // Curated optional services only. Core services (RPC, DCOM, WMI, NLA, EventLog, etc.) are intentionally excluded.
        var safeServices = new (string Service, int RestoreMode)[]
        {
            ("XblAuthManager", 3),
            ("XblGameSave", 3),
            ("XboxNetApiSvc", 3),
            ("XboxGipSvc", 3),
            ("Fax", 3),
            ("MapsBroker", 3),
            ("PhoneSvc", 3),
            ("WMPNetworkSvc", 3),
            ("WerSvc", 3)
        };

        var hardcoreExtras = new (string Service, int RestoreMode)[]
        {
            ("WpnService", 2),
            ("WpnUserService", 3),
            ("Spooler", 2),
            ("SysMain", 2),
            ("WSearch", 2)
        };

        var sb = new StringBuilder();
        sb.AppendLine(enableTrim
            ? $"Applying {(hardcore ? "hardcore" : "safe")} competitive service trim..."
            : "Reverting competitive service trim to defaults...");

        var services = safeServices.ToList();
        if (hardcore)
        {
            services.AddRange(hardcoreExtras);
        }

        foreach (var entry in services)
        {
            try
            {
                if (enableTrim)
                {
                    SetServiceStartMode(entry.Service, 4);
                    sb.AppendLine(RunProcess("sc.exe", $"stop {entry.Service}"));
                }
                else
                {
                    SetServiceStartMode(entry.Service, entry.RestoreMode);
                    if (entry.RestoreMode == 2)
                    {
                        sb.AppendLine(RunProcess("sc.exe", $"start {entry.Service}"));
                    }
                }
            }
            catch
            {
                sb.AppendLine($"Skipped: {entry.Service} (not present or access denied).");
            }
        }

        if (enableTrim)
        {
            sb.AppendLine(RunProcess("taskkill", "/IM OneDrive.exe /F"));
            sb.AppendLine(RunProcess("taskkill", "/IM PhoneExperienceHost.exe /F"));
            sb.AppendLine(RunProcess("taskkill", "/IM YourPhone.exe /F"));
        }

        sb.AppendLine("Done. Re-enable toggle to restore default startup modes.");
        if (hardcore)
        {
            sb.AppendLine("Warning: Hardcore trim may disable notifications, search indexing, printing, or Xbox features until reverted.");
        }

        return sb.ToString();
    }

    private static string ApplyPowerHardcoreMode(bool enable)
    {
        if (!IsRunningAsAdmin())
        {
            return "Admin required.";
        }

        var sb = new StringBuilder();
        if (enable)
        {
            sb.AppendLine(ApplyUltimatePowerPlan(true));
            sb.AppendLine(RunProcess("powercfg", "/setacvalueindex scheme_current 501a4d13-42af-4429-9fd1-a8218c268e20 ee12f906-d277-404b-b6da-e5fa1a576df5 0"));
            sb.AppendLine(RunProcess("powercfg", "/setdcvalueindex scheme_current 501a4d13-42af-4429-9fd1-a8218c268e20 ee12f906-d277-404b-b6da-e5fa1a576df5 0"));
            sb.AppendLine("Power hardcore mode enabled.");
        }
        else
        {
            sb.AppendLine(RunProcess("powercfg", "/setacvalueindex scheme_current 501a4d13-42af-4429-9fd1-a8218c268e20 ee12f906-d277-404b-b6da-e5fa1a576df5 2"));
            sb.AppendLine(RunProcess("powercfg", "/setdcvalueindex scheme_current 501a4d13-42af-4429-9fd1-a8218c268e20 ee12f906-d277-404b-b6da-e5fa1a576df5 2"));
            sb.AppendLine(ApplyUltimatePowerPlan(false));
            sb.AppendLine("Power hardcore mode disabled.");
        }

        sb.AppendLine(RunProcess("powercfg", "/setactive scheme_current"));
        return sb.ToString();
    }

    private static bool? IsPowerHardcoreModeEnabled()
    {
        try
        {
            var q = RunProcessRaw("powercfg", "/q scheme_current 501a4d13-42af-4429-9fd1-a8218c268e20 ee12f906-d277-404b-b6da-e5fa1a576df5");
            return q.Contains("0x00000000", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return null;
        }
    }

    private static string ToggleTimerResolutionMode(bool enable)
    {
        if (enable)
        {
            if (NtSetTimerResolution(5000, true, out _) == 0)
            {
                return "High timer resolution requested (0.5 ms).";
            }

            return "Failed to request timer resolution.";
        }

        _ = NtSetTimerResolution(5000, false, out _);
        return "Timer resolution request released.";
    }

    private static bool IsServiceStopped(string serviceName)
    {
        var output = RunProcessRaw("sc.exe", $"query {serviceName}");
        return output.Contains("STOPPED", StringComparison.OrdinalIgnoreCase);
    }

    [DllImport("ntdll.dll")]
    private static extern int NtSetTimerResolution(uint desiredResolution, bool setResolution, out uint currentResolution);

    private static string RunProcess(string fileName, string arguments)
    {
        var output = RunProcessRaw(fileName, arguments);
        return string.IsNullOrWhiteSpace(output) ? "OK" : output.Trim();
    }

    private static string RunProcessRaw(string fileName, string arguments)
    {
        using var proc = new Process();
        proc.StartInfo = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        proc.Start();
        var stdout = proc.StandardOutput.ReadToEnd();
        var stderr = proc.StandardError.ReadToEnd();
        proc.WaitForExit();

        var combined = (stdout + Environment.NewLine + stderr).Trim();
        if (proc.ExitCode != 0)
        {
            return $"ExitCode {proc.ExitCode}: {combined}";
        }

        return combined;
    }

    private static T SafeGet<T>(Func<T> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return default!;
        }
    }

    private static void SaveState(OptimizationState state)
    {
        var dir = Path.GetDirectoryName(StateFilePath)!;
        Directory.CreateDirectory(dir);

        var json = JsonSerializer.Serialize(state, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(StateFilePath, json);
    }

    private static bool? GetManagedState(string key)
    {
        if (!File.Exists(ManagedStateFilePath))
        {
            return false;
        }

        try
        {
            var json = File.ReadAllText(ManagedStateFilePath);
            var data = JsonSerializer.Deserialize<Dictionary<string, bool>>(json);
            if (data is null)
            {
                return false;
            }

            return data.TryGetValue(key, out var value) ? value : false;
        }
        catch
        {
            return null;
        }
    }

    private static void SetManagedState(string key, bool value)
    {
        var dir = Path.GetDirectoryName(ManagedStateFilePath)!;
        Directory.CreateDirectory(dir);

        Dictionary<string, bool> data;
        if (File.Exists(ManagedStateFilePath))
        {
            try
            {
                data = JsonSerializer.Deserialize<Dictionary<string, bool>>(File.ReadAllText(ManagedStateFilePath))
                    ?? new Dictionary<string, bool>();
            }
            catch
            {
                data = new Dictionary<string, bool>();
            }
        }
        else
        {
            data = new Dictionary<string, bool>();
        }

        data[key] = value;
        File.WriteAllText(ManagedStateFilePath, JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static OptimizationState? LoadState()
    {
        if (!File.Exists(StateFilePath))
        {
            return null;
        }

        var json = File.ReadAllText(StateFilePath);
        return JsonSerializer.Deserialize<OptimizationState>(json);
    }
    private static HttpClient CreateUpdateClient()
    {
        var client = new HttpClient();
        client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("AsherasTweakingUtility", "1.0"));
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
        return client;
    }

    private static async Task<GitHubContentItem?> GetLatestExecutableAsync()
    {
        try
        {
            using var response = await UpdateClient.GetAsync(LatestExeApiUrl);
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<GitHubContentItem>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
        }
        catch
        {
            return null;
        }
    }

    private sealed class GitHubContentItem
    {
        public string? name { get; init; }
        public long size { get; init; }
        public string? download_url { get; init; }
    }
}


