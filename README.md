# ZFS shares report — version 1.5

Create an offline Excel inventory of ZFS shares from a standalone appliance, simulator, or two-controller clustered system. The report generator runs on Windows or Linux, including RHEL9 and Exadata nodes. The appliance-side collector runs in the Solaris x86 native shell.

All controller addresses, usernames, and output paths are runtime parameters. No environment-specific addresses or accounts are built into the script. Without an output path, timestamped files are saved in the current directory.

## Requirements

- Python 3.6 or later.
- An OpenSSH client available as `ssh` on the command search path.
- SSH connectivity to each requested controller.
- An appliance account permitted to enter `confirm shell` and read `zpool` and `zfs` information.

No internet access, pip packages, Excel installation, or external workbook service is required. The workbook template is embedded. The script does not install software, change appliance configuration, or modify Exadata databases or services.

## Windows

Save the script in your chosen folder. Open Command Prompt or PowerShell there and run:

```cmd
python --version
ssh -V
python zfs_shares_report.py
```

If your Python installation uses the Windows launcher, use `py -3 zfs_shares_report.py` instead. Quote paths containing spaces. Output files are created in the current folder unless `--output` specifies another location.

Windows support from version 1.4 is retained: the script handles terminal control characters in SSH output, displays connection stages, and uses bounded prompt waits.

## Linux, RHEL9, and Exadata

Copy the same script to your chosen folder and run:

```sh
python3 --version
ssh -V
python3 zfs_shares_report.py
```

Use the name or path of your installed Python 3 interpreter if it differs. No separate Linux edition is needed.

## Controller prompts and clustered systems

The script asks for:

1. **Controller 1 address:** required; an IP address or hostname.
2. **Controller 2 address:** optional. For a clustered system, enter the second controller's IP address or hostname. Leave it blank for a standalone appliance or simulator.
3. **SSH username:** shared by both controller connections.

Controllers are contacted sequentially. SSH may ask for each controller's password and host-key confirmation separately. Passwords are handled by SSH and are not stored by the script.

Providing two controller addresses creates **one XLSX workbook with two tabs**, each named from that controller's `uname -n` hostname. A single-controller run creates one hostname-named tab. Long, invalid, or duplicate names are adjusted to fit Excel's 31-character tab-name limit.

Two addresses select a two-controller report; the script does not verify cluster membership. It reports local share datasets visible on each controller independently, without summing or deduplicating across controllers. A controller with no local share datasets gets a tab explaining that result. If collection fails on either requested controller, no combined workbook is written.

## Command-line examples

Replace uppercase placeholders with your parameters. On Windows, use `python` in place of `python3` when appropriate.

```sh
# Clustered system: both controllers in one workbook.
python3 zfs_shares_report.py --host CONTROLLER_1 --host2 CONTROLLER_2 --user SSH_USERNAME

# Standalone appliance or simulator; no second-controller prompt.
python3 zfs_shares_report.py --host CONTROLLER_1 --user SSH_USERNAME --single

# Select a pool and destination; connection details are requested.
python3 zfs_shares_report.py --pool POOL_NAME --output "REPORT_DIRECTORY/report.xlsx"

# Start with all dataset rows expanded.
python3 zfs_shares_report.py --expanded

# Unattended clustered collection with an existing SSH key.
python3 zfs_shares_report.py --host CONTROLLER_1 --host2 CONTROLLER_2 --user SSH_USERNAME --identity "SSH_KEY_PATH" --batch

# Rebuild a capture without connecting to the appliance.
python3 zfs_shares_report.py --from-json "CAPTURE_FILE.json" --output "REBUILT_REPORT.xlsx"

# Display version or options.
python3 zfs_shares_report.py --version
python3 zfs_shares_report.py --help
```

In interactive mode, omitted connection parameters are requested. `--single` suppresses the second-controller question. With `--batch`, `--host` and `--user` are required; omitting `--host2` selects a single controller without prompting. `--from-json` asks for no connection details.

## Output

By default, the current directory receives `zfs_shares_YYYYMMDD_HHMMSS.xlsx` and a matching `.json` capture. Existing output files are not deliberately overwritten. Two-controller JSON captures contain both controllers. Regeneration supports both older single-controller captures and the combined format.

Columns: dataset path, kind, readable used space, exact used bytes, readable snapshot space, exact snapshot bytes, direct snapshot count, compression setting, and compression ratio.

**No sheet has frozen or split panes.** Collapsible outlines remain available. Pool `/local` and project rows are initially visible; share rows are collapsed. Use Excel's outline controls to expand them, or pass `--expanded`. Excel supports up to seven nested outline levels.

## Changes from version 1.4

- Optional second controller address for clustered systems, with `--host2` and `--single` options.
- One workbook containing separate hostname-named tabs for the controllers.
- All frozen and split panes removed; dataset row outlines retained.
- Combined JSON captures and regeneration support.
- An explicit empty-result tab for controllers with no local share datasets.
- Windows/Linux support, offline generation, Solaris-compatible dataset checks, and connection diagnostics retained.

## Metric definitions

- Only filesystems and volumes under each non-system pool's `/local` dataset are included. Pool names containing lowercase `system` are excluded; other pools without `/local` are reported as skipped.
- Parent `used` values already include descendants. Summing all hierarchy rows would double-count space.
- `usedbysnapshots` is the space charged to snapshots of that dataset. It is not a recursive project total or the sum of individual snapshot `used` values.
- Snapshot counts are direct counts for that dataset, not counts across descendants.
- `compression` is the effective setting; `compressratio` is the achieved ratio reported by ZFS.
- Readable sizes use powers of 1024. Exact bytes come from `zfs get -p`. Exact byte values exceeding Excel's 15-digit numeric precision are stored as text; JSON preserves original readings.
- Readings are collected sequentially, including between controllers, and are not an atomic snapshot. Failed commands or missing required properties stop report generation instead of substituting zeros.

## Diagnostics

```sh
python3 zfs_shares_report.py --host CONTROLLER_1 --user SSH_USERNAME --single --probe-native-shell --debug-ssh
```

The probe checks native-shell access and the availability of `zpool` and `zfs`, without creating a workbook. To probe both controllers, replace `--single` with `--host2 CONTROLLER_2`.

The appliance route opens interactive SSH, waits for the appliance prompt, enters `confirm shell`, waits for the raw-shell prompt, and runs the read-only collector. For a target already providing an ordinary SSH shell, use `--shell-mode native`.

By default, login allows 60 seconds, subsequent shell transitions allow 30 seconds each (`--prompt-timeout`), and collection allows 600 seconds (`--timeout`). Waiting messages appear every 15 seconds. `--debug-ssh` displays connection diagnostics and remote output. Prompt timeouts include a raw terminal tail for troubleshooting. Ctrl+C cancels the run and closes the script's SSH session.

## Validation

Version 1.4 successfully collected live appliance data and generated a workbook from Windows. Version 1.5 passed 25 local tests, including single- and dual-controller input, hostname tab naming, empty-controller results, and removal of panes. A generated two-controller workbook was reopened, recalculated, inspected, and visually reviewed. Live dual-controller collection has not yet been verified.
