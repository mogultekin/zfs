# ZFS Shares Report

Export Oracle ZFS Storage Appliance share usage to an Excel workbook with collapsible pool, project, and share rows.

The report generator runs on a Linux, RHEL, or Exadata node and collects readings over SSH from the appliance’s **Solaris x86** shell. **No internet access or additional Python packages are required.**

## Requirements

- Python 3.6 or later and an SSH client on the reporting node.
- SSH access to the ZFS appliance.
- An appliance account allowed to enter `confirm shell` and run `zpool` and `zfs`.

## Quick start

Copy [`zfs_shares_report.py`](zfs_shares_report.py) to your reporting node, then run:

```bash
python3 zfs_shares_report.py
```

Enter the appliance IP address or hostname and SSH username when prompted. SSH then asks for the password.

## Example session

The address, timestamps, and dataset count below are placeholders.

```text
[root@rhel9 ~]# python3 zfs_shares_report.py
ZFS IP address or hostname: ZFS_Appliance_IP
ZFS SSH username: root
Waiting for appliance login prompt (up to 60s)...
(root@ZFS_Appliance_IP) Password:
Received appliance login prompt.
Waiting for raw shell prompt after confirm shell (up to 30s)...
Received raw shell prompt after confirm shell.
Waiting for raw shell readiness marker (up to 30s)...
Received raw shell readiness marker.
Waiting for collector completion marker (up to 600s)...
Received collector completion marker.
Connection to ZFS_Appliance_IP closed.
Created /root/zfs_shares_YYYYMMDD_hhmmss.xlsx (99 datasets)
Raw capture: /root/zfs_shares_YYYYMMDD_hhmmss.json
```

## Output files

| File | Contents |
| --- | --- |
| `zfs_shares_YYYYMMDD_hhmmss.xlsx` | Formatted Excel report with collapsible rows |
| `zfs_shares_YYYYMMDD_hhmmss.json` | Raw readings for inspection or report regeneration |

Files are saved in the current directory by default. Use the **+ / −** outline controls in Excel to expand or collapse projects and shares.

## Report columns

| Column | Description |
| --- | --- |
| Dataset | Full ZFS dataset path |
| Kind | Pool/local, project, filesystem, or volume |
| Used / Used (bytes) | Readable used space and exact bytes |
| Snapshot space / Snapshots (bytes) | Readable snapshot space and exact bytes |
| Snapshots (direct) | Snapshot count for that dataset |
| Compression | Current compression setting |
| Ratio | Compression ratio reported by ZFS |

**Reading the figures:** Parent used space includes its children, so adding all hierarchy rows would double-count usage. Snapshot space and counts apply to each dataset directly. Readable sizes use powers of 1024.
