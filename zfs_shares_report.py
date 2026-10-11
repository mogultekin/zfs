#!/usr/bin/env python3
"""Offline ZFS inventory report for Windows, Linux, RHEL9, and Exadata nodes.
Requires Python 3.6+ and OpenSSH; no internet or additional Python packages.
Windows: python zfs_shares_report.py
Linux: python3 zfs_shares_report.py
Prompts for controller 1, optional controller 2, and a shared SSH username.
For clustered systems enter both controller addresses; otherwise leave 2 blank.
Creates one XLSX with hostname-named tabs, collapsible rows, and no frozen panes.
All controller addresses, credentials, and output paths are supplied at runtime.
Solaris appliance access uses 'confirm shell'; --shell-mode native skips it.
"""
import argparse
import base64
from collections import Counter
import copy
from datetime import datetime, timezone
import io
import json
import math
import os
from pathlib import Path
import re
import queue
import threading
import time
import uuid
import shlex
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

TEMPLATE = "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"
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('x', NS)
def q(name):
    return '{' + NS + '}' + name

PROPERTIES = ('type', 'used', 'usedbysnapshots', 'compression', 'compressratio')
COLLECTOR = r'''set -eu
PATH=/usr/sbin:/usr/bin:/sbin:/bin:$PATH
export PATH
LC_ALL=C
export LC_ALL
command -v zpool >/dev/null
command -v zfs >/dev/null
pools=$(zpool list -H -o name)
datasets=$(zfs list -H -t filesystem,volume -o name)
newline='
'
printf 'BEGIN\t1\n'
controller_hostname=$(uname -n)
printf 'HOSTNAME\t%s\n' "$controller_hostname"
for pool in $pools; do
    case "$pool" in *system*) continue ;; esac
    if [ -n "$SELECTED_POOLS" ]; then
        case " $SELECTED_POOLS " in *" $pool "*) ;; *) continue ;; esac
    fi
    root="$pool/local"
    case "$newline$datasets$newline" in
        *"$newline$root$newline"*) ;;
        *) printf 'SKIP\t%s\tno /local dataset\n' "$pool"; continue ;;
    esac
    printf 'ROOT\t%s\n' "$root"
    data=$(zfs list -H -r -t filesystem,volume -o name "$root")
    printf '%s\n' "$data" | awk '{print "DATASET\t" $0}'
    props=$(zfs get -H -p -r -o name,property,value,source type,used,usedbysnapshots,compression,compressratio "$root")
    printf '%s\n' "$props" | awk '{print "PROP\t" $0}'
    snaps=$(zfs list -H -r -t snapshot -o name "$root")
    if [ -n "$snaps" ]; then
        printf '%s\n' "$snaps" | awk '{print "SNAP\t" $0}'
    fi
done
printf 'END\t1\n'
'''

def parse_capture(text, source, started, finished):
    roots, names, props, snaps = [], set(), {}, set()
    skipped = []
    hostname = None
    began = ended = False
    for line in text.splitlines():
        if line == 'BEGIN\t1':
            if began:
                raise ValueError('Duplicate capture start')
            began = True
            continue
        if not began:
            continue  # An SSH login banner can precede the capture.
        if ended:
            # Appliance shell exit messages can follow the completed capture.
            continue
        if line == 'END\t1':
            ended = True
            continue
        parts = line.split('\t')
        if parts[0] == 'HOSTNAME' and len(parts) == 2:
            if hostname is not None or not parts[1].strip():
                raise ValueError('Invalid controller hostname record')
            hostname = parts[1].strip()
        elif parts[0] == 'SKIP' and len(parts) == 3:
            skipped.append({'pool': parts[1], 'reason': parts[2]})
            print('Skipping %s: %s' % (parts[1], parts[2]), file=sys.stderr)
        elif parts[0] == 'ROOT' and len(parts) == 2:
            roots.append(parts[1])
        elif parts[0] == 'DATASET' and len(parts) == 2:
            if parts[1] in names:
                raise ValueError('Duplicate dataset: ' + parts[1])
            names.add(parts[1])
        elif parts[0] == 'PROP' and len(parts) == 5:
            name, prop, value, origin = parts[1:]
            if '@' not in name:
                entry = props.setdefault(name, {})
                if prop in entry:
                    raise ValueError('Duplicate property for ' + name)
                entry[prop] = {'value': value, 'source': origin}
        elif parts[0] == 'SNAP' and len(parts) == 2:
            if parts[1].strip() == 'no datasets available':
                continue
            if '@' not in parts[1] or parts[1] in snaps:
                raise ValueError('Invalid or duplicate snapshot name')
            snaps.add(parts[1])
        elif line.strip():
            raise ValueError('Unrecognized collection output: ' + line[:160])
    if not began or not ended:
        raise ValueError('Incomplete ZFS collection; native shell access may be required')
    if (not roots or not names) and not (hostname and not roots and not names):
        raise ValueError('No non-system pools with /local datasets were found')
    counts = Counter(name.split('@', 1)[0] for name in snaps)
    if set(counts) - names:
        raise ValueError('Dataset changed during collection; retry')
    records = []
    for name in sorted(names, key=lambda x: x.split('/')):
        if not any(name == r or name.startswith(r + '/') for r in roots):
            raise ValueError('Dataset outside selected roots')
        p = props.get(name, {})
        if any(prop not in p for prop in PROPERTIES):
            raise ValueError('Missing required properties: ' + name)
        records.append({'name': name, 'properties': p, 'snapshots': counts[name]})
    data = {'schema': 1, 'source': source, 'started_utc': started,
            'finished_utc': finished, 'roots': roots, 'datasets': records,
            'skipped_pools': skipped}
    if hostname:
        data['hostname'] = hostname
    validate_data(data)
    return data

def validate_data(data):
    if data.get('schema') != 1 or not isinstance(data.get('datasets'), list):
        raise ValueError('Expected a schema 1 controller capture')
    if not data['datasets'] and (data.get('roots') or not data.get('hostname')):
        raise ValueError('Empty controller capture must have a hostname and no dataset roots')
    names = set()
    for item in data['datasets']:
        name = item['name']
        if name in names or '@' in name or any(ord(c) < 32 for c in name):
            raise ValueError('Invalid or duplicate dataset: ' + name)
        names.add(name)
        p = item['properties']
        for prop in PROPERTIES:
            if prop not in p or 'value' not in p[prop]:
                raise ValueError('Missing property ' + prop + ' on ' + name)
        for prop in ('used', 'usedbysnapshots'):
            if not re.fullmatch(r'\d+', p[prop]['value']):
                raise ValueError('Expected exact bytes from zfs get -p: ' + name + '/' + prop)
        if p['type']['value'] not in ('filesystem', 'volume'):
            raise ValueError('Unexpected dataset type: ' + name)
        ratio = p['compressratio']['value']
        if not re.fullmatch(r'\d+(?:\.\d+)?x', ratio) or not math.isfinite(float(ratio[:-1])):
            raise ValueError('Invalid compression ratio: ' + name)
        if p['compression']['value'] in ('', '-'):
            raise ValueError('Missing compression setting: ' + name)
        if type(item['snapshots']) is not int or item['snapshots'] < 0:
            raise ValueError('Invalid snapshot count: ' + name)
    for root in data['roots']:
        if root not in names or not root.endswith('/local'):
            raise ValueError('Invalid /local root')
    for name in names:
        if not any(name == r or name.startswith(r + '/') for r in data['roots']):
            raise ValueError('Dataset outside /local roots')
        if name not in data['roots'] and name.rsplit('/', 1)[0] not in names:
            raise ValueError('Missing dataset parent: ' + name)

def stamp():
    return datetime.now(timezone.utc).isoformat(timespec='seconds')

def terminal_text(text):
    """Remove terminal decoration before matching prompts or parsing records."""
    text = re.sub(r'\x1b\][^\x07]*?(?:\x07|\x1b\\)', '', text)
    text = re.sub(r'\x1b\[[0-?]*[ -/]*[@-~]', '', text)
    text = re.sub(r'\x1b[()][0-2A-Z]', '', text)  # character-set selection
    text = re.sub(r'\x1b[=>78DMHEc]', '', text)  # keypad/cursor modes
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    while re.search(r'[^\n]\x08', text):
        text = re.sub(r'[^\n]\x08', '', text)
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1a\x1c-\x1f\x7f\ufeff\u200b\u200e\u200f]', '', text)
    return text.replace('\u00a0', ' ')

def appliance_dialogue(send, expect, program, token):
    """Advance only after each remote prompt/marker has actually arrived."""
    expect(r'(?m)^[ \t]*[A-Za-z0-9_.-]+:[^\r\n]*>[ \t]*$', 'appliance login prompt')
    send('confirm shell\n')
    expect(r'(?m)^[^\r\n|]*#[ \t]*$', 'raw shell prompt after confirm shell')
    # Disable remote terminal echo before sending the collector. The marker
    # must occur on its own line, so an echoed command cannot satisfy it.
    send("stty -echo; printf '\\n" + token + "_READY\\n'\n")
    expect(r'(?m)^' + token + r'_READY\r?$', 'raw shell readiness marker')
    delimiter = token + '_SCRIPT'
    send("/bin/sh <<'" + delimiter + "'\nprintf '\\n'\n" + program +
         '\n' + delimiter + "\nprintf '\\n" + token + ":%s\\n' \"$?\"\n")
    output, match = expect(r'(?m)^' + token + r':(\d+)\r?$', 'collector completion marker')
    return subprocess.CompletedProcess([], int(match.group(1)), output[:match.start()].replace('\r\n', '\n'))

def run_appliance(command, program, timeout, prompt_timeout=30, debug=False):
    # No remote command: reproduce the user's interactive login, including PTY.
    proc = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.PIPE, bufsize=0)
    chunks = queue.Queue()
    def reader():
        try:
            while True:
                chunk = os.read(proc.stdout.fileno(), 65536)
                if not chunk:
                    break
                chunks.put(chunk)
        finally:
            chunks.put(None)
    threading.Thread(target=reader, daemon=True).start()
    buffer = ''
    completed = False
    def send(text):
        proc.stdin.write(text.encode('utf-8'))
        proc.stdin.flush()
    def expect(pattern, stage):
        nonlocal buffer
        limit = timeout if stage == 'collector completion marker' else min(timeout, prompt_timeout)
        if stage == 'appliance login prompt':
            limit = min(timeout, max(60, prompt_timeout))
        deadline = time.monotonic() + limit
        print('Waiting for %s (up to %ss)...' % (stage, limit), file=sys.stderr, flush=True)
        def details():
            return ('\nRemote output:\n' + (terminal_text(buffer[-3000:]) or '(No remote stdout received.)') +
                    '\nRaw terminal tail: ' + repr(buffer[-500:]))
        while True:
            # Keep original chunks until a match, including split escape sequences.
            view = terminal_text(buffer)
            match = re.search(pattern, view)
            if match:
                seen = view
                buffer = ''
                print('Received ' + stage + '.', file=sys.stderr, flush=True)
                return seen, match
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise RuntimeError('Timed out waiting for ' + stage + '.' + details())
            try:
                chunk = chunks.get(timeout=min(remaining, 15))
            except queue.Empty:
                if time.monotonic() >= deadline:
                    raise RuntimeError('Timed out waiting for ' + stage + '.' + details())
                print('Still waiting for ' + stage + '...', file=sys.stderr, flush=True)
                continue
            if chunk is None:
                raise RuntimeError('SSH ended while waiting for ' + stage + '.' + details())
            decoded = chunk.decode('utf-8', errors='replace')
            if debug:
                print(terminal_text(decoded), end='', file=sys.stderr, flush=True)
            buffer += decoded
    try:
        result = appliance_dialogue(send, expect, program, 'ZFSREPORT_' + uuid.uuid4().hex)
        completed = True
        return result
    finally:
        if proc.poll() is None:
            try:
                if completed:
                    send('exit\nexit\n')
                else:
                    # Do not send further shell input after an unknown state.
                    proc.terminate()
                proc.wait(timeout=5)
            except (OSError, subprocess.TimeoutExpired):
                proc.terminate()
                try:
                    proc.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    proc.wait()
        proc.stdin.close()
        proc.stdout.close()

def collect(args):
    started = stamp()
    for pool in args.pool:
        if not re.fullmatch(r'[A-Za-z][A-Za-z0-9_.:-]*', pool):
            raise ValueError('Invalid pool name: ' + pool)
    program = 'SELECTED_POOLS=' + shlex.quote(' '.join(args.pool)) + '\n' + COLLECTOR
    if args.probe_native_shell:
        program = ("PATH=/usr/sbin:/usr/bin:/sbin:/bin:$PATH; export PATH\n"
                   "printf '__ZFS_NATIVE_SHELL_OK__\\n'\n"
                   "command -v zpool || exit 127\n"
                   "command -v zfs || exit 127\n")
    command = ['ssh', '-tt' if args.shell_mode == 'appliance' else '-T', '-p', str(args.port), '-o', 'ConnectTimeout=15',
               '-o', 'ServerAliveInterval=15', '-o', 'ServerAliveCountMax=3']
    if args.debug_ssh:
        command.append('-v')
    if args.batch:
        command += ['-o', 'BatchMode=yes']
    if args.identity:
        command += ['-i', args.identity]
    if not re.fullmatch(r'[A-Za-z0-9_.:-]+', args.host) or args.host.startswith('-'):
        raise ValueError('Invalid host')
    if not re.fullmatch(r'[A-Za-z0-9_.-]+', args.user) or args.user.startswith('-'):
        raise ValueError('Invalid SSH user')
    if args.shell_mode == 'appliance':
        command += [args.user + '@' + args.host]
        result = run_appliance(command, program, args.timeout,
                               getattr(args, 'prompt_timeout', 30), args.debug_ssh)
    else:
        command += [args.user + '@' + args.host, '/bin/sh -s']
        result = subprocess.run(command, input=program, universal_newlines=True, stdout=subprocess.PIPE,
                                timeout=args.timeout, check=False)
    if result.returncode:
        detail = result.stdout.strip()
        if detail:
            print('Remote standard output (last 30 lines):\n' +
                  '\n'.join(detail.splitlines()[-30:]), file=sys.stderr)
        raise RuntimeError('SSH/remote command failed (exit %s). No report was created. '
                           'This exit code alone does not identify the cause. '
                           'Run with --probe-native-shell --debug-ssh to distinguish SSH '
                           'authentication/connection errors from remote shell rejection.' % result.returncode)
    if args.probe_native_shell:
        print(result.stdout.rstrip())
        if '__ZFS_NATIVE_SHELL_OK__' not in result.stdout.splitlines():
            raise RuntimeError('The remote response did not confirm a native shell. '
                               'The account may enter the appliance management CLI.')
        print('Native shell probe passed. No storage data or configuration was changed.')
        return None
    try:
        data = parse_capture(result.stdout, args.user + '@' + args.host, started, stamp())
    except ValueError:
        print('Remote standard output (first 20 lines):\n' +
              '\n'.join(result.stdout.splitlines()[:20]), file=sys.stderr)
        raise
    missing = set(args.pool) - {r.split('/')[0] for r in data['roots']}
    if missing:
        raise ValueError('Requested pools missing or without /local: ' + ', '.join(sorted(missing)))
    return data

def cell(row, column, value, style, formula=None):
    c = ET.SubElement(row, q('c'), {'r': column + row.attrib['r'], 's': str(style)})
    if isinstance(value, str):
        c.set('t', 'inlineStr')
        t = ET.SubElement(ET.SubElement(c, q('is')), q('t'))
        t.text = value
    else:
        c.set('t', 'n')
        if formula:
            ET.SubElement(c, q('f')).text = formula
        ET.SubElement(c, q('v')).text = str(value)
    return c

def write_single_xlsx(data, output, expanded=False):
    validate_data(data)
    records = sorted(data['datasets'], key=lambda r: r['name'].split('/'))
    if len(records) > 1048568:
        raise ValueError('Too many rows for one Excel sheet')
    source = zipfile.ZipFile(io.BytesIO(base64.b64decode(TEMPLATE)))
    sheet = ET.fromstring(source.read('xl/worksheets/sheet1.xml'))
    styles = ET.fromstring(source.read('xl/styles.xml'))
    sd = sheet.find(q('sheetData'))
    prototypes = [[int(c.get('s', 0)) for c in sd.find(q('row') + "[@r='%s']" % row)]
                  for row in (9, 10, 11)]
    xfs = styles.find(q('cellXfs'))
    unit_styles = [int(c.get('s', 0)) for c in sd.find(q('row') + "[@r='14']")][:7]
    unit_formats = [xfs[i].get('numFmtId', '0') for i in unit_styles]
    cache = {}
    def styled(base, fmt=None, indent=None):
        key = (base, fmt, indent)
        if key not in cache:
            xf = copy.deepcopy(xfs[base])
            if fmt is not None:
                xf.set('numFmtId', fmt)
                xf.set('applyNumberFormat', '1')
            if indent is not None:
                a = xf.find(q('alignment'))
                if a is None:
                    a = ET.SubElement(xf, q('alignment'))
                a.set('indent', str(min(indent, 15)))
                a.set('horizontal', 'left')
                xf.set('applyAlignment', '1')
            cache[key] = len(xfs)
            xfs.append(xf)
        return cache[key]
    for row in list(sd):
        if int(row.get('r')) >= 9:
            sd.remove(row)
    for rownum, text in ((3, 'Source: %s; collected %s to %s' %
                         (data['source'], data['started_utc'], data['finished_utc'])),):
        row = sd.find(q('row') + "[@r='%s']" % rownum)
        row.clear()
        row.set('r', str(rownum))
        cell(row, 'A', text, 3)
    props = ET.Element(q('sheetPr'))
    ET.SubElement(props, q('outlinePr'), {'summaryBelow': '0', 'summaryRight': '0', 'showOutlineSymbols': '1'})
    sheet.insert(0, props)
    if not records:
        row = sd.find(q('row') + "[@r='7']")
        row.clear(); row.set('r', '7')
        cell(row, 'A', 'No local share datasets found on this controller.', 3)
    max_depth = 0
    for index, item in enumerate(records):
        rownum = index + 9
        name = item['name']
        root = next(r for r in data['roots'] if name == r or name.startswith(r + '/'))
        depth = name.count('/') - root.count('/')
        max_depth = max(max_depth, min(depth, 7))
        has_children = index + 1 < len(records) and records[index + 1]['name'].startswith(name + '/')
        attrs = {'r': str(rownum), 'ht': '23', 'customHeight': '1', 'outlineLevel': str(min(depth, 7))}
        if not expanded and depth >= 2:
            attrs['hidden'] = '1'
        if not expanded and has_children and depth >= 1:
            attrs['collapsed'] = '1'
        row = ET.SubElement(sd, q('row'), attrs)
        base = prototypes[min(depth, 2)]
        p = {k: v['value'] for k, v in item['properties'].items()}
        cell(row, 'A', name, styled(base[0], indent=depth))
        kind = 'Pool / local' if depth == 0 else ('Project' if depth == 1 else p['type'])
        cell(row, 'B', kind, base[1])
        for prop, human_col, bytes_col, human_style, bytes_style in (
            ('used', 'C', 'D', base[2], base[3]),
            ('usedbysnapshots', 'E', 'F', base[4], base[5])):
            amount = int(p[prop])
            unit = 0
            while unit < 6 and amount >= 1024 ** (unit + 1):
                unit += 1
            # Formula plus cached value means human sizes remain useful in Excel.
            cell(row, human_col, amount / 1024 ** unit, styled(human_style, unit_formats[unit]),
                 '%s%d/%d' % (bytes_col, rownum, 1024 ** unit))
            cell(row, bytes_col, str(amount) if amount >= 10**15 else amount, bytes_style)
        cell(row, 'G', item['snapshots'], base[6])
        cell(row, 'H', p['compression'], base[7])
        cell(row, 'I', float(p['compressratio'][:-1]), base[8])
    xfs.set('count', str(len(xfs)))
    sheet.find(q('sheetFormatPr')).set('outlineLevelRow', str(max_depth))
    # No frozen or split panes on any controller tab.
    for view in sheet.findall('.//' + q('sheetView')):
        for child in list(view):
            if child.tag in (q('pane'), q('selection')):
                view.remove(child)
    dimension = ET.Element(q('dimension'), {'ref': 'A1:I%d' % (len(records) + 8)})
    sheet.insert(1, dimension)
    replacements = {'xl/worksheets/sheet1.xml': ET.tostring(sheet, encoding='utf-8', xml_declaration=True),
                    'xl/styles.xml': ET.tostring(styles, encoding='utf-8', xml_declaration=True)}
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as dest:
        for info in source.infolist():
            dest.writestr(info.filename, replacements.get(info.filename, source.read(info.filename)))

def controller_captures(data):
    captures = data.get('controllers', []) if data.get('schema') == 2 else [data]
    if not 1 <= len(captures) <= 2:
        raise ValueError('Expected one or two controller captures')
    for capture in captures:
        validate_data(capture)
    return captures

def write_xlsx(data, output, expanded=False):
    captures = controller_captures(data)
    packages = []
    for capture in captures:
        buffer = io.BytesIO()
        write_single_xlsx(capture, buffer, expanded)
        with zipfile.ZipFile(buffer) as z:
            packages.append({name: z.read(name) for name in z.namelist()})
    package = packages[0]
    workbook = ET.fromstring(package['xl/workbook.xml'])
    sheets = workbook.find(q('sheets'))
    relations = ET.fromstring(package['xl/_rels/workbook.xml.rels'])
    types = ET.fromstring(package['[Content_Types].xml'])
    styles = ET.fromstring(package['xl/styles.xml'])
    xfs = styles.find(q('cellXfs'))
    rns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    cns = 'http://schemas.openxmlformats.org/package/2006/content-types'
    used = set()
    for index, capture in enumerate(captures):
        hostname = capture.get('hostname') or capture.get('source', 'Controller').split('@')[-1]
        name = re.sub(r'[\x00-\x1f\[\]:*?/\\]', '_', hostname).strip("'") or 'Controller'
        base = name[:31]; name = base; suffix = 2
        while name.lower() in used:
            tail = ' (%d)' % suffix
            name = base[:31-len(tail)] + tail; suffix += 1
        used.add(name.lower())
        if index == 0:
            sheets[0].set('name', name)
            continue
        other = ET.fromstring(packages[index]['xl/styles.xml']).find(q('cellXfs'))
        offset = len(xfs)
        for xf in other:
            xfs.append(copy.deepcopy(xf))
        sheet = ET.fromstring(packages[index]['xl/worksheets/sheet1.xml'])
        for c in sheet.findall('.//' + q('c')):
            c.set('s', str(offset + int(c.get('s', '0'))))
        path = 'xl/worksheets/sheet%d.xml' % (index + 1)
        package[path] = ET.tostring(sheet, encoding='utf-8', xml_declaration=True)
        rid = 'controllerSheet%d' % (index + 1)
        ET.SubElement(sheets, q('sheet'), {'name': name, 'sheetId': str(index + 1), '{'+rns+'}id': rid})
        ET.SubElement(relations, '{'+pns+'}Relationship', {'Id': rid, 'Type': rns+'/worksheet', 'Target': '/'+path})
        ET.SubElement(types, '{'+cns+'}Override', {'PartName': '/'+path,
                       'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
    xfs.set('count', str(len(xfs)))
    for path, element in [('xl/workbook.xml',workbook), ('xl/_rels/workbook.xml.rels',relations),
                          ('[Content_Types].xml',types), ('xl/styles.xml',styles)]:
        if path == 'xl/_rels/workbook.xml.rels':
            ET.register_namespace('', pns)
        elif path == '[Content_Types].xml':
            ET.register_namespace('', cns)
        package[path] = ET.tostring(element, encoding='utf-8', xml_declaration=True)
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
        for path, content in package.items():
            z.writestr(path, content)

def collect_controllers(args):
    captures = []
    for host in [args.host] + ([args.host2] if args.host2 else []):
        controller_args = copy.copy(args)
        controller_args.host = host
        print('Connecting to controller ' + host + '...', file=sys.stderr, flush=True)
        captures.append(collect(controller_args))
    if args.probe_native_shell:
        return None
    return captures[0] if len(captures) == 1 else {'schema': 2, 'controllers': captures}

def ask_connection(args):
    if args.from_json:
        return
    missing = [name for name in ('host', 'user') if not getattr(args, name)]
    if args.batch and missing:
        raise ValueError('--batch requires ' + ' and '.join('--' + name for name in missing))
    for name, label in (('host', 'Controller 1 IP address or hostname'),):
        while not getattr(args, name):
            try:
                value = input(label + ': ').strip()
            except EOFError:
                raise ValueError('No input available. Supply --host and --user on the command line.')
            if value:
                setattr(args, name, value)
            else:
                print('Please enter a value.', file=sys.stderr)
    if getattr(args, 'host2', None) is None:
        if args.batch:
            args.host2 = ''
        else:
            try:
                args.host2 = input('Controller 2 IP address or hostname (Enter for standalone/simulator): ').strip()
            except EOFError:
                args.host2 = ''
    if args.host2 and args.host2.lower() == args.host.lower():
        raise ValueError('Controller addresses must be different')
    while not args.user:
        try:
            args.user = input('ZFS SSH username (used for both controllers): ').strip()
        except EOFError:
            raise ValueError('No username provided; supply --user')

def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--version', action='version', version='zfs_shares_report 1.5 (two controllers, no frozen panes)')
    parser.add_argument('--host', help='ZFS IP address or hostname; asks when omitted')
    parser.add_argument('--host2', help='Optional second controller; blank means standalone/simulator')
    parser.add_argument('--single', action='store_true', help='Skip the second-controller prompt')
    parser.add_argument('--user', help='ZFS SSH username; asks when omitted')
    parser.add_argument('--port', type=int, default=22)
    parser.add_argument('--identity', help='SSH private key path (optional)')
    parser.add_argument('--batch', action='store_true', help='Never prompt for SSH credentials')
    parser.add_argument('--shell-mode', choices=('appliance', 'native'), default='appliance',
                        help="appliance: enter with 'confirm shell' (default); native: ordinary SSH shell")
    parser.add_argument('--debug-ssh', action='store_true', help='Show SSH connection/authentication diagnostics')
    parser.add_argument('--probe-native-shell', action='store_true',
                        help='Check native shell and zfs/zpool command availability; create no report')
    parser.add_argument('--pool', action='append', default=[], help='Restrict pool; repeat for multiple pools')
    parser.add_argument('--timeout', type=int, default=600, help='Collection timeout in seconds')
    parser.add_argument('--prompt-timeout', type=int, default=30,
                        help='Timeout for each shell transition; login allows at least 60 seconds')
    parser.add_argument('--output', type=Path, help='XLSX filename; default is timestamped')
    parser.add_argument('--expanded', action='store_true', help='Open with all outline rows visible')
    parser.add_argument('--from-json', type=Path, help='Rebuild an earlier JSON capture without SSH')
    args = parser.parse_args()
    if args.single:
        if args.host2:
            parser.error('--single cannot be combined with a second controller')
        args.host2 = ''
    if not 1 <= args.port <= 65535 or args.timeout < 1 or args.prompt_timeout < 1:
        parser.error('Port or timeout out of range')
    output = args.output or Path('zfs_shares_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '.xlsx')
    if output.suffix.lower() != '.xlsx':
        parser.error('--output must end in .xlsx')
    raw = output.with_suffix('.json')
    if output.exists() or raw.exists():
        parser.error('Output XLSX or JSON already exists; choose a new output filename')
    try:
        ask_connection(args)
        if args.probe_native_shell:
            if args.from_json:
                parser.error('--probe-native-shell cannot be combined with --from-json')
            collect_controllers(args)
            return 0
        data = json.loads(args.from_json.read_text(encoding='utf-8')) if args.from_json else collect_controllers(args)
        captures = controller_captures(data)
        output.parent.mkdir(parents=True, exist_ok=True)
        # Render completely before creating either final file.
        buffer = io.BytesIO()
        write_xlsx(data, buffer, args.expanded)
        with output.open('xb') as f:
            f.write(buffer.getvalue())
        with raw.open('x', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print('Created %s (%d controller tabs, %d datasets)' %
              (output.resolve(), len(captures), sum(len(c['datasets']) for c in captures)))
        print('Raw capture: ' + str(raw.resolve()))
        return 0
    except KeyboardInterrupt:
        print('\nCancelled.', file=sys.stderr)
        return 130
    except (OSError, ValueError, KeyError, RuntimeError, subprocess.TimeoutExpired) as exc:
        print('ERROR: ' + str(exc), file=sys.stderr)
        return 1

if __name__ == '__main__':
    sys.exit(main())
