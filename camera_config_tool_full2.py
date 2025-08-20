
#!/usr/bin/env python3
"""
camera_config_tool_full.py
(WIDE Excel + Targets sheet + JSON defaults + per-section overrides)
"""
import argparse, sys, json
from typing import Dict, Any, List, Tuple, Optional
from collections import defaultdict
import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

requests.packages.urllib3.disable_warnings()

def build_url(scheme: str, host: str, path: str) -> str:
    path = path.lstrip('/')
    return f"{scheme}://{host}/{path}"

def auth_params(user: str, password: str, rest: List[Tuple[str, Any]]) -> List[Tuple[str, Any]]:
    return [('userName', user), ('password', password)] + rest

def http_get(url: str, params: List[Tuple[str, Any]], verify_ssl: bool) -> str:
    r = requests.get(url, params=params, timeout=12, verify=verify_ssl)
    r.raise_for_status()
    return r.text

def parse_kv(body: str) -> Dict[str, str]:
    out = {}
    for line in body.splitlines():
        line = line.strip()
        if not line or '=' not in line or line.lower().startswith('http'):
            continue
        k, v = line.split('=', 1)
        out[k.strip()] = v.strip()
    return out

BUILTIN_DEFAULTS = {
    "get": {"enabled": True, "action": "get", "requires_type": True},
    "set": {"enabled": True, "action": "set", "requires_type": True},
}

def resolve_mode(defaults: Dict[str,Any], section_modes: Optional[Dict[str,Any]], mode: str) -> Dict[str,Any]:
    out = dict(BUILTIN_DEFAULTS.get(mode, {}))
    if defaults and mode in defaults:
        for k,v in defaults[mode].items():
            out[k] = v
    if section_modes and mode in section_modes:
        for k,v in section_modes[mode].items():
            out[k] = v
    out.setdefault("enabled", True)
    out.setdefault("action", "get" if mode=="get" else "set")
    out.setdefault("requires_type", True)
    return out

def read_targets(path: str) -> Dict[str, List[Dict[str,str]]]:
    wb = load_workbook(path, data_only=True)
    if 'Targets' not in wb.sheetnames:
        return {}
    ws = wb['Targets']
    headers = []
    c = 1
    while c <= ws.max_column:
        h = ws.cell(row=1, column=c).value
        if h is None or str(h).strip() == '':
            break
        headers.append(str(h).strip())
        c += 1
    if not headers or headers[0].lower() != 'section':
        raise ValueError("Targets sheet: first column header must be 'Section'")
    result: Dict[str, List[Dict[str,str]]] = defaultdict(list)
    r = 2
    while r <= ws.max_row:
        sec = ws.cell(row=1*r//1, column=1).value if False else ws.cell(row=r, column=1).value
        if sec in (None, ''):
            r += 1
            continue
        sec = str(sec).strip()
        entry: Dict[str,str] = {}
        any_value = False
        for idx, name in enumerate(headers[1:], start=2):
            val = ws.cell(row=r, column=idx).value
            if val in (None, ''):
                continue
            entry[name] = str(val).strip()
            any_value = True
        result[sec].append(entry if any_value else {})
        r += 1
    return result

def write_wide(path: str, sections: List[str], data_by_section: Dict[str, List[Tuple[str,str]]]):
    wb = Workbook()
    ws = wb.active
    ws.title = "Config"
    for idx, sec in enumerate(sections):
        c = idx*3 + 1
        cell = ws.cell(row=1, column=c, value=sec)
        cell.font = Font(bold=True, underline="single")
    max_rows = max([len(data_by_section.get(sec, [])) for sec in sections] + [0])
    for r_off in range(max_rows):
        r = r_off + 2
        for idx, sec in enumerate(sections):
            c = idx*3 + 1
            pairs = data_by_section.get(sec, [])
            if r_off < len(pairs):
                k, v = pairs[r_off]
                ws.cell(row=r, column=c, value=k)
                ws.cell(row=r, column=c+1, value=v)
    wb.save(path)

def read_wide(path: str) -> Dict[str, List[Tuple[str,str]]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    sections = []
    col = 1
    empty_headers = 0
    while col <= ws.max_column:
        header = ws.cell(row=1, column=col).value
        if header and str(header).strip():
            sections.append((str(header).strip(), col))
            empty_headers = 0
        else:
            empty_headers += 1
            if empty_headers >= 5:
                break
        col += 3
    data = {}
    for sec, c in sections:
        pairs: List[Tuple[str,str]] = []
        blanks = 0
        row = 2
        while row <= ws.max_row:
            k = ws.cell(row=row, column=c).value
            v = ws.cell(row=row, column=c+1).value
            if (k in (None,'') and v in (None,'')):
                blanks += 1
                if blanks >= 3:
                    break
            else:
                blanks = 0
                ks = '' if k is None else str(k)
                vs = '' if v is None else str(v)
                if ks or vs:
                    pairs.append((ks.strip(), vs.strip()))
            row += 1
        data[sec] = pairs
    return data

def export_configs(ip, user, password, scheme, verify_ssl, caps, selectors) -> Dict[str, List[Tuple[str,str]]]:
    data_by_section: Dict[str, List[Tuple[str,str]]] = {}
    defaults = caps.get("defaults", {})
    secmap = caps.get("sections", caps)
    preferred = ['deviceInfo','deviceName','localNetwork','devicePort','AVStream','OSD','motionAlarm','diskAlarm','perimeterParam','NTP','DDNS','SMTP','User','streamAbility','setNorthPosition']
    types = preferred + [t for t in sorted(secmap.keys()) if t not in preferred]
    for t in types:
        entry = secmap.get(t, {})
        endpoint = entry.get('endpoint', 'param.cgi')
        modes = entry.get('modes', {})
        mget = resolve_mode(defaults, modes, "get")
        pairs: List[Tuple[str,str]] = []
        sels = selectors.get(t) or [None]
        for sel in sels:
            if not mget.get("enabled", True):
                continue
            params = [('action', mget.get('action', 'get'))]
            if mget.get("requires_type", True):
                params.append(('type', t))
            if sel:
                for k,v in sel.items():
                    params.append((k,v))
            try:
                url = build_url(scheme, ip, f'/cgi-bin/{endpoint}')
                body = http_get(url, auth_params(user, password, params), verify_ssl)
                kv = parse_kv(body)
                if sel:
                    for k,v in sel.items():
                        pairs.append((k, str(v)))
                for k,v in kv.items():
                    pairs.append((k, v))
            except Exception as e:
                pairs.append(('#error', f"{t} {mget.get('action','get')} failed: {e}"))
        if pairs:
            data_by_section[t] = pairs
    return data_by_section

def apply_configs(data_by_section: Dict[str, List[Tuple[str,str]]], ip, user, password, scheme, verify_ssl, caps, allow_unknown=False) -> Dict[str, Any]:
    report = {'applied': [], 'skipped': []}
    defaults = caps.get("defaults", {})
    secmap = caps.get("sections", caps)
    for section, pairs in data_by_section.items():
        entry = secmap.get(section, {})
        endpoint = entry.get('endpoint', 'param.cgi')
        allowed = set(entry.get('set_params', []))
        modes = entry.get('modes', {})
        mset = resolve_mode(defaults, modes, "set")
        if not mset.get("enabled", True):
            report['skipped'].append({'section': section, 'reason': 'set disabled by capabilities'})
            continue
        payload = [('action', mset.get('action', 'set'))]
        if mset.get("requires_type", True):
            payload.append(('type', section))
        unknown = []
        for k,v in pairs:
            ku = (k or '').upper()
            if ku in ('', 'SECTION', '#ERROR'):
                continue
            if (k in allowed) or allow_unknown:
                payload.append((k, v))
            else:
                unknown.append(k)
        if unknown and not allow_unknown:
            report['skipped'].append({'section': section, 'reason': 'non-settable keys', 'keys': unknown})
        if len(payload) > (1 if not mset.get("requires_type", True) else 2):
            try:
                url = build_url(scheme, ip, f'/cgi-bin/{endpoint}')
                resp = http_get(url, auth_params(user, password, payload), verify_ssl)
                report['applied'].append({'section': section, 'keys':[k for (k,_) in payload if k not in ('action','type')], 'response': resp[:200]})
            except Exception as e:
                report['skipped'].append({'section': section, 'reason': f"set failed: {e}", 'keys':[k for (k,_) in payload]})
    return report

def struct_from_wide(data_by_section: Dict[str, List[Tuple[str,str]]]) -> Dict[str, Dict[str,str]]:
    out: Dict[str, Dict[str,str]] = {}
    for sec, pairs in data_by_section.items():
        out.setdefault(sec, {})
        for k,v in pairs:
            if k and not k.startswith('#'):
                out[sec][k] = v
    return out

def diff_struct(actual: Dict[str,Any], template: Dict[str,Any]) -> List[str]:
    diffs = []
    secs = set(actual.keys()) | set(template.keys())
    for sec in sorted(secs):
        a = actual.get(sec, {})
        b = template.get(sec, {})
        keys = set(a.keys()) | set(b.keys())
        for k in sorted(keys):
            av = str(a.get(k,''))
            bv = str(b.get(k,''))
            if av != bv:
                diffs.append(f"{sec}.{k}: actual='{av}' expected='{bv}'")
    return diffs

def parse_selectors_cli(sel_list: List[str]) -> Dict[str, List[Dict[str, str]]]:
    out = defaultdict(list)
    for s in sel_list or []:
        if ':' not in s:
            continue
        section, kvs = s.split(':',1)
        sel = {}
        for kv in kvs.split(','):
            if '=' in kv:
                k,v = kv.split('=',1)
                sel[k.strip()] = v.strip()
        if sel:
            out[section].append(sel)
    return out

def main():
    ap = argparse.ArgumentParser(description="Camera CGI tool (WIDE + Targets + JSON defaults/overrides).")
    sub = ap.add_subparsers(dest='cmd', required=True)
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument('--ip', required=True)
    common.add_argument('-u','--user', required=True)
    common.add_argument('-p','--password', required=True)
    common.add_argument('--scheme', default='http', choices=['http','https'])
    common.add_argument('--insecure', action='store_true')
    common.add_argument('--caps', required=True, help='Path to capabilities JSON')

    pe = sub.add_parser('export', parents=[common])
    pe.add_argument('-o','--output', default='export.xlsx')
    pe.add_argument('--targets', help='Excel file with a "Targets" sheet')
    pe.add_argument('--selectors', nargs='*', help='(optional) selectors if no Targets provided')

    pa = sub.add_parser('apply', parents=[common])
    pa.add_argument('-i','--input', required=True)
    pa.add_argument('--allow-unknown-keys', action='store_true')

    pd = sub.add_parser('diff', parents=[common])
    pd.add_argument('-t','--template', required=True)
    pd.add_argument('--targets', help='Excel file with a "Targets" sheet')
    pd.add_argument('--selectors', nargs='*', help='(optional) selectors if no Targets provided')

    args = ap.parse_args()
    verify_ssl = not args.insecure

    with open(args.caps, 'r', encoding='utf-8') as f:
        caps = json.load(f)

    selectors = {}
    if getattr(args, 'targets', None):
        try:
            selectors = read_targets(args.targets)
        except Exception as e:
            print(f"[WARN] Failed to read Targets from {args.targets}: {e}")
            selectors = {}
    if not selectors and hasattr(args, 'selectors'):
        selectors = parse_selectors_cli(args.selectors or [])

    if args.cmd == 'export':
        data = export_configs(args.ip, args.user, args.password, args.scheme, verify_ssl, caps, selectors)
        sections = list(data.keys())
        write_wide(args.output, sections, data)
        print(f"Exported to {args.output}")
        return 0

    if args.cmd == 'apply':
        data = read_wide(args.input)
        rpt = apply_configs(data, args.ip, args.user, args.password, args.scheme, verify_ssl, caps, allow_unknown=args.allow_unknown_keys)
        print("Apply report:")
        for a in rpt['applied']:
            print("  APPLIED:", a)
        for s in rpt['skipped']:
            print("  SKIPPED:", s)
        return 0

    if args.cmd == 'diff':
        live = export_configs(args.ip, args.user, args.password, args.scheme, verify_ssl, caps, selectors)
        live_struct = struct_from_wide(live)
        tmpl_struct = struct_from_wide(read_wide(args.template))
        diffs = diff_struct(live_struct, tmpl_struct)
        if not diffs:
            print("No differences.")
            return 0
        print("Differences:")
        for d in diffs:
            print("  -", d)
        return 1

if __name__ == "__main__":
    sys.exit(main())
