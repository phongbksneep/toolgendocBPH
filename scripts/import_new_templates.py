#!/usr/bin/env python3
import json
import os
import re
import unicodedata
import zipfile
import html
from collections import defaultdict, Counter


SRC_DIR = '/var/www/Hồ sơ chạy thử VBA'
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
TEMPLATES_DIR = os.path.join(ROOT, 'templates')
CATALOG_PATH = os.path.join(ROOT, 'catalog.json')
DEMO_PATH = os.path.join(ROOT, 'project-data.demo.friendly.json')
SUMMARY_PATH = os.path.join(ROOT, 'TONG-HOP-BIEN-LAP.md')
BUNDLE_PATH = os.path.join(ROOT, 'assets', 'templates_bundle.zip')

RUN_RE = re.compile(r'<w:r\b[\s\S]*?</w:r>')
P_RE = re.compile(r'<w:p\b[\s\S]*?</w:p>')
WT_RE = re.compile(r'(<w:t(?:\s[^>]*)?>)([\s\S]*?)(</w:t>)')


def nfc(s: str) -> str:
    return unicodedata.normalize('NFC', s)


def strip_accents(s: str) -> str:
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return unicodedata.normalize('NFC', s)


def norm_name(s: str) -> str:
    s = nfc(s).lower()
    s = strip_accents(s)
    s = re.sub(r'[^a-z0-9]+', '', s)
    return s


def key_norm_text(s: str) -> str:
    s = html.unescape(nfc(s))
    s = s.replace('\xa0', ' ')
    s = s.strip()
    s = strip_accents(s).lower()
    s = re.sub(r'[_\-–—.,;:!?/\\()\[\]{}"“”‘’`]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def slugify(s: str, max_len=64) -> str:
    s = html.unescape(nfc(s))
    s = strip_accents(s).lower()
    s = re.sub(r'[^a-z0-9]+', '_', s)
    s = re.sub(r'_+', '_', s).strip('_')
    if not s:
        s = 'field'
    return s[:max_len].rstrip('_')


def extract_text(run_xml: str) -> str:
    chunks = [html.unescape(m.group(2)) for m in WT_RE.finditer(run_xml)]
    return ''.join(chunks)


def replace_first_text(run_xml: str, new_text: str) -> str:
    esc = html.escape(new_text)
    done = False

    def repl(m):
        nonlocal done
        if not done:
            done = True
            return f'{m.group(1)}{esc}{m.group(3)}'
        return f'{m.group(1)}{m.group(2)}{m.group(3)}'

    out = WT_RE.sub(repl, run_xml)
    if not done:
        out = out.replace('</w:r>', f'<w:t xml:space="preserve">{esc}</w:t></w:r>', 1)
    return out


def clear_all_text(run_xml: str) -> str:
    return WT_RE.sub(lambda m: f'{m.group(1)}{m.group(3)}', run_xml)


def should_skip(owner: str, text: str) -> bool:
    t = html.unescape(nfc(text)).replace('\xa0', ' ')
    t = re.sub(r'\s+', ' ', t).strip()
    if not t:
        return True
    # Skip pure punctuation
    if re.fullmatch(r'[-–—.,;:!?/\\()\[\]{}"“”‘’`]+', t):
        return True
    # Skip tiny numeric/punct fragments for GLOBAL only
    if owner == 'GLOBAL':
        compact = re.sub(r'\s+', '', t)
        if re.fullmatch(r'[0-9]+', compact) and len(compact) <= 2:
            return True
    return False


def split_order_prefix(name: str):
    """Return (order_no:int|None, base_name:str). E.g. '15. PHỤ LỤC.docx' -> (15, 'PHỤ LỤC.docx')"""
    m = re.match(r'^\s*(\d+)\s*[\._-]?\s*(.+)$', nfc(name))
    if not m:
        return None, nfc(name)
    return int(m.group(1)), m.group(2).strip()


def build_source_list():
    files = [f for f in os.listdir(SRC_DIR) if f.lower().endswith('.docx') and not should_ignore_source(f)]

    def _k(f):
      no, base = split_order_prefix(f)
      return (0 if no is not None else 1, no if no is not None else 10**9, norm_name(base), norm_name(f))

    files.sort(key=_k)
    return files

def convert_doc_to_docx_files():
    # Best-effort conversion for legacy .doc files.
    # - Uses textutil on macOS if available.
    # - Otherwise skips and logs.
    doc_files = [f for f in os.listdir(SRC_DIR) if f.lower().endswith('.doc')]
    if not doc_files:
        return []

    converted = []
    textutil = '/usr/bin/textutil'
    has_textutil = os.path.exists(textutil)

    for fn in doc_files:
        src = os.path.join(SRC_DIR, fn)
        dst = os.path.join(SRC_DIR, os.path.splitext(fn)[0] + '.docx')
        if os.path.exists(dst):
            converted.append(os.path.basename(dst))
            continue
        if has_textutil:
            rc = os.system(f'"{textutil}" -convert docx "{src}" -output "{dst}" >/dev/null 2>&1')
            if rc == 0 and os.path.exists(dst):
                converted.append(os.path.basename(dst))
    return converted


def classify_owner(src_name: str):
    n = norm_name(src_name)
    if 'nghiemthuvatlieu' in n:
        return 'LIST_VAT_LIEU'
    if 'bienbannghiemthucongviec' in n:
        return 'LIST_NTCV'
    if 'yeucaunghiemthucongvieclan' in n:
        return 'LIST_YC_NTCV'
    return 'GLOBAL'


def should_ignore_source(src_name: str) -> bool:
    # Ignore temp/system files only
    if src_name.startswith('~$') or src_name.startswith('.'):
        return True
    return False


def choose_group_base(files_by_owner):
    base = {}
    for owner in ('LIST_VAT_LIEU', 'LIST_NTCV', 'LIST_YC_NTCV'):
        cands = files_by_owner.get(owner, [])
        if not cands:
            continue
        # Use first file in numbered order as base template for that group
        base[owner] = cands[0]
    return base


def map_target_templates():
    tpls = [f for f in os.listdir(TEMPLATES_DIR) if f.lower().endswith('.docx')]
    by_norm = {norm_name(f): f for f in tpls}

    def find_like(keyword):
        k = norm_name(keyword)
        for n, real in by_norm.items():
            if k in n:
                return real
        raise RuntimeError(f'Cannot map template: {keyword}')

    return {
        'LIST_NTCV': find_like('BB Nghiệm thu công việc.docx'),
        'LIST_VAT_LIEU': find_like('Nghiệm thu vật liệu.docx'),
        'LIST_YC_NTCV': find_like('Yêu cầu nghiệm thu công việc.docx'),
    }


def source_to_template_name(src_name: str) -> str:
    # Source files are numbered by desired order.
    # Template files in app do NOT keep numeric prefixes, so normalize by removing prefix.
    no, base = split_order_prefix(src_name)
    base_n = norm_name(base)

    # Keep special known aliases if any mismatch remains
    mapping = {
        norm_name('BBBG đưa vào sử dụng.docx'): 'BBBG đưa vào sử dụng.docx',
    }
    if base_n in mapping:
        return mapping[base_n]
    return nfc(base) if no is not None else nfc(src_name)


def pretty_output_name(src_name: str) -> str:
    # Keep customer's numbered order in output filenames
    return nfc(src_name)


def process_xml(xml: str, owner: str, target_template: str, key_db, catalog, sample_capture, *, old_catalog=None, order_state=None, allow_new_keys=True):
    out_parts = []
    cur = 0

    for pm in P_RE.finditer(xml):
        out_parts.append(xml[cur:pm.start()])
        para = pm.group(0)
        runs = list(RUN_RE.finditer(para))
        if not runs:
            out_parts.append(para)
            cur = pm.end()
            continue

        rebuilt = []
        pcur = 0
        i = 0
        while i < len(runs):
            r = runs[i]
            rxml = r.group(0)
            is_hl = '<w:highlight' in rxml
            if not is_hl:
                rebuilt.append(para[pcur:r.end()])
                pcur = r.end()
                i += 1
                continue

            j = i
            grp = []
            while j < len(runs) and ('<w:highlight' in runs[j].group(0)):
                grp.append(runs[j].group(0))
                j += 1

            combined = ''.join(extract_text(x) for x in grp)
            combined = combined.replace('\xa0', ' ')
            combined = re.sub(r'\s+', ' ', combined).strip()

            new_runs = grp
            if not should_skip(owner, combined):
                key_norm = key_norm_text(combined)
                if key_norm:
                    owner_map = key_db[owner]
                    if key_norm in owner_map:
                        key = owner_map[key_norm]
                    else:
                        key = None

                        # Reuse same normalized text from previous catalog for stable keys
                        if old_catalog:
                            for ok, ov in old_catalog.items():
                                if ov.get('owner') != owner:
                                    continue
                                olabel = ov.get('label') or ov.get('sample') or ''
                                if key_norm_text(olabel) == key_norm:
                                    key = ok
                                    break

                        # Optional: do not create new keys when caller only wants extracting values
                        if key is None and not allow_new_keys:
                            key = None
                        elif key is None:
                            base = slugify(key_norm)
                            key = f'{owner}__{base}'
                            idx = 2
                            while key in catalog:
                                key = f'{owner}__{base}_{idx}'
                                idx += 1

                        if key is not None:
                            owner_map[key_norm] = key
                            if key not in catalog:
                                # Preserve old label/sample if available to avoid raw-value labels
                                old = (old_catalog or {}).get(key, {})
                                stable_label = old.get('label') or combined
                                stable_sample = old.get('sample') or combined
                                catalog[key] = {
                                    'owner': owner,
                                    'label': stable_label,
                                    'files': [],
                                    'count': 0,
                                    'sample': stable_sample,
                                }

                    if key is not None:
                        if target_template not in catalog[key]['files']:
                            catalog[key]['files'].append(target_template)
                        catalog[key]['count'] += 1
                        sample_capture[key] = combined

                        # Keep first-seen order index (for UI ordering by source file)
                        if order_state is not None:
                            order_state.setdefault('counter', 0)
                            if key not in order_state['index']:
                                order_state['index'][key] = order_state['counter']
                                order_state['counter'] += 1

                    if key is not None:
                        ph = '{{' + key + '}}'
                        new_runs = [replace_first_text(grp[0], ph)] + [clear_all_text(x) for x in grp[1:]]

            rebuilt.append(para[pcur:r.start()])
            rebuilt.extend(new_runs)
            pcur = runs[j - 1].end()
            i = j

        rebuilt.append(para[pcur:])
        out_parts.append(''.join(rebuilt))
        cur = pm.end()

    out_parts.append(xml[cur:])
    return ''.join(out_parts)


def main():
    converted_docx = convert_doc_to_docx_files()
    src_files = build_source_list()
    files_by_owner = defaultdict(list)
    for f in src_files:
        files_by_owner[classify_owner(f)].append(f)

    target_group_templates = map_target_templates()
    group_base = choose_group_base(files_by_owner)

    # Map GLOBAL source docs to template files by normalized filename
    existing_templates = [f for f in os.listdir(TEMPLATES_DIR) if f.lower().endswith('.docx')]
    tpl_norm_map = {norm_name(f): f for f in existing_templates}

    # Load previous catalog to keep stable labels/keys (avoid raw value labels)
    old_catalog = {}
    if os.path.exists(CATALOG_PATH):
        try:
            with open(CATALOG_PATH, 'r', encoding='utf-8') as f:
                old_catalog = json.load(f)
        except Exception:
            old_catalog = {}

    key_db = defaultdict(dict)   # owner -> norm_text -> key
    catalog = {}

    # Track field order by first appearance following customer file order
    order_state = {'index': {}, 'counter': 0}

    # Gather list-jobs values from each source file
    list_job_values = defaultdict(dict)  # src_name -> {key: value}

    # Update shared group templates from selected base files
    for owner, base_src in group_base.items():
        target_tpl = target_group_templates[owner]
        src_path = os.path.join(SRC_DIR, base_src)
        dst_path = os.path.join(TEMPLATES_DIR, target_tpl)

        with zipfile.ZipFile(src_path, 'r') as zin:
            items = {n: zin.read(n) for n in zin.namelist()}

        sample_capture = {}
        xml = items['word/document.xml'].decode('utf-8', 'ignore')
        new_xml = process_xml(
            xml, owner, target_tpl, key_db, catalog, sample_capture,
            old_catalog=old_catalog, order_state=order_state, allow_new_keys=True
        )
        items['word/document.xml'] = new_xml.encode('utf-8')

        with zipfile.ZipFile(dst_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for n, b in items.items():
                zout.writestr(n, b)

    # Process each source group file to collect list job values using existing keys (without writing template)
    for owner in ('LIST_VAT_LIEU', 'LIST_NTCV', 'LIST_YC_NTCV'):
        target_tpl = target_group_templates[owner]
        for src_name in files_by_owner.get(owner, []):
            src_path = os.path.join(SRC_DIR, src_name)
            with zipfile.ZipFile(src_path, 'r') as zin:
                xml = zin.read('word/document.xml').decode('utf-8', 'ignore')

            sample_capture = {}
            _ = process_xml(
                xml, owner, target_tpl, key_db, catalog, sample_capture,
                old_catalog=old_catalog, order_state=order_state, allow_new_keys=False
            )
            list_job_values[src_name] = sample_capture

    # Update GLOBAL templates directly from matching source docs
    for src_name in files_by_owner.get('GLOBAL', []):
        mapped = source_to_template_name(src_name)
        s_norm = norm_name(mapped)
        if s_norm not in tpl_norm_map:
            # try fuzzy contains
            cand = None
            for tn, real in tpl_norm_map.items():
                if s_norm in tn or tn in s_norm:
                    cand = real
                    break
            if not cand:
                # New global template requested by customer (e.g. PHỤ LỤC HĐ PS)
                target_tpl = mapped if mapped.lower().endswith('.docx') else (mapped + '.docx')
                tpl_norm_map[s_norm] = target_tpl
                if target_tpl not in existing_templates:
                    existing_templates.append(target_tpl)
            else:
                target_tpl = cand
        else:
            target_tpl = tpl_norm_map[s_norm]

        src_path = os.path.join(SRC_DIR, src_name)
        dst_path = os.path.join(TEMPLATES_DIR, target_tpl)

        with zipfile.ZipFile(src_path, 'r') as zin:
            items = {n: zin.read(n) for n in zin.namelist()}

        sample_capture = {}
        xml = items['word/document.xml'].decode('utf-8', 'ignore')
        new_xml = process_xml(
            xml, 'GLOBAL', target_tpl, key_db, catalog, sample_capture,
            old_catalog=old_catalog, order_state=order_state, allow_new_keys=True
        )
        items['word/document.xml'] = new_xml.encode('utf-8')

        with zipfile.ZipFile(dst_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for n, b in items.items():
                zout.writestr(n, b)

    # Sort catalog by owner then by first-appearance order from customer source files
    owner_order = {'GLOBAL': 0, 'LIST_NTCV': 1, 'LIST_VAT_LIEU': 2, 'LIST_YC_NTCV': 3}
    sorted_items = sorted(
        catalog.items(),
        key=lambda kv: (
            owner_order.get(kv[1]['owner'], 9),
            order_state['index'].get(kv[0], 10**9),
            kv[0],
        ),
    )
    catalog = {k: v for k, v in sorted_items}

    with open(CATALOG_PATH, 'w', encoding='utf-8') as f:
        json.dump(catalog, f, ensure_ascii=False, indent=2)
        f.write('\n')

    # Build demo JSON
    global_fields = [
        {'key': k, 'label': v.get('label') or k, 'value': v.get('sample', '')}
        for k, v in catalog.items() if v.get('owner') == 'GLOBAL'
    ]

    jobs = []

    # Add singleton/global jobs (from actual template files that are not list shared templates)
    list_tpls = set(target_group_templates.values())
    for tpl in sorted(existing_templates):
        if tpl in list_tpls:
            continue
        jobs.append({
            'template': tpl,
            'output': tpl,
            'enabled': True,
            'fields': [],
        })

    # Add grouped jobs by source files
    for owner in ('LIST_NTCV', 'LIST_VAT_LIEU', 'LIST_YC_NTCV'):
        target_tpl = target_group_templates[owner]
        for src_name in files_by_owner.get(owner, []):
            fields = []
            for k, v in sorted(list_job_values.get(src_name, {}).items()):
                if catalog.get(k, {}).get('owner') == owner:
                    fields.append({'key': k, 'label': catalog[k].get('label') or k, 'value': v})
            jobs.append({
                'template': target_tpl,
                'output': pretty_output_name(src_name),
                'enabled': True,
                'fields': fields,
            })

    demo = {
        'meta': {
            'version': 4,
            'description': 'Friendly list-based JSON for human input (updated from highlighted source docs)',
            'updated_at': __import__('datetime').datetime.now().strftime('%Y-%m-%d'),
        },
        'global_fields': global_fields,
        'jobs': jobs,
        'notes': {
            'source_folder': SRC_DIR,
            'group_templates': {
                'LIST_NTCV': target_group_templates['LIST_NTCV'],
                'LIST_VAT_LIEU': target_group_templates['LIST_VAT_LIEU'],
                'LIST_YC_NTCV': target_group_templates['LIST_YC_NTCV'],
            },
            'doc_files_skipped': [
                f for f in os.listdir(SRC_DIR)
                if f.lower().endswith('.doc')
                and not f.startswith('~$')
                and not os.path.exists(os.path.join(SRC_DIR, os.path.splitext(f)[0] + '.docx'))
            ],
            'doc_files_converted': converted_docx,
        },
    }

    with open(DEMO_PATH, 'w', encoding='utf-8') as f:
        json.dump(demo, f, ensure_ascii=False, indent=2)
        f.write('\n')

    # Summary markdown
    cnt = Counter(v['owner'] for v in catalog.values())
    md = []
    md.append('# Tổng hợp biến lặp từ các file Word (đợt mới)')
    md.append('')
    md.append(f'Nguồn: `{SRC_DIR}` (.docx)')
    md.append('')
    md.append('## Thống kê nhanh')
    md.append('')
    md.append(f'- GLOBAL: **{cnt.get("GLOBAL", 0)}** biến')
    md.append(f'- LIST_NTCV: **{cnt.get("LIST_NTCV", 0)}** biến')
    md.append(f'- LIST_VAT_LIEU: **{cnt.get("LIST_VAT_LIEU", 0)}** biến')
    md.append(f'- LIST_YC_NTCV: **{cnt.get("LIST_YC_NTCV", 0)}** biến')
    md.append(f'- Tổng: **{len(catalog)}** biến')
    md.append('')
    md.append('## Template nhóm dùng chung')
    md.append('')
    md.append(f'- LIST_NTCV → `{target_group_templates["LIST_NTCV"]}`')
    md.append(f'- LIST_VAT_LIEU → `{target_group_templates["LIST_VAT_LIEU"]}`')
    md.append(f'- LIST_YC_NTCV → `{target_group_templates["LIST_YC_NTCV"]}`')
    md.append('')
    md.append('## Ghi chú')
    md.append('')
    md.append('- Đã gom trùng theo chuẩn hóa: bỏ khác biệt hoa/thường, dấu câu và khoảng trắng dư.')
    md.append('- Với nhóm file cùng loại vẫn dùng chung 1 template như yêu cầu.')
    md.append('')

    for owner in ('GLOBAL', 'LIST_NTCV', 'LIST_VAT_LIEU', 'LIST_YC_NTCV'):
        md.append(f'## {owner} ({cnt.get(owner, 0)} biến)')
        md.append('')
        owner_keys = [k for k, v in catalog.items() if v['owner'] == owner]
        for k in owner_keys:
            v = catalog[k]
            md.append(f'- `{k}` — {v.get("label", "")}'.rstrip())
        md.append('')

    with open(SUMMARY_PATH, 'w', encoding='utf-8') as f:
        f.write('\n'.join(md).rstrip() + '\n')

    # Rebuild templates_bundle.zip
    os.makedirs(os.path.dirname(BUNDLE_PATH), exist_ok=True)
    with zipfile.ZipFile(BUNDLE_PATH, 'w', zipfile.ZIP_DEFLATED) as z:
        z.write(CATALOG_PATH, arcname='catalog.json')
        for tpl in sorted(os.listdir(TEMPLATES_DIR)):
            if tpl.lower().endswith('.docx'):
                z.write(os.path.join(TEMPLATES_DIR, tpl), arcname=f'templates/{tpl}')

    print('Done.')
    print('Catalog fields:', len(catalog), dict(cnt))


if __name__ == '__main__':
    main()
