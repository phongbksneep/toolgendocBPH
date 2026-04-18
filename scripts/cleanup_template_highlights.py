#!/usr/bin/env python3
import os
import re
import zipfile

ROOT = '/root/.openclaw/workspace/bphh_docgen'
TPL_DIR = os.path.join(ROOT, 'templates')
BUNDLE = os.path.join(ROOT, 'assets', 'templates_bundle.zip')
CATALOG = os.path.join(ROOT, 'catalog.json')

RUN_RE = re.compile(r'<w:r\b[\s\S]*?</w:r>')
HIGHLIGHT_RE = re.compile(r'<w:highlight[^>]*/>')
TEXT_RE = re.compile(r'<w:t(?:\s[^>]*)?>([\s\S]*?)</w:t>')


def cleanup_docx(path):
    with zipfile.ZipFile(path, 'r') as zin:
        items = {n: zin.read(n) for n in zin.namelist()}

    changed = False
    for name, data in list(items.items()):
        if not (name.startswith('word/') and name.endswith('.xml')):
            continue
        xml = data.decode('utf-8', 'ignore')

        def repl_run(m):
            run = m.group(0)
            if '<w:highlight' not in run:
                return run
            txt = ''.join(t.group(1) for t in TEXT_RE.finditer(run)).strip()
            # keep highlight only if this run contains placeholder
            if '{{' in txt and '}}' in txt:
                return run
            new_run = HIGHLIGHT_RE.sub('', run)
            return new_run

        new_xml = RUN_RE.sub(repl_run, xml)
        if new_xml != xml:
            changed = True
            items[name] = new_xml.encode('utf-8')

    if changed:
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for n, b in items.items():
                zout.writestr(n, b)
    return changed


def rebuild_bundle():
    with zipfile.ZipFile(BUNDLE, 'w', zipfile.ZIP_DEFLATED) as z:
        z.write(CATALOG, arcname='catalog.json')
        for fn in sorted(os.listdir(TPL_DIR)):
            if fn.lower().endswith('.docx'):
                z.write(os.path.join(TPL_DIR, fn), arcname=f'templates/{fn}')


def main():
    changed = []
    for fn in sorted(os.listdir(TPL_DIR)):
        if not fn.lower().endswith('.docx'):
            continue
        p = os.path.join(TPL_DIR, fn)
        if cleanup_docx(p):
            changed.append(fn)
    rebuild_bundle()
    print('cleaned templates:', len(changed))
    for x in changed:
        print('-', x)


if __name__ == '__main__':
    main()
