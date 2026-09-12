#!/usr/bin/env python3
"""Package this skill only; never include adjacent user libraries or projects."""
import argparse
import hashlib
import json
from pathlib import Path
import re
import zipfile


REPO_ROOT = Path(__file__).resolve().parents[3]
DISTRIBUTION = REPO_ROOT / 'skill'


def build(output):
    root = Path(__file__).resolve().parents[1]
    output = output.resolve()
    if output == root or root in output.parents:
        raise ValueError('Choose an output directory outside the skill source')
    files = sorted(p for p in root.rglob('*') if p.is_file()
                   and '__pycache__' not in p.parts and p.suffix != '.pyc')
    for p in files:
        if p.suffix == '.md':
            for target in re.findall(r'\]\(([^)]+)\)', p.read_text(encoding='utf-8')):
                if '://' not in target and not target.startswith('#'):
                    if not (p.parent / target.split('#')[0]).exists():
                        raise ValueError(f'Missing resource: {p.name}: {target}')
    output.mkdir(parents=True, exist_ok=True)
    parts = ['# AutoBericht — complete portable workflow\n\n'
             'Follow this workflow when the user asks. All skill resources are '
             'embedded below under their original relative filenames. Read '
             'SKILL.md first, then the relevant resources. Relative resource '
             'links refer to the corresponding embedded sections. Materialise '
             'helper code at the named paths only when execution is needed and '
             'available. The user supplies permitted project inputs separately. '
             'Site photos stay on the work computer.\n']
    ordered = [root / 'SKILL.md'] + [p for p in files if p != root / 'SKILL.md']
    for p in ordered:
        parts.append('\n\n---\n\n## Embedded resource: '
                     + p.relative_to(root).as_posix() + '\n\n')
        text = p.read_text(encoding='utf-8')
        if p.suffix == '.md':
            parts.append(text)
        else:
            lang = {'.py': 'python', '.cjs': 'javascript', '.yaml': 'yaml'}.get(p.suffix, 'text')
            parts.append('````' + lang + '\n' + text.rstrip() + '\n````\n')
    portable = output / 'AutoBericht_portable.md'
    portable.write_text(''.join(parts), encoding='utf-8', newline='\n')
    archive = output / 'autobericht-skill.zip'
    with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as z:
        for p in files:
            # Fixed metadata makes identical source produce identical ZIP bytes.
            info = zipfile.ZipInfo('autobericht/' + p.relative_to(root).as_posix(),
                                   date_time=(1980, 1, 1, 0, 0, 0))
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, p.read_bytes())
    with zipfile.ZipFile(archive) as z:
        if z.testzip() is not None:
            raise ValueError('Archive integrity check failed')
        for p in files:
            if z.read('autobericht/' + p.relative_to(root).as_posix()) != p.read_bytes():
                raise ValueError('Archive content mismatch')
    manifest = {p.name: {'bytes': p.stat().st_size,
                        'sha256': hashlib.sha256(p.read_bytes()).hexdigest()}
                for p in [portable, archive]}
    manifest_path = (REPO_ROOT / 'program/build/skill-manifest.json'
                     if output == DISTRIBUTION else output / 'package_manifest.json')
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8', newline='\n')
    print(json.dumps({'sourceFiles': len(files), 'artifacts': manifest}, indent=2))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, default=DISTRIBUTION,
                        help='Output directory (default: repo skill/)')
    build(parser.parse_args().output)
