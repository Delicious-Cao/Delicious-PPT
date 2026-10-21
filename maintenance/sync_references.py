"""Synchronize shared Skill references and check repository-local Markdown links."""

import argparse
from pathlib import Path
import re
import sys

ROOT = Path(__file__).resolve().parent.parent


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument('--check', action='store_true')
    mode.add_argument('--write', action='store_true')
    args = parser.parse_args()
    source = ROOT / 'skills/delicious-ppt-chatgpt/references'
    target = ROOT / 'skills/delicious-ppt-deepseek/references'
    errors = []
    pairs = [(source / name, target / name) for name in
             ['requirements-intake.md', 'visual-design.md', 'delivery-checklist.md']]
    for directory in [source, target]:
        pairs.append((ROOT / 'docs/RUNTIME.md', directory / 'runtime.md'))
    for origin, destination in pairs:
        if args.write:
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(origin.read_bytes())
        if not destination.exists() or origin.read_bytes() != destination.read_bytes():
            errors.append('Reference mismatch: ' + destination.relative_to(ROOT).as_posix())
    for file in ROOT.rglob('*.md'):
        for match in re.finditer(r'\]\((<[^>]+>|[^)\s]+)\)', file.read_text(encoding='utf-8')):
            href = match.group(1).strip('<>')
            if href.startswith(('https://', 'http://', '#', 'mailto:')):
                continue
            relative = href.split('#', 1)[0]
            if not (file.parent / relative).exists():
                errors.append(file.relative_to(ROOT).as_posix() + ': broken link ' + href)
    if errors:
        print('\n'.join(errors), file=sys.stderr)
        return 1
    print('Shared references and local Markdown links are consistent.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
