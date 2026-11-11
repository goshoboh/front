#!/usr/bin/env python3
"""Generate the PWA release version from deployable files (no dependencies)."""
import argparse
import hashlib
import json
import re
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true', help='Verify without writing files')
    args = parser.parse_args()
    root = Path(__file__).resolve().parent.parent
    index = root / 'index.html'
    html = index.read_bytes()
    pattern = rb'(<meta name="app-version" content=")[^"]*("\s*/?>)'
    if len(re.findall(pattern, html)) != 1:
        parser.error('index.html must contain exactly one app-version meta tag')
    normalized = re.sub(pattern, rb'\g<1>__APP_VERSION__\g<2>', html)
    assets = ['index.html', 'app.js', 'styles.css', 'pwa-update.js', 'sw.js',
              'manifest.webmanifest', 'favicon.ico']
    assets += [p.relative_to(root).as_posix() for p in (root / 'icons').rglob('*') if p.is_file()]
    digest = hashlib.sha256()
    for name in sorted(assets):
        content = normalized if name == 'index.html' else (root / name).read_bytes()
        digest.update(name.encode() + b'\0' + str(len(content)).encode() + b'\0' + content)
    version = digest.hexdigest()
    updated_html = re.sub(pattern, lambda m: m[1] + version.encode() + m[2], html)
    release = (json.dumps({'version': version}, indent=2) + '\n').encode()
    version_file = root / 'version.json'
    if args.check:
        if html != updated_html or not version_file.exists() or version_file.read_bytes() != release:
            parser.exit(1, 'Release version is stale. Run: python3 scripts/prepare-release.py\n')
        print('Release version is current: ' + version[:12])
        return
    index.write_bytes(updated_html)
    version_file.write_bytes(release)
    print('Prepared release: ' + version[:12])


if __name__ == '__main__':
    main()
