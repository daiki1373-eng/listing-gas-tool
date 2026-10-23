"""Bundle local source assets into a single Apps Script HTML Service file."""
from pathlib import Path
import re
import json

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / 'src'


def bundle(templates=None):
    html = (SRC / 'index.html').read_text(encoding='utf-8')

    def css(match):
        path = (SRC / match.group(1)).resolve()
        if not path.is_relative_to(SRC.resolve()):
            raise ValueError('Asset path escapes src/')
        return '<style>' + path.read_text(encoding='utf-8') + '</style>'

    def js(match):
        path = (SRC / match.group(1)).resolve()
        if not path.is_relative_to(SRC.resolve()):
            raise ValueError('Asset path escapes src/')
        source = path.read_text(encoding='utf-8')
        if '</script' in source.lower():
            raise ValueError('Inline script contains a closing script tag: ' + path.name)
        return '<script>\n' + source + '\n</script>'

    html = re.sub(r'<link rel="stylesheet" href="([^"]+)">', css, html)
    html = re.sub(r'<script src="([^"]+)"></script>', js, html)
    if templates is not None:
        marker = 'window.REPORT_TEMPLATES={};'
        if html.count(marker) != 1:
            raise ValueError('Template initialization marker is missing or duplicated')
        html = html.replace(marker, 'window.REPORT_TEMPLATES=' + json.dumps(templates, separators=(',', ':')) + ';', 1)
    if re.search(r'<(?:script[^>]* src=|link[^>]*stylesheet)', html):
        raise ValueError('External assets remain in GAS output')
    return html


if __name__ == '__main__':
    output = ROOT / 'gas' / 'Index.html'
    output.write_text(bundle(), encoding='utf-8')
    print(f'Built gas/Index.html ({output.stat().st_size:,} bytes)')
