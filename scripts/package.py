"""Create a self-contained transfer ZIP using locally supplied PPT originals."""
import argparse
import base64
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from build import ROOT, bundle

NAMES = {'single': 'ListingTemplateSingle.pptx', 'multi': 'ListingTemplateMulti.pptx'}


def template_bytes(folder):
    result = {}
    for mode, name in NAMES.items():
        data = (folder / name).read_bytes()
        if len(data) > 20 * 1024 * 1024:
            raise ValueError(f'{name}: template must be 20MB or smaller')
        with ZipFile(folder / name) as z:
            count = 26 if mode == 'single' else 25
            for page in range(1, count + 1):
                if f'ppt/slides/slide{page}.xml' not in z.namelist():
                    raise ValueError(f'{name}: incompatible slide structure')
            marker = '日別予算到達状況' if mode == 'single' else 'キャンペーン別 当月実績'
            if f'name="{marker}"' not in z.read('ppt/slides/slide14.xml').decode('utf-8'):
                raise ValueError(f'{name}: single/multi template does not match')
            if z.testzip():
                raise ValueError(f'{name}: corrupted ZIP entry')
        result[mode] = data
    return result


def files_for_package(folder):
    originals = template_bytes(folder)
    files = {}
    for prefix in ['src', 'scripts', 'gas', 'tests', 'examples', 'docs']:
        for path in (ROOT / prefix).rglob('*'):
            if path.is_file() and '__pycache__' not in path.parts and path.suffix != '.pyc':
                files[path.relative_to(ROOT).as_posix()] = path.read_bytes()
    for name in ['README.md', '.gitignore', 'package.json']:
        files[name] = (ROOT / name).read_bytes()
    files['gas/Index.html'] = bundle().encode('utf-8')
    files['start.html'] = bundle({mode:base64.b64encode(data).decode('ascii') for mode,data in originals.items()}).encode('utf-8')
    for mode, data in originals.items():
        files['templates/' + NAMES[mode]] = data
    version = json.loads((ROOT / 'package.json').read_text(encoding='utf-8'))['version']
    files['package-info.json'] = json.dumps({'version':version,'offline':'start.html','gas':'gas/Index.html','templates':list(NAMES.values())},ensure_ascii=False,indent=2).encode('utf-8')
    files['SHA256.json'] = json.dumps({name:hashlib.sha256(data).hexdigest() for name,data in sorted(files.items())},ensure_ascii=False,indent=2).encode('utf-8')
    return files


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--templates-dir', type=Path, required=True)
    parser.add_argument('--output', type=Path, default=ROOT/'dist/listing-gas-tool-portable.zip')
    args = parser.parse_args()
    files = files_for_package(args.templates_dir.resolve())
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(args.output, 'w', ZIP_DEFLATED, compresslevel=9) as z:
        for name, data in sorted(files.items()):
            z.writestr('listing-gas-tool/' + name, data)
    with ZipFile(args.output) as z:
        assert z.testzip() is None
        manifest = json.loads(z.read('listing-gas-tool/SHA256.json'))
        for name, digest in manifest.items():
            assert hashlib.sha256(z.read('listing-gas-tool/' + name)).hexdigest() == digest
    print(f'Created {args.output} ({len(files)} files, {args.output.stat().st_size:,} bytes)')
