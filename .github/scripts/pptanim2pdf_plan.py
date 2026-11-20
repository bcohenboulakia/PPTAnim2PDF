"""Select tracked presentations for automatic or manual conversion."""
import argparse
import json
from pathlib import Path
import subprocess


def paths(*args):
    return set(subprocess.check_output(['git', *args]).decode('utf-8').split('\0')) - {''}


def plan(event, before='', input_path='', output_path=''):
    current = {p for p in paths('ls-tree', '-r', '--name-only', '-z', 'HEAD')
               if p.lower().endswith('.pptx')}
    if event != 'push':
        if input_path not in current:
            raise ValueError(f'Tracked source presentation not found: {input_path}')
        selected, deleted = {input_path}, set()
    else:
        changed = paths('diff', '--no-renames', '--name-only', '-z', before, 'HEAD') if before and set(before) != {'0'} else current
        tooling = any(p.startswith('.github/pptanim2pdf/') or p in {
            '.github/workflows/convert-pptx.yml', '.github/scripts/pptanim2pdf_plan.py',
            '.github/scripts/publish_outputs.py'} for p in changed)
        selected = current if tooling else current & changed
        deleted = {p for p in changed - current if p.lower().endswith('.pptx')}
    return {
        'matrix': {'include': [{'input': p, 'output': output_path if event != 'push' else ''}
                               for p in sorted(selected)]},
        'removed_matrix': {'include': [{'input': p, 'output': str(Path(p).with_suffix('.pdf')).replace('\\', '/')}
                                       for p in sorted(deleted)]},
        'has_conversions': bool(selected), 'has_removals': bool(deleted),
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--event', required=True)
    parser.add_argument('--before', default='')
    parser.add_argument('--input', default='')
    parser.add_argument('--output', default='')
    args = parser.parse_args()
    result = plan(args.event, args.before, args.input, args.output)
    print(json.dumps(result))
    import os
    if 'GITHUB_OUTPUT' in os.environ:
        with open(os.environ['GITHUB_OUTPUT'], 'a', encoding='utf-8') as stream:
            for key, value in result.items():
                stream.write(f'{key}={json.dumps(value)}\n')
            revision = subprocess.check_output(['git', 'rev-parse', 'HEAD']).decode().strip()
            stream.write(f'source_sha={revision}\n')


if __name__ == '__main__':
    main()
