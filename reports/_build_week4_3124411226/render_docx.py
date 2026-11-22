"""Use the packaged DOCX renderer with Microsoft Word as the PDF converter.

The Windows runtime bundles Poppler but does not bundle LibreOffice. This
adapter leaves the managed renderer and the input document unchanged.
"""
import importlib.util
import os
from pathlib import Path
import subprocess
import sys

SKILL = Path(r'C:\Users\ADMIN\.codex\plugins\cache\openai-primary-runtime\documents\26.1119.11507\skills\documents')
RUNTIME = Path(r'C:\Users\ADMIN\.cache\codex-runtimes\codex-primary-runtime\dependencies')
POPPLER = RUNTIME / 'native/poppler/Library/bin'
os.environ['PATH'] = str(POPPLER) + os.pathsep + os.environ.get('PATH', '')
spec = importlib.util.spec_from_file_location('packaged_renderer', SKILL / 'render_docx.py')
renderer = importlib.util.module_from_spec(spec)
spec.loader.exec_module(renderer)

def word_to_pdf(doc_path, user_profile, convert_tmp_dir, stem, verbose):
    target = str(Path(convert_tmp_dir) / (stem + '.pdf'))
    cmd = ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
           str(Path(__file__).with_name('export_word.ps1')), '-InputDocx',
           str(Path(doc_path).resolve()), '-OutputPdf', target]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=180)
    logs = result.stdout + result.stderr
    if result.returncode or not Path(target).exists():
        raise RuntimeError(logs)
    if verbose:
        print(logs)
    return target, logs

renderer.convert_to_pdf = word_to_pdf
if __name__ == '__main__':
    renderer.main()
