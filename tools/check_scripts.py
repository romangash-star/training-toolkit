"""Syntax-check every inline <script> and every shared .js file with `node --check`."""
import glob, os, re, subprocess, tempfile, sys
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
bad = 0
pages = glob.glob('**/index.html', recursive=True)
tmp = os.path.join(tempfile.gettempdir(), 'site_check.js')
for f in pages:
    for js in re.findall(r'<script>(.*?)</script>', open(f, encoding='utf-8').read(), flags=re.S):
        open(tmp, 'w', encoding='utf-8').write(js)
        r = subprocess.run(['node', '--check', tmp], capture_output=True, text=True)
        if r.returncode: bad += 1; print('SYNTAX', f, r.stderr[:300])
for f in glob.glob('**/assets/*.js', recursive=True):
    r = subprocess.run(['node', '--check', f], capture_output=True, text=True)
    if r.returncode: bad += 1; print('SYNTAX', f, r.stderr[:300])
print(len(pages), 'pages checked,', bad, 'syntax errors')
sys.exit(1 if bad else 0)
