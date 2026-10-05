"""Add the site-wide stylesheet + script (the RomanAI shell) to every page. Idempotent.

Bump V whenever a shared asset changes, so a browser never runs a new page
with an old cached script.
"""
import glob, os, re

V = '17'
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HOME, DD, MZ, AI = 'כל ההכשרות', 'קבלת החלטות מבוססת נתונים', 'מאיצים דיגיטליים', 'הכשרת בינה מלאכותית'


def crumbs(rel):
    parts = rel.split('/')[:-1]            # folders leading to index.html
    if not parts: return None
    if parts == ['data'] or parts in (['ai'], ['maazim']): return f'{HOME}|../'
    if parts[0] == 'data': return f'{HOME}|../../;{DD}|../'
    if parts[0] == 'role-worksheet': return f'{HOME}|../;x|auto'
    if parts[0] == 'ai-prompt': return f'{HOME}|../;{AI}|../ai/'
    return f'{HOME}|../;{MZ}|../maazim/'


for path in glob.glob(os.path.join(ROOT, '**', 'index.html'), recursive=True):
    rel = '/'.join(os.path.relpath(path, ROOT).split(os.sep))
    up = '../' * rel.count('/')
    s = open(path, encoding='utf-8').read()
    s = re.sub(r'\n?<link rel="stylesheet" href="[./]*assets/site\.css[^"]*">', '', s)
    s = re.sub(r'\n?<script src="[./]*assets/site\.js[^"]*"[^>]*></script>', '', s)
    s = re.sub(r'(assets/(?:dd|hub)\.(?:css|js))(?:\?v=\w+)?"', r'\1?v=' + V + '"', s)
    c = crumbs(rel)
    attrs = f" data-crumbs='{c}'" if c else ''
    s = re.sub(r'\n?<link rel="(?:icon|apple-touch-icon)"[^>]*>', '', s)
    icons = (f'<link rel="icon" href="{up}assets/favicon.svg" type="image/svg+xml">\n'
             f'<link rel="icon" href="{up}assets/favicon-32.png" sizes="32x32" type="image/png">\n'
             f'<link rel="apple-touch-icon" href="{up}assets/apple-touch-icon.png">\n')
    s = s.replace('</head>', icons + f'<link rel="stylesheet" href="{up}assets/site.css?v={V}">\n</head>', 1)
    i = s.rindex('</body>')
    s = s[:i] + f'<script src="{up}assets/site.js?v={V}"{attrs}></script>\n' + s[i:]
    open(path, 'w', encoding='utf-8', newline='\n').write(s)
