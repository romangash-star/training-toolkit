"""Add the site-wide stylesheet + script (back strip, animated drawings) to every page. Idempotent."""
import glob, os, re
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
    depth = rel.count('/'); up = '../' * depth
    s = open(path, encoding='utf-8').read()
    s = re.sub(r'\n?<link rel="stylesheet" href="[./]*assets/site\.css">', '', s)
    s = re.sub(r'\n?<script src="[./]*assets/site\.js"[^>]*></script>', '', s)
    c = crumbs(rel)
    is_tool = depth == 2 or (depth == 1 and rel.split('/')[0] not in ('data', 'ai', 'maazim'))
    attrs = (f" data-crumbs='{c}'" if c else '') + (' data-side="1"' if is_tool else '')
    s = s.replace('</head>', f'<link rel="stylesheet" href="{up}assets/site.css">\n</head>', 1)
    i = s.rindex('</body>')
    s = s[:i] + f'<script src="{up}assets/site.js"{attrs}></script>\n' + s[i:]
    open(path, 'w', encoding='utf-8', newline='\n').write(s)
    print(rel, attrs)
