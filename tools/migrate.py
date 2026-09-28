"""One-time migration: the pages downloaded from mezim.netlify.app -> this repo.

Usage: python tools/migrate.py <dir-with-downloaded-page_*.html>
Rewrites absolute links to relative ones (so the site works on GitHub Pages
with or without a custom domain), pulls the inline avatar out to assets/,
and creates the copies of the מאיצים tools under data/.
"""
import base64, os, re, sys

SRC = sys.argv[1]
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

PAGES = ['ai', 'ai-prompt', 'maazim', 'role-worksheet', 'problem-tree',
         'challenge', 'prompt', 'insight', 'prototype', 'pitch']
# tools copied into the data-driven-decisions section, with their day there
DATA_COPIES = {
    'role-worksheet': 'יום 1 · שאלה',
    'problem-tree':   'יום 1 · שאלה',
    'challenge':      'יום 1 · שאלה',
    'prompt':         'יום 2 · תובנה',
    'insight':        'יום 2 · תובנה',
    'prototype':      'יום 3 · החלטה',
    'pitch':          'יום 3 · החלטה',
}
NETLIFY_COMMENT = re.compile(r'<!-- This site is hosted on Netlify.*?-->\n?', re.S)
AVATAR = re.compile(r'data:image/jpeg;base64,([A-Za-z0-9+/=]+)')


def read(name):
    with open(os.path.join(SRC, f'page_{name}.html'), encoding='utf-8') as f:
        return f.read()


def write(rel, text):
    path = os.path.join(ROOT, rel)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8', newline='\n') as f:
        f.write(text)


def fix_avatar(html, up):
    m = AVATAR.search(html)
    if m:
        img = os.path.join(ROOT, 'assets', 'roman.jpg')
        if not os.path.exists(img):
            os.makedirs(os.path.dirname(img), exist_ok=True)
            with open(img, 'wb') as f:
                f.write(base64.b64decode(m.group(1)))
        html = AVATAR.sub(up + 'assets/roman.jpg', html)
    return html


def relink(html, up):
    """href='/x?q' -> href='{up}x/?q' ; href='/' -> href='{up}'"""
    def sub(m):
        q, path = m.group(1), m.group(2)
        if path == '':
            return f'href={q}{up or "./"}{q}'
        name, _, query = path.partition('?')
        return f'href={q}{up}{name}/{"?" + query if query else ""}{q}'
    return re.sub(r"""href=(['"])/([^'"/][^'"]*|)\1""", sub, html)


ROLE_SETHUB = '''  function setHub(){
    try{
      var p=new URLSearchParams(location.search).get('from');
      var hub = p==='maazim' ? '../maazim/' : '../ai/';
      var el=document.getElementById('hubLink'); if(el) el.setAttribute('href',hub);
    }catch(e){}
  }'''


def main():
    for name in PAGES + ['root']:
        html = NETLIFY_COMMENT.sub('', read(name))
        up = '' if name == 'root' else '../'
        html = relink(fix_avatar(html, up), up)
        if name == 'role-worksheet':
            html = re.sub(r'  function setHub\(\)\{.*?\n  \}\n', ROLE_SETHUB + '\n', html, count=1, flags=re.S)
        write('index.html' if name == 'root' else f'{name}/index.html', html)

        if name in DATA_COPIES:
            d = NETLIFY_COMMENT.sub('', read(name))
            d = fix_avatar(d, '../../')
            # every internal link in a tool points back to its hub
            d = re.sub(r"""href=(['"])/[^'"]*\1""", r"href=\1../\1", d)
            if name == 'role-worksheet':
                d = re.sub(r'  function setHub\(\)\{.*?\n  \}\n', '  function setHub(){}\n', d, count=1, flags=re.S)
            # separate saved work from the מאיצים copies of the same tool
            d = re.sub(r"(LS(?:_KEY)?\s*=\s*')", r"\1dd_", d)
            d = re.sub(r"(CLOUD_KEY\s*=\s*')", r"\1dd_", d)
            d = re.sub(r'(<span class="daytag">)[^<]*', r'\g<1>' + DATA_COPIES[name], d)
            d = re.sub(r'<title>([^<—]*?)\s*—[^<]*</title>', r'<title>\1 — קבלת החלטות מבוססת נתונים</title>', d)
            write(f'data/{name}/index.html', d)


if __name__ == '__main__':
    main()
