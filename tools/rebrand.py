"""One-time recolor of every page to the RomanAI design system (tokens.css).

The pages carry their palette inline (CSS variables + literal hex values, also inside
SVG icons and the canvas code that draws exported images), so the old palette is
mapped to the new one value by value. `--report` only lists emoji left in the pages.
"""
import glob, os, re, sys, collections

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = [p for pat in ('**/index.html', '**/assets/*.css', '**/assets/*.js')
         for p in glob.glob(os.path.join(ROOT, pat), recursive=True)]

HEX = {  # old -> new
    # surfaces / text / lines
    'EEF1F6': 'FAF5EC', 'F7F9FB': 'F2EADC', 'F6F8FB': 'FAF5EC', 'D5DBE3': 'E3D8C6', 'DCE2EA': 'E3D8C6',
    'E6EAF0': 'EDE4D5', '15242E': '2A211B', '54636F': '6B5D52', '1F3540': '2A211B', '0C1820': '1C1512',
    'B3BDC9': 'CBBDA6', 'C2CCD6': 'CBBDA6', 'C5CED7': 'D6CBBB', 'B8C2CC': 'D6CBBB', '9AA6B1': 'A3978A',
    # primary action colour: old green-teal -> terracotta
    '127C71': 'B4502F', '0E655C': '8F3D22',
    # secondary: old blue -> deep teal
    '1F5673': '1F5F5B', '16405C': '174A47', '1F4460': '1F5F5B',
    # accents
    'C77D3E': 'D99A2B', 'A9662E': '7D5210', 'E9A94F': 'F0B95A', '8A7358': '6B5D52',
    '2E8B6B': '2E7D4F', 'B23A2C': 'A8362D', 'E4463A': 'A8362D',
    # dark hero backgrounds -> teal
    '0C2233': '1F5F5B', '123040': '1F5F5B', '10403B': '1F5F5B', '102A43': '1F5F5B', '0F2230': '174A47',
    # light-on-dark
    '4FD6BE': 'F0B95A', '7FD8C6': 'C9E2DE', '8FD6C2': 'C9E2DE', 'CFE6DF': 'C9E2DE', 'E8F1F0': 'FFFDF8',
    # tints
    'E7F1EE': 'DCEDEA', 'EAF0F5': 'DCEDEA', 'EEF3F8': 'DCEDEA', 'C9D7E5': 'A9D3CC', 'FBF1E6': 'F4DED3',
    'F5FBFA': 'FBF1EA', 'F0FAF8': 'FBF1EA', 'F4FAF8': 'FBF1EA',
}
RGBA = {'18,124,113': '180,80,47', '21,36,46': '42,33,27', '31,86,115': '31,95,91',
        '46,139,107': '46,125,79', '247,249,251': '250,245,236', '238,241,246': '250,245,236',
        '79,214,190': '240,185,90', '233,169,79': '217,154,43'}
# the brand uses no emoji; plain marks (✓ ✗ ✕ ☐ ☑ ★) stay
EMOJI = re.compile('[\U0001F000-\U0001FAFF☀-☄☆-☏☒-✒✔✖✘-➿⬀-⯿]️?[  ]?')
SWAP = {'🤖': 'AI', '⭐': '★', '➕': '+', '❝': '"'}


def strip_emoji(s):
    for a, b in SWAP.items():
        s = s.replace(a + '️', b).replace(a, b)
    return EMOJI.sub('', s).replace('️', '')


def recolor(s):
    s = re.sub(r'#([0-9a-fA-F]{6})\b', lambda m: '#' + HEX.get(m.group(1).upper(), m.group(1)), s)
    s = re.sub(r'#(?:fff|FFF)\b(?![0-9a-fA-F])', '#FFFDF8', s)          # cards are warm white
    for a, b in RGBA.items():
        s = re.sub(r'rgba\(\s*' + a.replace(',', r'\s*,\s*'), 'rgba(' + b, s)
    return s


if __name__ == '__main__':
    if '--report' in sys.argv:
        sys.stdout.reconfigure(encoding='utf-8')
        for f in FILES:
            c = collections.Counter(EMOJI.findall(open(f, encoding='utf-8').read()))
            if c: print(os.path.relpath(f, ROOT), ' '.join(k + str(v) for k, v in c.most_common()))
    else:
        for f in FILES:
            s = open(f, encoding='utf-8').read(); t = strip_emoji(recolor(s))
            if t != s: open(f, 'w', encoding='utf-8', newline='\n').write(t)
        print('recolored', len(FILES), 'files')
