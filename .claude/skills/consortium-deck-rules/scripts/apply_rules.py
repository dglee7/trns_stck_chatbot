# Applies the 협의체 deck formatting rules (see ../SKILL.md). Usage: python3 apply_rules.py in.pptx out.pptx
import re, zipfile, sys, copy
SRC, DST = sys.argv[1], sys.argv[2]
sys.argv[3:] = ['/tmp/_slide2_ref.xml']
GRAY = '<a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill>'
def line(tag):  # 0.5pt gray
    return f'<a:{tag} w="6350" cap="flat" cmpd="sng" algn="ctr">{GRAY}<a:prstDash val="solid"/></a:{tag}>'
def set_ln(tcpr, tag):
    pat = re.compile(rf'<a:{tag}\b.*?</a:{tag}>', re.S)
    assert pat.search(tcpr), tag
    return pat.sub(lambda m: line(tag), tcpr, count=1)
def fix_cell_border(c, tag):
    return re.sub(r'<a:tcPr.*?</a:tcPr>', lambda m: set_ln(m.group(0), tag), c, count=1, flags=re.S)
def is_hdr(row): return 'DDEBF7' in row
def ctext(c): return ''.join(re.findall(r'<a:t>([^<]*)</a:t>', c))
stats = {'hdr':0,'val':0,'lbl':0}

def fix_table(tb, pp_all):
    rows = re.findall(r'<a:tr .*?</a:tr>', tb, re.S)
    nh = 0
    while nh < len(rows) and is_hdr(rows[nh]): nh += 1
    new = []
    for ri, r in enumerate(rows):
        cells = re.findall(r'<a:tc[ >].*?</a:tc>', r, re.S)
        out = []
        for ci, c in enumerate(cells):
            if ri == nh-1: c = fix_cell_border(c, 'lnB'); stats['hdr'] += 1
            if ri == nh:   c = fix_cell_border(c, 'lnT')
            if ri < nh:    # label
                c2 = re.sub(r'<a:t>변화', '<a:t>증감', c)
                stats['lbl'] += c2 != c; c = c2
            else:
                label = ctext(cells[0])
                pp = pp_all or '비중' in label
                def rep(m):
                    t = m.group(1)
                    if t.startswith('▼'): t = '▲' + t[1:]
                    elif t.startswith('▲'): t = ('+' if pp else '') + t[1:].lstrip()
                    stats['val'] += t != m.group(1)
                    return f'<a:t>{t}</a:t>'
                c = re.sub(r'<a:t>([▲▼][^<]*)</a:t>', rep, c)
            out.append(c)
        # rebuild row
        i = 0
        def sub(m):
            nonlocal i; v = out[i]; i += 1; return v
        new.append(re.sub(r'<a:tc[ >].*?</a:tc>', sub, r, flags=re.S))
    j = 0
    def sub2(m):
        nonlocal j; v = new[j]; j += 1; return v
    return re.sub(r'<a:tr .*?</a:tr>', sub2, tb, flags=re.S)

def process(n, xml):
    first = [True]
    def tbl(m):
        tb = m.group(0)
        pp_all = (n == 1)
        return fix_table(tb, pp_all)
    xml = re.sub(r'<a:tbl>.*?</a:tbl>', tbl, xml, flags=re.S)
    # unit boxes: right-align to table right edge
    if n == 1:
        sp = [b for b in re.findall(r'<p:sp>.*?</p:sp>', open(sys.argv[3], encoding='utf8').read(), re.S) if 'name="TextBox 10"' in b][0]
        runs = re.findall(r'<a:r>.*?</a:r>', sp, re.S)
        sp = sp.replace(runs[0], re.sub(r'<a:t>[^<]*</a:t>', '<a:t>(단위: %, %p)</a:t>', runs[0]), 1)
        for r_ in runs[1:]: sp = sp.replace(r_, '', 1)
        ids = [int(x) for x in re.findall(r'<p:cNvPr id="(\d+)"', xml)]
        sp = re.sub(r'<p:cNvPr id="\d+" name="[^"]*"', f'<p:cNvPr id="{max(ids)+1}" name="TextBox Unit"', sp, count=1)
        # right edge = table right edge
        m = re.search(r'<p:graphicFrame>.*?<a:off x="(\d+)" y="(\d+)"/>\s*<a:ext cx="(\d+)" cy="(\d+)"', xml, re.S)
        tx, ty, tcx, tcy = map(int, m.groups())
        sp = re.sub(r'<a:off x="\d+" y="\d+"/>\s*<a:ext cx="\d+" cy="(\d+)"/>',
                    lambda k: f'<a:off x="{tx}" y="1911553"/><a:ext cx="{tcx}" cy="{k.group(1)}"/>', sp, count=1)
        xml = xml.replace('</p:spTree>', sp + '</p:spTree>', 1)
    else:
        # snap every right-aligned unit box to the right edge of its table
        frames = [tuple(map(int, g)) for g in re.findall(r'<p:graphicFrame>.*?<a:off x="(\d+)" y="(\d+)"/>\s*<a:ext cx="(\d+)" cy="(\d+)"', xml, re.S)]
        def snap(m):
            b = m.group(0)
            if '(단위' not in b or 'algn="r"' not in b: return b
            om = re.search(r'<a:off x="(\d+)" y="(\d+)"/>\s*<a:ext cx="(\d+)" cy="(\d+)"', b)
            x, y, cx, cy = map(int, om.groups())
            # table directly below the unit box
            cand = [f for f in frames if f[1] >= y and abs(f[0]-x) < 100000]
            f = min(cand, key=lambda f: f[1])
            nx = f[0] + f[2] - cx
            return b.replace(om.group(0), f'<a:off x="{nx}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"')
        xml = re.sub(r'<p:sp>.*?</p:sp>', snap, xml, flags=re.S)
    return xml

zin = zipfile.ZipFile(SRC)
slide2 = zin.read('ppt/slides/slide2.xml').decode('utf8')
open('/tmp/_slide2_ref.xml','w',encoding='utf8').write(slide2)
sys.argv[3] = '/tmp/_slide2_ref.xml'
zout = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    data = zin.read(it.filename)
    m = re.fullmatch(r'ppt/slides/slide(\d+)\.xml', it.filename)
    if m: data = process(int(m.group(1)), data.decode('utf8')).encode('utf8')
    zout.writestr(it, data)
zout.close()
print(stats)
