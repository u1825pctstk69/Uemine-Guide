"""台ごとのQR（印刷用PDF）を作る

使い方:
  python3 tools/qr/make_qr.py <レート...> [-o 出力.pdf]
    例: python3 tools/qr/make_qr.py 5S
        python3 tools/qr/make_qr.py 1P 5S -o 赤枠.pdf

- data/data.json の指定レートの全台を台番号順に並べる（QRの中身は dai/?no=台番号）
- 1枚 縦70mm×横60mm。枠色はレートで決まる（4円・20円＝緑、1円・5円＝赤）
- 上から「N番台」→ QR → 「この台の機種ガイドになります」→ 下の色帯に店ロゴ（tools/qr/logo_white.png）
- A4に3列×4行=12枚。Chromium（Playwright）で HTML→PDF
- 作成後、全QRを1枚ずつ切り出して読み取り、台番号と一致するか確認する
"""
import sys, os, io, json, base64, glob, subprocess, tempfile, argparse
import qrcode

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
URL = 'https://u1825pctstk69.github.io/Uemine-Guide/dai/?no={}'
GREEN, RED = '#23AC38', '#E60012'
COLOR = {'4P': GREEN, '20S': GREEN, '1P': RED, '5S': RED}
NAME = {'4P': '4円パチンコ', '1P': '1円パチンコ', '20S': '20円スロット', '5S': 'プラスファイブ'}
PER = 12

CSS = '''@page{size:A4;margin:5mm 0}body{margin:0;font-family:"Noto Sans CJK JP",sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.g{display:grid;grid-template-columns:repeat(3,60mm);grid-auto-rows:70mm;gap:2mm;justify-content:center;align-content:start;height:287mm;overflow:hidden;break-after:page}.g:last-child{break-after:auto}
.c{box-sizing:border-box;border-radius:3mm;padding:2.2mm 2.2mm 0;display:flex;flex-direction:column}
.w{background:#fff;border-radius:1.2mm;flex:1;display:flex;flex-direction:column;align-items:center}
.no{font-size:21pt;font-weight:900;line-height:1;margin-top:2.6mm;color:#111;letter-spacing:.02em}.no small{font-size:10pt;font-weight:700;margin-left:.8mm}
.qr{width:31mm;height:31mm;margin-top:1.8mm;image-rendering:pixelated}
.t{font-size:8.6pt;font-weight:700;margin-top:1.4mm;color:#111;white-space:nowrap}
.lg{height:17mm;display:flex;justify-content:center;align-items:center}.lg img{height:13.5mm}'''

PDFJS = '''const {chromium}=require('playwright');(async()=>{const b=await chromium.launch();const p=await b.newPage();
await p.goto('file://'+process.argv[2]);await p.pdf({path:process.argv[3],format:'A4',printBackground:true,preferCSSPageSize:true});await b.close()})();'''


def b64(data):
    return base64.b64encode(data).decode()


def build(rates, out):
    d = json.load(open(os.path.join(ROOT, 'data', 'data.json')))
    seats = sorted((n, m['rate']) for m in d['machines'] if m['rate'] in rates for n in m['nos'])
    logo = b64(open(os.path.join(ROOT, 'tools', 'qr', 'logo_white.png'), 'rb').read())
    cells = []
    for n, r in seats:
        q = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, border=1, box_size=10)
        q.add_data(URL.format(n)); q.make(fit=True)
        buf = io.BytesIO(); q.make_image().save(buf, format='PNG')
        cells.append(f'<div class="c" style="background:{COLOR[r]}"><div class="w"><div class="no">{n}<small>番台</small></div>'
                     f'<img class="qr" src="data:image/png;base64,{b64(buf.getvalue())}"><div class="t">この台の機種ガイドになります</div></div>'
                     f'<div class="lg"><img src="data:image/png;base64,{logo}"></div></div>')
    pages = ''.join('<div class="g">' + ''.join(cells[i:i + PER]) + '</div>' for i in range(0, len(cells), PER))
    tmp = tempfile.mkdtemp()
    html = os.path.join(tmp, 'qr.html'); js = os.path.join(tmp, 'pdf.js')
    open(html, 'w').write(f'<!doctype html><meta charset="utf-8"><style>{CSS}</style>{pages}')
    open(js, 'w').write(PDFJS)
    npm_root = subprocess.run(['npm', 'root', '-g'], capture_output=True, text=True).stdout.strip()
    subprocess.run(['node', js, html, os.path.abspath(out)], check=True, env={**os.environ, 'NODE_PATH': npm_root})
    return [n for n, _ in seats], tmp


def verify(pdf, seats, tmp):
    import cv2
    subprocess.run(['pdftoppm', '-r', '300', '-png', pdf, os.path.join(tmp, 'p')], check=True)
    mm = 300 / 25.4; det = cv2.QRCodeDetector(); got = []
    for f in sorted(glob.glob(os.path.join(tmp, 'p-*.png'))):
        im = cv2.imread(f)
        for r in range(4):
            for c in range(3):
                x = int((13 + c * 62) * mm); y = int((5 + r * 72) * mm)
                cell = im[y:y + int(70 * mm), x:x + int(60 * mm)]
                if cell.size == 0 or cell.mean() > 250:
                    continue
                q = cell[int(12 * mm):int(49 * mm), int(10 * mm):int(50 * mm)]; v = ''
                for s in (1, .5, .35, .75, 1.5):
                    v, _, _ = det.detectAndDecode(cv2.resize(q, None, fx=s, fy=s) if s != 1 else q)
                    if v:
                        break
                got.append(v)
    exp = [URL.format(n) for n in seats]
    bad = [(i, exp[i] if i < len(exp) else None, g) for i, g in enumerate(got) if i >= len(exp) or g != exp[i]]
    return len(got) == len(exp) and not bad, len(got), bad


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('rates', nargs='+', choices=list(COLOR))
    ap.add_argument('-o', '--out')
    a = ap.parse_args()
    out = a.out or os.path.join(os.getcwd(), '台別QR_' + '・'.join(NAME[r] for r in a.rates) + '.pdf')
    seats, tmp = build(a.rates, out)
    ok, n, bad = verify(out, seats, tmp)
    print(f'{out}\n台数 {len(seats)} / 読み取り {n} / 全件一致 {ok}')
    if not ok:
        print('不一致:', bad[:10]); sys.exit(1)


if __name__ == '__main__':
    main()
