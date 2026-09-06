# -*- coding: utf-8 -*-
"""NBC本戦プレゼン「一杯の設計図」— 予備の PowerPoint 版を生成する。

  使い方:  cd nbc-deck && python3 tools/build_pptx.py
  必要なもの: python-pptx（pip install python-pptx）

NBC_Presentation.html を読んで発表者ノートを取り込み、
同じ 1920x1080 の座標系でスライドを描き直して NBC_Presentation.pptx を書き出す。
HTML 版とページ番号がずれるのは1か所だけ：
9ページ（成分の溶け出し方）は HTML では → を3回押して1枚の中で3段階に伸ばすが、
PowerPoint にはその仕掛けを持ち込めないため、各段階を 9・10・11 の3枚に分けている。
そのため全25ページになるが、押す回数（18回）ときっかけは HTML 版と同じ。
"""
import sys, os, re, html
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deckkit import *

HTML = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'NBC_Presentation.html')

def load_notes():
    """HTML の .notes ブロックを取り出して、発表者ノート用のプレーンテキストにする。"""
    s = io.open(HTML, encoding='utf-8').read()
    out = []
    for m in re.finditer(r'<section class="s[^"]*" data-t="([^"]*)".*?</section>', s, re.S):
        n = re.search(r'<div class="notes">(.*?)</div>\s*</section>', m.group(0), re.S)
        t = ''
        if n:
            t = n.group(1)
            t = re.sub(r'<h4>(.*?)</h4>', r'【\1】\n', t, flags=re.S)
            t = re.sub(r'<hr[^>]*>', '\n────────\n', t)
            t = html.unescape(re.sub(r'<[^>]+>', '', t.replace('<br>', '\n')))
            t = re.sub(r'\n{3,}', '\n\n', '\n'.join(l.strip() for l in t.split('\n'))).strip()
        out.append(t)
    if len(out) != 23:
        raise SystemExit('HTML のセクション数が %d です（23 を想定）。tools/build_pptx.py の対応表を見直してください。' % len(out))
    return out

import io
NOTES = load_notes()
def nt(i): return NOTES[i-1]        # HTML 側の 1 始まりページ番号


def s07(prs, nt):
    s = slide(prs)
    head(s, "THE VARIABLES", "淹れ手が動かせるのは、この4つ")
    txt(s, L,891,CW,93, "茶葉にない味は作れない。[[表れ方を変えられる。]]", 64, INK, True,'c',1.45,SANS)
    cards = [("湯量","200((mL))","1椀40mL × 5碗","今回は固定",0),
             ("茶葉量","7((g))","湯量に対して決まる","今回は固定",0),
             ("温度","","2種類の湯を使い分ける","今回動かす",1),
             ("時間","60 / 10((秒))","どの湯に、何秒浸すか","今回動かす",1)]
    for i,(k,v,sub,fix,act) in enumerate(cards):
        x = 104 + i*435; y = 331; w = 405; h = 543
        rrect(s, x,y,w,h, r=10, fill=WHITE, line=GOLD if act else INK3, lw=6 if act else 4)
        txt(s, x,y+24,w,57, k, 38, GOLD if act else INK2, True,'c',1.0,SANS,38*0.06)
        fx, fy = x+102, y+90
        if i==0:      # 湯量：湯呑みに200mLまで注がれている
            rrect(s, fx+50,fy+15,100,220, r=14, line=INK, lw=5)
            rect(s, fx+54,fy+100,92,131, fill=INK3, alpha=46)
            rect(s, fx+54,fy+97,92,6, fill=GOLD)
            for ty in (150,185,220): rect(s, fx+120,fy+ty,26,4, fill=INK3)
        elif i==1:    # 茶葉：茶則に7gの茶葉
            pie(s, fx+25,fy+70,150,150, 0,180, fill=LEAF, line=INK, lw=5)
            rect(s, fx+8,fy+142,184,5, fill=INK)
        elif i==2:    # 温度：50℃まで青、その上に90℃まで赤
            rrect(s, fx+84,fy+15,32,175, r=16, fill=WHITE, line=INK, lw=5)
            rect(s, fx+90,fy+110,20,78, fill=COLD)
            rect(s, fx+90,fy+42,20,68, fill=HOT)
            oval(s, fx+70,fy+163,60,60, fill=COLD, line=INK, lw=5)
            rect(s, fx+122,fy+108,26,5, fill=COLD)
            rect(s, fx+122,fy+40,26,5, fill=HOT)
        else:         # 時間：外周が60秒、内周がその1/6＝10秒
            donut(s, fx+25,fy+35,150,150, th=0.30, fill=COLD)
            donut(s, fx+52,fy+62,96,96, th=0.35, fill=INK3, alpha=22)
            pie(s, fx+52,fy+62,96,96, 270,330, fill=HOT)
            oval(s, fx+69,fy+79,62,62, fill=WHITE)
        if i==2:
            txt(s, x,y+349,w,74, "50 / 90((℃))", 70, INK, True,'c',1.05,SANS, wrap=False)
        elif v:
            txt(s, x,y+349,w,74, v, 70, INK, True,'c',1.05,SANS, wrap=False)
        txt(s, x,y+431,w,42, sub, 32, INK2, True,'c',1.3,SANS)
        txt(s, x,y+481,w,38, fix, 30, GOLD if act else INK3, True,'c',1.0,SANS,30*0.06)
    notes(s, nt(7)); return s

def s08(prs, nt):
    s = slide(prs)
    head(s, "THE FIRST BREW", "一煎目", mb=30)
    txt(s, L,306,CW,90, "<<総湯量>>　{{200mL}}　<<／>>　<<茶葉>>　{{7g}}", 80, INK, True,'c',1.1,SANS, small=44)
    # PHASE 1 / PHASE 2
    for x, bd, bg, tag, big, n2, why in (
        (104, COLD, COLDS, "PHASE 1 ／ 低温", "50((℃))", "30((mL))　　60((秒))",
         "少量の湯で茶葉を\nひたひたに濡らし、時間をかける"),
        (1190, HOT, HOTS, "PHASE 2 ／ 追い湯", "90((℃))", "+170((mL))　　10((秒))",
         "温度を上げて、\n短時間で注ぎ分ける")):
        rect(s, x,462,626,485, fill=bg, line=bd, lw=6)
        txt(s, x,492,626,60, tag, 40, bd, True,'c',1.0,SANS,40*0.10)
        txt(s, x,566,626,132, big, 132, INK, True,'c',1.0,SANS, wrap=False)
        txt(s, x,714,626,84, n2, 84, INK, True,'c',1.0,SANS, wrap=False)
        txt(s, x,816,626,101, why, 36, INK2, True,'c',1.4,SANS)
    # 煎茶碗（一段目・二段目が重なった一杯）
    b = pie(s, 810,472,300,290, 0,180, line=INK, lw=5); grad(b, TEAG, YAMA, 270)
    rect(s, 800,614,320,5, fill=INK)
    rect(s, 878,758,8,24, fill=INK); rect(s, 1034,758,8,24, fill=INK)
    rect(s, 870,780,180,6, fill=INK)
    txt(s, 760,800,400,46, "一煎目の一杯", 33, INK2, True,'c',1.2,SANS)
    notes(s, nt(8)); return s

ROWS = [("テアニン","アミノ酸類／旨味・甘味", 40.5, 49.3, None),
        ("EGC","遊離型カテキン／爽やかな渋み", 17.5, 33.2, 44.2),
        ("EGCg","ガレート型／重い苦渋味", 6.9, 17.1, 31.5),
        ("カフェイン","アルカロイド／苦味", 21.6, 34.6, 62.7)]
CAPS = ["旨味・甘みの成分は、[[低い温度でも溶け出しやすい。]]",
        "テアニンも増える。だが[[苦味・渋みの成分の増え方が上回る。]]",
        "高温では、[[苦味も渋みも短時間で出る。]]味わいのバランスが変わる。"]

def s09(prs, nt, state):
    s = slide(prs)
    head(s, "THE SCIENCE", "成分は、同じようには溶け出さない", tsize=80, mb=10)
    txt(s, L,276,CW,48, "同じ茶葉を {{50℃}} ／ {{70℃}} ／ {{90℃}} で、それぞれ 60秒 淹れたとき",
        33, INK2, True,'c',1.2,SANS, gold=INK2)
    for bi,(name,sub,v1,v2,v3) in enumerate(ROWS):
        by = 324 + bi*128
        txt(s, 104,by,440,112, "%s\n<<%s>>"%(name,sub), 48, INK, True,'r',1.2,SANS, small=28)
        for j,(tt,col,val) in enumerate(((("50℃"),COLD,v1),("70℃",WARM,v2),("90℃",HOT,v3))):
            ry = by + j*40
            on = j <= state
            fade = None if on else 16
            txt(s, 574,ry-6,126,44, tt, 34, col, True,'l',1.0,SANS, anchor='m', wrap=False)
            rect(s, 720,ry+2,906,28, fill=INK3, alpha=(16 if on else 8))
            if val is None:
                txt(s, 734,ry-6,300,44, "90℃条件の測定なし", 26, INK2, True,'l',1.0,SANS, anchor='m')
                txt(s, 1646,ry-6,170,44, "—", 30, INK3, True,'l',1.0,SANS, anchor='m', wrap=False)
            else:
                if on: rect(s, 720,ry+2,int(906*val/100.0),28, fill=col)
                txt(s, 1646,ry-8,170,48, "%.1f%%"%val, 40, col if on else INK3, True,'l',1.0,SANS, anchor='m', wrap=False)
            if not on:
                rect(s, 574,ry-8,1242,48, fill=PAPER, alpha=78)
    txt(s, L,828,CW,60, CAPS[state], 46, INK, True,'c',1.3,SANS)
    txt(s, L,896,CW,88,
        "坂本ら（2002）茶業研究報告 94:45-55、表2・表3（p.49）／深蒸し上級煎茶4g・湯120mL・1煎目60秒。"
        "3つの温度はそれぞれ別々に淹れた1煎目で、低温のあとに高温を重ねた結果ではない。"
        "アミノ酸類（表3）に90℃条件はない。拓朗（午）の実測値でもない。",
        26, INK2, False,'l',1.5,SANS)
    notes(s, nt(9) if state==0 else "【%d段目】\n→ を押すと %s のバーが伸びます。\n\n%s"%(
        state, ("70℃","90℃")[state-1], CAPS[state].replace('[[','').replace(']]','')))
    return s

def s12(prs, nt):
    s = slide(prs)
    head(s, "THE FIRST BREW ／ DESIGN", "だから、一煎目は二段で淹れる", mb=12)
    txt(s, L,288,CW,54, "合組された二つの茶の良さを際立たせるために、成分が{{溶け出す性質の差}}を使う",
        34, INK2, True,'c',1.2,SANS)
    k = 0.8419; ox, oy = 286.5, 352
    def fx(v): return ox + v*k
    def fy(v): return oy + v*k
    def tl(sx, sy, w, s_, size, col, al='c'):
        txt(s, fx(sx-w/2.0), fy(sy)-size*k*0.85, w*k, size*k*1.3, s_, size*k, col, True, al,1.0,SANS)
    line(s, fx(110),fy(380), fx(1520),fy(380), INK3, 3)
    line(s, fx(110),fy(300), fx(1520),fy(300), COLD, 3, dash='dash')
    line(s, fx(110),fy(96),  fx(1520),fy(96),  HOT,  3, dash='dash')
    txt(s, fx(110)-260, fy(300)-30, 250, 60, "50℃", 40*k, COLD, True,'r',1.0,SANS)
    txt(s, fx(110)-260, fy(96)-30,  250, 60, "約84℃", 40*k, HOT, True,'r',1.0,SANS)
    line(s, fx(150),fy(300), fx(1180),fy(300), COLD, 16)
    line(s, fx(1180),fy(300), fx(1180),fy(96), HOT, 16)
    line(s, fx(1180),fy(96), fx(1450),fy(96), HOT, 16)
    line(s, fx(150),fy(346), fx(1180),fy(346), COLD, 4)
    line(s, fx(150),fy(336), fx(150),fy(356), COLD, 4)
    line(s, fx(1180),fy(336), fx(1180),fy(356), COLD, 4)
    tl(665,400,500,"60秒",44,COLD); tl(665,256,600,"50℃ ・ 30mL",44,INK)
    line(s, fx(1180),fy(142), fx(1450),fy(142), HOT, 4)
    line(s, fx(1180),fy(132), fx(1180),fy(152), HOT, 4)
    line(s, fx(1450),fy(132), fx(1450),fy(152), HOT, 4)
    tl(1315,196,400,"10秒",44,HOT); tl(1315,60,600,"＋90℃ ・ 170mL",44,INK)
    for x, bd, bg, t_, lis in (
        (104, COLD, COLDS, "PHASE 1 ／ 低温・時間をかける",
         ["アミノ酸類を先に引き出す","カテキン類はまだ出にくい"]),
        (975, HOT, HOTS, "PHASE 2 ／ 追い湯・高温・短時間",
         ["90℃の湯を加え、急須の中で約84℃に","爽やかな渋みを後から立ち上げる","10秒で切り、苦味・重い苦渋味の出過ぎを避ける"])):
        rect(s, x,724,841,258, fill=bg, line=bd, lw=5)
        txt(s, x+26,748,789,54, t_, 36, bd, True,'l',1.0,SANS,36*0.04)
        txt(s, x+26,812,789,150, "\n".join("・"+v for v in lis), 32, INK, True,'l',1.45,SANS)
    notes(s, nt(10)); return s


def warn(s, y, h, text, size=36):
    rect(s, L,y,CW,h, fill=HOT, alpha=7)
    rect(s, L,y,10,h, fill=HOT)
    txt(s, L+30,y+18,CW-60,h-36, text, size, INK, False,'l',1.45,SANS)

def s13(prs, nt):
    s = slide(prs)
    head(s, "THE SENSORY", "一杯の中に、二つの表情", mb=26)
    for x, col, body in ((104, TEAG, "若草のような爽やかな香り\n煎茶らしいキレのある渋み"),
                         (1016, WARM, "瓜のような\n瑞々しい甘みと香り")):
        rrect(s, x,302,800,200, r=12, fill=WHITE, line=col, lw=5)
        txt(s, x+24,302,752,200, body, 46, INK, True,'c',1.45,SANS)
    line(s, 504,512, 782,584, INK3, 6, dash='sysDot')
    line(s, 1416,512, 1138,584, INK3, 6, dash='sysDot')
    b = pie(s, 810,445,300,290, 0,180, line=INK, lw=5); grad(b, TEAG, YAMA, 270)
    rect(s, 800,587,320,5, fill=INK)
    rect(s, 878,731,8,24, fill=INK); rect(s, 1034,731,8,24, fill=INK)
    rect(s, 870,753,180,6, fill=INK)
    txt(s, L,855,CW,54, "茶師に伺った設計は、川根本町の{{「山の香気」}}に、鹿児島県産の{{「穏やかな甘み」}}を合わせたもの。",
        36, INK2, True,'c',1.5,SANS)
    txt(s, L,917,CW,42, "※ 図の二つは私自身の官能表現です。どちらの香味がどの産地に由来するか、という対応づけではありません。",
        28, INK2, False,'c',1.5,SANS)
    notes(s, nt(11)); return s

def s14(prs, nt):
    s = slide(prs)
    head(s, "THE SECOND BREW", "二煎目", mb=30)
    txt(s, L,306,CW,72, "茶葉は同じ。条件だけを変える。", 48, INK, True,'c',1.2,SANS)
    rect(s, 104,437,1712,523, fill=HOTS, line=HOT, lw=6)
    txt(s, 104,477,1712,69, "高温・短時間", 46, HOT, True,'c',1.0,SANS,46*0.10)
    txt(s, 104,576,1712,172, "90((℃))　　200((mL))　　10((秒))", 172, INK, True,'c',1.0,SANS, wrap=False)
    txt(s, 204,788,1512,62, "一煎目で十分に吸水した茶葉は、一煎目よりも成分が溶出しやすい", 44, INK2, True,'c',1.3,SANS)
    notes(s, nt(12)); return s

def s15(prs, nt):
    s = slide(prs)
    y = 307
    txt(s, L,y,CW,48, "THE SECOND BREW", 38, GOLD, True,'c',1.0,SANS,38*0.22); y+=48+22
    txt(s, L,y,CW,150, "爽快に、[[さっぱりと。]]", 120, INK, True,'c',1.25,SANS);  y+=150+44
    txt(s, L,y,CW,180, "一煎目で残ったものを、\n高温・短時間で引き出す。", 60, INK2, True,'c',1.5,SANS)
    notes(s, nt(13)); return s

def s16(prs, nt):
    s = slide(prs)
    head(s, "SYNTHESIS", "同じ茶葉、異なる表情")
    txt(s, L,891,CW,93, "どちらが好きかは、[[人によって違っていい。]]", 64, INK, True,'c',1.45,SANS)
    cx = [104, 384, 1100]; cw = [280, 716, 716]
    rect(s, L,336,CW,4, fill=INK)
    rect(s, 384,340,1432,90, fill=GOLD, alpha=10)
    for x,w,t_ in ((384,716,"一煎目"),(1100,716,"二煎目")):
        txt(s, x,340,w,90, t_, 54, INK, True,'c',1.0,SANS)
    rowsy = [(430,150),(580,120),(700,130)]
    heads = ["条件","狙い","味わい"]
    cells = [("{{50℃}} 30mL 60秒\n→ {{90℃}} 170mL 10秒", "{{90℃}} 200mL 10秒"),
             ("低温で時間を取り、\nあとから温度を上げる", "吸水済みの茶葉に\n高温を短く当てる"),
             ("複層的\n<<爽やかな香りと渋み＋瑞々しい甘み>>", "爽快・さっぱり\n<<そのお茶らしいキレ>>")]
    for (y,h), hd, (a,b) in zip(rowsy, heads, cells):
        txt(s, 128,y,232,h, hd, 40, INK2, True,'l',1.3,SANS)
        txt(s, 384,y,716,h, a, 46, INK, True,'c',1.3,SANS, gold=INK, small=38)
        txt(s, 1100,y,716,h, b, 46, INK, True,'c',1.3,SANS, gold=INK, small=38)
        rect(s, L,y+h,CW,3, fill=INK3, alpha=55)
    notes(s, nt(14)); return s

def s17(prs, nt):
    s = slide(prs)
    head(s, None, "これは「正解の淹れ方」ではなく、\n[[一つの仮説]]です。", center=True)
    x = 264
    for jp, en, hi in (("知る","LEARN",0),("試す","TRY",0),("味わう","TASTE",0),("調整","ADJUST",1)):
        oval(s, x,415,246,246, fill=GOLD if hi else INK)
        txt(s, x,470,246,94, jp, 62, PAPER, True,'c',1.0,SANS)
        txt(s, x,572,246,45, en, 30, PAPER, True,'c',1.0,SANS,30*0.10)
        x += 246
        if not hi:
            txt(s, x,490,136,96, "→", 96, GOLD, False,'c',1.0,SANS); x += 136
    txt(s, 104,668,52,63, "◀", 52, GOLD, False,'c',1.0,SANS)
    rect(s, 150,694,290,6, fill=GOLD)
    txt(s, 440,668,540,63, "また、次の一杯へ", 42, GOLD, True,'c',1.0,SANS)
    rect(s, 980,694,836,6, fill=GOLD)
    txt(s, L,824,CW,160, "知識は正解ではなく、[[自分で考えるための材料]]。\n<<積み重ねられてきた知見を土台に、そこから一歩だけ踏み出してみた淹れ方です。>>",
        64, INK, True,'c',1.45,SANS, small=44)
    notes(s, nt(15)); return s

def s18(prs, nt):
    s = slide(prs)
    head(s, "EPILOGUE ／ 一人の飲み手として", "日本茶は、何も知らなくても美味しく飲めます。", tsize=68)
    txt(s, L,318,CW,78, "その上で、", 52, INK2, True,'c',1.5,SANS)
    for x, body in ((104,"「このお茶は、\nどんなお茶だろう」"), (980,"「どうやって\n淹れてみよう」")):
        rect(s, x,432,836,220, fill=WHITE, line=INK3, lw=5)
        txt(s, x+28,432,780,220, body, 54, INK, True,'c',1.35,SANS)
    txt(s, L,688,CW,186, "そう思ったなら、その時点で、\n[[その一杯を自分で考え始めている。]]", 64, INK, True,'c',1.45,SANS)
    txt(s, L,903,CW,81, "私が伝えたいのは、まさにその[[楽しさ]]です。", 56, INK, True,'c',1.45,SANS)
    notes(s, nt(16)); return s

def s19(prs, nt):
    s = slide(prs, double=True)
    txt(s, L,274,CW,156, "10年、20年経てば、また次の世代が日本茶に出会い、\n何度でも再発見される。", 52, INK2, True,'c',1.5,SANS)
    txt(s, L,482,CW,302, "飲む人が、\n[[日本茶を面白がり続けること。]]", 108, INK, True,'c',1.4,MIN)
    notes(s, nt(17)); return s

def s20(prs, nt):
    s = slide(prs)
    y = 308
    txt(s, L,y,CW,48, "APPENDIX", 38, GOLD, True,'c',1.0,SANS,38*0.22); y+=48+22
    txt(s, L,y,CW,168, "質疑応答用", 150, INK, True,'c',1.12,MIN); y+=168+40
    txt(s, L,y,CW,163, "なぜ50℃・60秒か　／　なぜ30mLか\nなぜ90℃・10秒か　／　産地を淹れ分けているのか", 48, INK2, True,'c',1.7,SANS)
    notes(s, nt(18)); return s

def s21(prs, nt):
    s = slide(prs)
    head(s, "Q&A 01", "なぜ 50℃・60秒 なのか。", tsize=74, mb=20)
    txt(s, L,279,CW,53, "同じ茶葉を 60秒 淹れたとき、旨味と重い苦渋味の出方に、どれだけ差がつくか", 33, INK2, True,'c',1.2,SANS)
    for i,(k,v,sub,hi) in enumerate((
        ("40℃以下","時間切れ","低いほど差は開くが、\n必要な浸出時間が延び、\n10分のプレゼンに収まらない",0),
        ("50℃（採用）","差が残る","旨味は出ているのに、\n重い苦渋味は\nまだ大きく出ていない",1),
        ("70℃","差が縮む","旨味も増えるが、\n重い苦渋味の\n増え方が上回る",0))):
        x = 104 + i*582
        rrect(s, x,350,548,390, r=10, fill=WHITE if hi else None, line=GOLD if hi else INK3, lw=6 if hi else 4)
        txt(s, x,392,548,57, k, 38, GOLD if hi else INK2, True,'c',1.0,SANS,38*0.06)
        txt(s, x,459,548,79, v, 72, GOLD if hi else INK2, True,'c',1.1,SANS)
        txt(s, x+26,560,496,150, sub, 33, INK2, True,'c',1.4,SANS)
    txt(s, L,748,CW,88, "坂本ら（2002）表2・表3（p.49）／深蒸し上級煎茶4g・湯120mL・1煎目60秒。50℃：テアニン40.5%／EGCg 6.9%　→　70℃：テアニン49.3%／EGCg 17.1%。40℃条件は論文になく、傾向からの判断。",
        29, INK2, False,'c',1.5,SANS)
    warn(s, 844, 140, "言い方の境界：「50℃なら旨味だけが出る」とは言わない。低温でも溶出しやすい成分と、低温では溶出が遅い成分の差を利用している、と説明する。")
    notes(s, nt(19)); return s

def s22(prs, nt):
    s = slide(prs)
    head(s, "Q&A 02", "なぜ最初の湯は 30mL なのか。", tsize=74)
    for i,(k,v,sub,hi) in enumerate((
        ("20mL","約86℃","7gの茶葉を\n均一に濡らせない恐れ",0),
        ("30mL（採用）","約84℃","ひたひたに濡らす実用性と、\n追い湯後の温度の立ち上がりを両立",1),
        ("40mL","約82℃","低温相の液量が増え、\n混合温度が下がる",0))):
        x = 104 + i*582
        rrect(s, x,370,548,390, r=10, fill=WHITE if hi else None, line=GOLD if hi else INK3, lw=6 if hi else 4)
        txt(s, x,412,548,57, k, 38, GOLD if hi else INK2, True,'c',1.0,SANS,38*0.06)
        txt(s, x,479,548,97, v, 88, GOLD if hi else INK2, True,'c',1.1,SANS)
        txt(s, x+16,588,516,110, sub, 30, INK2, True,'c',1.4,SANS)
    txt(s, L,780,CW,45, "※ 90℃の湯を加えたときの、熱損失を無視した単純混合温度。", 30, INK2, False,'c',1.5,SANS)
    warn(s, 844, 140, "言い方の境界：「30mLは科学的に導出された最適湯量」とは言わない。原理は研究から、具体値は実飲から決めた、と説明する。")
    notes(s, nt(20)); return s

def s23(prs, nt):
    s = slide(prs)
    head(s, "Q&A 03", "なぜ二煎目は 90℃・10秒 なのか。", tsize=66, mb=20)
    txt(s, L,269,CW,58, "吸水済みの茶葉なら、90℃・10秒 でもこれだけ追加で出る（同じ茶葉での連続抽出）", 36, INK2, True,'c',1.2,SANS)
    rect(s, 560,360,34,22, fill=INK);  txt(s, 606,343,420,52, "1煎目 90℃・60秒", 34, INK, True,'l',1.0,SANS)
    rect(s, 1000,360,34,22, fill=HOT); txt(s, 1046,343,520,52, "2煎目 90℃・10秒の追加分", 34, INK, True,'l',1.0,SANS)
    for i,(name,sub,g1,g2,tot) in enumerate((("EGC","遊離型／爽やかな渋み",44.2,35.0,79.2),
                                             ("EGCg","ガレート型／重い苦渋味",31.5,15.4,46.9),
                                             ("カフェイン","苦味",62.7,20.1,82.8))):
        y = 419 + i*84
        txt(s, 104,y-12,430,84, "%s\n<<%s>>"%(name,sub), 44, INK, True,'r',1.2,SANS, small=26)
        rect(s, 564,y,846,60, fill=INK3, alpha=16)
        rect(s, 564,y,int(846*g1/100.0),60, fill=INK)
        rect(s, 564+int(846*g1/100.0),y,int(846*g2/100.0),60, fill=HOT)
        txt(s, 1430,y,386,60, "%.1f%% {{+%.1f}} → %.1f%%"%(g1,g2,tot), 34, INK, True,'l',1.0,SANS, gold=HOT, wrap=False)
    txt(s, L,690,CW,132, "坂本ら（2002）表2（p.49）。EC は 44.6%→79.2%、ECg は 27.6%→41.5%。10秒でも十分に出る。だから長く置かないという設計。"
        "ただし研究の1煎目は90℃・60秒であり、今回の一煎目（50℃30mL60秒＋約84℃10秒）とは茶葉の状態が異なるため、"
        "この追加分がそのまま拓朗（午）に当てはまるわけではない。", 28, INK2, False,'l',1.5,SANS)
    warn(s, 844, 140, "言い方の境界：「90℃・10秒が拓朗（午）の科学的最適値」とは言わない。「出し切る」も言わない。既存研究の挙動と実飲目的に整合する実践的な基準値として採用した、と説明する。")
    notes(s, nt(21)); return s

def s24(prs, nt):
    s = slide(prs)
    head(s, "Q&A 04", "川根本町と鹿児島を、淹れ分けているのか。", tsize=74)
    for x, col, mark_, body, note_ in (
        (104, HOT, "✕", "産地ごとの成分を\n分けて抽出している", "そのような証拠はありません"),
        (980, INK, "〇", "茶師が一つにまとめた茶から\n異なる側面を表現している", "今回行っているのはこちら")):
        rect(s, x,352,836,400, fill=(WHITE if col is INK else HOT), line=col, lw=6,
             alpha=(None if col is INK else 6))
        txt(s, x,392,836,90, mark_, 90, col, True,'c',1.0,SANS)
        txt(s, x+30,500,776,140, body, 52, INK, True,'c',1.35,SANS)
        txt(s, x+30,662,776,52, note_, 34, INK2, True,'c',1.2,SANS)
    warn(s, 792, 192, "区別する三層：① 文献で確認された事実　② 文献を土台にした今回の設計判断　③ 拓朗（午）での実飲・官能評価。"
                      "「瓜のような瑞々しい香り」も香気分析はしておらず、あくまで「私はそう感じている」という官能表現。")
    notes(s, nt(22)); return s

def s25(prs, nt):
    s = slide(prs)
    head(s, "REFERENCES", "主要参考文献")
    refs = ["坂本 彬・中川 致之・杉山 弘成・堀江 秀樹（2002）「煎茶の1煎, 2煎, 3煎液の成分組成に基づく溶出特性」茶業研究報告 94:45-55. DOI: 10.5979/cha.2002.94_45",
            "堀江 秀樹・氏原 ともみ・木幡 勝則（2001）「茶主要成分の茶浸出液への溶出特性」茶業研究報告 91:29-33. DOI: 10.5979/cha.2001.91_29",
            "池田 重美・中川 致之・岩浅 潔（1972）「煎茶の浸出条件と可溶成分との関係」茶業研究報告 37:69-78. DOI: 10.5979/cha.1972.69",
            "久保 智子・藤原 孝之・冨澤 代志子（2014）「異なる条件で浸出した緑茶の渋味およびうま味の味覚センサーによる評価」日本食品科学工学会誌 61(5):192-198. DOI: 10.3136/nskkk.61.192",
            "野村 幸子ほか（2023）「‘やぶきた’二番茶の煎茶および釜炒り茶の冷水浸出液における化学成分の違い」茶業研究報告 135:19-28"]
    txt(s, 148,316,1668,470, "\n".join("・"+r for r in refs), 28, INK, False,'l',1.62,SANS)
    txt(s, L,800,CW,190, "いずれも拓朗（午）を試料とした研究ではありません。原理の理解に用い、具体的な条件は実飲によって決定しています。\n"
        "坂本ら（2002）は J-STAGE で全文（PDF）を確認し、本資料の数値は表2（カテキン類・カフェイン）および表3（遊離アミノ酸）p.49 から直接引用しています。",
        29, INK2, False,'l',1.5,SANS)
    notes(s, nt(23)); return s


prs = deck()

# ---- 1 表紙 ----
s = slide(prs, double=True); y = 242
txt(s, L,y,CW,48, "NIHONCHA BREWERS CHAMPIONSHIP 2026", 38, GOLD, True,'c',1.0,SANS,38*0.22); y+=48+34
txt(s, L,y,CW,206, "一杯の設計図", 184, INK, True,'c',1.12,MIN,184*0.02);                    y+=206+26
txt(s, L,y,CW,62, "THE BLUEPRINT OF A CUP", 52, GOLD, True,'c',1.0,SANS,52*0.28);            y+=62+56
rect(s, CX-140,y,280,5, fill=INK);                                                            y+=5+56
txt(s, L,y,CW,80, "しもきた茶苑大山　茶師十段之茶「拓朗（午）」", 64, INK, True,'c',1.25,SANS)
notes(s, nt(1))

# ---- 2 論語 ----
s = slide(prs); y = 201
txt(s, L,y,CW,48, "論語 ／ 雍也篇", 38, GOLD, True,'c',1.0,SANS,38*0.22);   y+=48+22
txt(s, L,y,CW,81, "子曰、知之者不如好之者、好之者不如楽之者", 54, INK2, True,'c',1.0,MIN,54*0.16); y+=81+46
gx = 472
for ch, rd, last in (("知","知る",0),("好","好む",0),("楽","楽しむ",1)):
    txt(s, gx,y,200,200, ch, 200, GOLD if last else INK, True,'c',1.0,MIN)
    txt(s, gx,y+216,200,63, rd, 42, INK2, True,'c',1.0,SANS,42*0.10)
    gx += 200
    if not last:
        txt(s, gx,y+34,188,120, "＜", 96, GOLD, False,'c',1.0,SANS); gx += 188
y += 279+40
txt(s, L,y,CW,139, "これを知る者は、これを好む者に如かず。\nこれを好む者は、これを楽しむ者に如かず。", 42, INK2, True,'c',1.65,SANS)
notes(s, nt(2))

# ---- 3 入口が広い ----
s = slide(prs); y = 119
txt(s, L,y,CW,115, "日本茶は、[[入口が広い。]]", 96, INK, True,'c',1.2,SANS);            y+=115+22
txt(s, L,y,CW,78, "ただ美味しく飲むだけでも、十分な付き合い方。", 52, INK2, True,'c',1.5,SANS); y+=78+30
txt(s, L,y,CW,84, "▼", 70, GOLD, False,'c',1.2,SANS);                                     y+=84+26
txt(s, L,y,CW,115, "その先は、[[限りない。]]", 96, INK, True,'c',1.2,SANS);               y+=115+34
cx0 = CX-668
for w in ("産地","品種","栽培","製茶","歴史","文化","科学"):
    rrect(s, cx0,y,172,114, r=57, fill=WHITE, line=INK, lw=4)
    txt(s, cx0,y,172,114, w, 52, INK, True,'c',1.0,SANS)
    cx0 += 172+22
y += 114+40
txt(s, L,y,CW,162, "知ったことを手掛かりに、考え、試しながら味わう。\nそれも、日本茶を[[楽しむ]]ということ。", 56, INK, True,'c',1.45,SANS)
notes(s, nt(3))

# ---- 4 「茶師」って、なんだろう ----
s = slide(prs)
head(s, "2019 ／ 東京・下北沢", "「茶師」って、なんだろう。")
txt(s, L,798,CW,186, "産地ではない街に、\n[[お茶を仕上げる技術]]があった。", 64, INK, True,'c',1.45,SANS)
for (num, lab, hi), x in zip((("01","荒茶を\n仕入れる",0),("02","個性を\n見極める",0),
                              ("03","火入れと\n合組",1),("04","一つの茶に\n仕上げる",0)),
                             (104, 550, 996, 1442)):
    txt(s, x,424,374,66, num, 44, GOLD, True,'c',1.0,SANS,44*0.10)
    txt(s, x,508,374,160, lab, 64, GOLD if hi else INK, True,'c',1.25,SANS)
for ax in (478, 924, 1370):
    txt(s, ax,506,72,80, "→", 72, GOLD, False,'c',1.0,SANS)
notes(s, nt(4))

# ---- 5 拓朗（午） ----
s = slide(prs)
head(s, "THE DESIGN", "茶師が、一つにまとめた茶")
txt(s, L,798,CW,186, "熱湯で淹れても、低温で淹れても、美味しい。\n[[「これは、設計された飲み物だ。」]]", 64, INK, True,'c',1.45,SANS)
oval(s, 774,370,372,372, fill=GOLD, alpha=16)
oval(s, 790,386,340,340, fill=GOLD)
txt(s, 790,436,340,154, "拓朗\n（午）", 64, PAPER, True,'c',1.2,SANS)
txt(s, 790,596,340,45, "GOGUMI", 30, PAPER, True,'c',1.0,SANS,30*0.14)
txt(s, 650,514,140,88, "→", 88, GOLD, False,'c',1.0,SANS)
txt(s, 1130,514,140,88, "←", 88, GOLD, False,'c',1.0,SANS)
for x, tag, name, note_ in ((104,"BASE","静岡県・川根本町","ベースにしている茶"),
                            (1270,"BLEND","鹿児島県産","組み合わせている茶")):
    txt(s, x,444,546,54, tag, 36, GOLD, True,'c',1.0,SANS,36*0.18)
    txt(s, x,516,546,72, name, 60, INK, True,'c',1.2,SANS)
    txt(s, x,608,546,60, note_, 40, INK2, True,'c',1.2,SANS)
notes(s, nt(5))

# ---- 6 今日のテーマ ----
s = slide(prs); y = 136
txt(s, L,y,CW,48, "TODAY'S THEME", 38, GOLD, True,'c',1.0,SANS,38*0.22);  y+=48+22
txt(s, L,y,CW,156, "茶師が考え尽くした設計があるのなら、\n奇をてらわず、普通に淹れるのが筋。", 52, INK2, True,'c',1.5,SANS); y+=156+26
txt(s, L,y,CW,93, "それでも、あえて。", 62, GOLD, True,'c',1.2,SANS);      y+=93+24
txt(s, L,y,CW,216, "二種類のお茶が[[含まれていると分かる]]、\nそんな一杯を作ってみる。", 80, INK, True,'c',1.35,SANS); y+=216+44
txt(s, L,y,CW,157, "淹れ手が狙いを持って、お茶を淹れる。\nその[[プロセス自体]]も、一緒に楽しんでいただけたら。", 54, INK, True,'c',1.45,SANS)
notes(s, nt(6))

# ---- 7〜12 ----
s07(prs, nt); s08(prs, nt)
for k in (0,1,2): s09(prs, nt, k)
s12(prs, nt)
# ---- 13〜25 ----
for fn in (s13,s14,s15,s16,s17,s18,s19,s20,s21,s22,s23,s24,s25):
    fn(prs, nt)

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'NBC_Presentation.pptx')
prs.save(OUT)
print('書き出しました: %s（%d ページ）' % (os.path.normpath(OUT), len(prs.slides._sldIdLst)))
