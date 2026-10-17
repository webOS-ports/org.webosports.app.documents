#!/usr/bin/env python3
# (c) 2026 Herman van Hazendonk <github.com@herrie.org>
# SPDX-License-Identifier: GPL-3.0-only
#
# Draws the three icon designs considered for Documents, in the legacy app
# icon style (an object on the translucent app plate), into
# tools/icon/variants/. images/icon.svg is the chosen one, 2-folders.
#
#   python3 tools/icon/make-variants.py
#
# No SVG filters anywhere: Qt's renderer, which the launcher uses, has none,
# so every shadow is one gradient.
import os
import math
BLUE=("#4a8fe6","#2459b2"); GREEN=("#4cb05e","#2c7d3c"); ORANGE=("#f7a043","#d4660f")
class Svg:
    def __init__(s): s.defs=[]; s.body=[]; s.n=0
    def id(s,p): s.n+=1; return f"{p}{s.n}"
    def lin(s,stops,x1=0,y1=0,x2=0,y2=1):
        i=s.id("l"); st="".join(f'<stop offset="{o}" stop-color="{c}" stop-opacity="{a}"/>' for o,c,a in stops)
        s.defs.append(f'<linearGradient id="{i}" x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}">{st}</linearGradient>'); return f"url(#{i})"
    def rad(s,stops,cx,cy,r,fx=None,fy=None,user=True):
        i=s.id("r"); st="".join(f'<stop offset="{o}" stop-color="{c}" stop-opacity="{a}"/>' for o,c,a in stops)
        f=f' fx="{fx}" fy="{fy}"' if fx is not None else ""
        s.defs.append(f'<radialGradient id="{i}" cx="{cx}" cy="{cy}" r="{r}"{f} gradientUnits="userSpaceOnUse">{st}</radialGradient>'); return f"url(#{i})"
    def add(s,x): s.body.append(x)
    def out(s):
        body=s.body[:4]+['<g transform="translate(128 126) scale(1.14) translate(-128 -128)">']+s.body[4:]+['</g>']
        return ('<svg xmlns="http://www.w3.org/2000/svg" width="256" height="256" viewBox="0 0 256 256">\n<defs>'
                        + "".join(s.defs) + "</defs>\n" + "\n".join(body) + "\n</svg>\n")
def plate(s):
    # The legacy app plate: translucent white, brightest at the top, a faint dark rim.
    s.add(f'<circle cx="128" cy="129" r="121" fill="{s.rad([(0.93,"#000",0),(0.975,"#000",0.16),(1,"#000",0)],128,129,121)}"/>')
    s.add(f'<circle cx="128" cy="128" r="117" fill="{s.lin([(0,"#fff",0.5),(0.38,"#fff",0.2),(0.62,"#fff",0.03),(1,"#fff",0.06)])}"/>')
    s.add(f'<circle cx="128" cy="128" r="117" fill="{s.rad([(0.55,"#fff",0),(1,"#fff",0.22)],128,128,117)}"/>')
    s.add('<circle cx="128" cy="128" r="117" fill="none" stroke="#000" stroke-opacity="0.13" stroke-width="1.5"/>')
def ground_shadow(s,cx,cy,rx,ry,a=0.32):
    s.add(f'<ellipse cx="{cx}" cy="{cy}" rx="{rx}" ry="{ry}" fill="{s.rad([(0,"#000",a),(1,"#000",0)],cx,cy,rx)}" transform="translate({cx} {cy}) scale(1 {ry/rx}) translate({-cx} {-cy})"/>'.replace(f'rx="{rx}" ry="{ry}"',f'rx="{rx}" ry="{rx}"'))
def sheet_shadow(s,x,y,w,h,dx=2.5,dy=4,soft=7,a=0.28):
    # One shape, one gradient: a soft-edged rect shadow (radial falloff stretched).
    cx,cy=x+w/2+dx,y+h/2+dy
    g=s.rad([(0,"#000",a),(0.7,"#000",a*0.8),(1,"#000",0)],0,0,1)
    s.defs[-1]=s.defs[-1].replace('gradientUnits="userSpaceOnUse"','gradientUnits="objectBoundingBox"').replace('cx="0" cy="0" r="1"','cx="0.5" cy="0.5" r="0.62"')
    s.add(f'<rect x="{cx-w/2-soft}" y="{cy-h/2-soft}" width="{w+2*soft}" height="{h+2*soft}" rx="{soft*2}" fill="{g}"/>')
def paper(s,x,y,w,h,curl=0):
    fill=s.lin([(0,"#ffffff",1),(1,"#e9ecef",1)])
    if curl:
        c=curl
        s.add(f'<path d="M{x+3},{y} H{x+w-3} Q{x+w},{y} {x+w},{y+3} V{y+h-c} L{x+w-c},{y+h} H{x+3} Q{x},{y+h} {x},{y+h-3} V{y+3} Q{x},{y} {x+3},{y} Z" fill="{fill}" stroke="#aeb4bc" stroke-width="1.2"/>')
    else:
        s.add(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="3" fill="{fill}" stroke="#aeb4bc" stroke-width="1.2"/>')
def curl(s,x,y,w,h,c):
    # The turned-up corner, as on the Notes and Tasks icons.
    X,Y=x+w,y+h
    s.add(f'<path d="M{X-c},{Y} Q{X-c*0.55},{Y-c*0.35} {X-c*0.62},{Y-c*0.95} Q{X-c*0.1},{Y-c*0.62} {X},{Y-c} Z" fill="{s.lin([(0,"#ffffff",1),(0.55,"#e4e7eb",1),(1,"#b9bfc7",1)],1,0,0,1)}" stroke="#a3aab3" stroke-width="1"/>')
def text_lines(s,x,y,w,rows,gap=10.5,col="#99a1ab",h=4.2,head=None):
    yy=y
    if head:
        s.add(f'<rect x="{x}" y="{yy}" width="{w*0.6}" height="7" rx="3.5" fill="{head}"/>'); yy+=15
    for i in range(rows):
        ww=w*(0.68 if i==rows-1 else (0.92 if i%3==1 else 1))
        s.add(f'<rect x="{x}" y="{yy}" width="{ww:.1f}" height="{h}" rx="{h/2}" fill="{col}"/>'); yy+=gap
def sheet_grid(s,x,y,w,h,col=GREEN):
    s.add(f'<rect x="{x}" y="{y}" width="{w}" height="12" rx="2" fill="{s.lin([(0,col[0],1),(1,col[1],1)])}"/>')
    rows=int((h-12)/13); cw=[0.34,0.33,0.33]
    s.add(f'<rect x="{x}" y="{y+12}" width="{w*cw[0]:.1f}" height="{rows*13}" fill="#e2f1e4"/>')
    for r in range(rows+1):
        s.add(f'<line x1="{x}" y1="{y+12+r*13}" x2="{x+w}" y2="{y+12+r*13}" stroke="#b9d5bd" stroke-width="1"/>')
    xx=x
    for c in cw[:-1]:
        xx+=w*c; s.add(f'<line x1="{xx:.1f}" y1="{y}" x2="{xx:.1f}" y2="{y+12+rows*13}" stroke="#b9d5bd" stroke-width="1"/>')
    for r in range(rows):
        for ci,c in enumerate(cw[1:]):
            x0=x+w*(cw[0]+sum(cw[1:1+ci])); bw=w*c*(0.35+0.4*((r*7+ci*3)%5)/4)
            s.add(f'<rect x="{x0+w*c-bw-3:.1f}" y="{y+12+r*13+4.5}" width="{bw:.1f}" height="4" rx="2" fill="#7f8a96"/>')
def slide(s,x,y,w,h,col=ORANGE):
    s.add(f'<rect x="{x}" y="{y}" width="{w}" height="15" rx="2" fill="{s.lin([(0,col[0],1),(1,col[1],1)])}"/>')
    s.add(f'<rect x="{x+6}" y="{y+5}" width="{w*0.5}" height="5" rx="2.5" fill="#fff"/>')
    r=min(w,h)*0.27; cx=x+w*0.32; cy=y+15+(h-15)/2
    a0=math.radians(-90); a1=math.radians(-90+130)
    s.add(f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="#f8cf9f"/>')
    s.add(f'<path d="M{cx},{cy} L{cx+r*math.cos(a0):.1f},{cy+r*math.sin(a0):.1f} A{r},{r} 0 0 1 {cx+r*math.cos(a1):.1f},{cy+r*math.sin(a1):.1f} Z" fill="{s.lin([(0,col[0],1),(1,col[1],1)])}"/>')
    for i in range(3):
        s.add(f'<rect x="{x+w*0.62}" y="{cy-r*0.7+i*r*0.7:.1f}" width="{w*0.3:.1f}" height="4" rx="2" fill="#99a1ab"/>')

def concept_papers():
    s=Svg(); plate(s); ground_shadow(s,128,222,92,14)
    W,H=112,148; X,Y=128-W/2,40
    for rot,kind in ((-15,"grid"),(15,"slide")):
        s.add(f'<g transform="rotate({rot} 128 232)">'); sheet_shadow(s,X,Y,W,H,a=0.18); paper(s,X,Y,W,H)
        if kind=="grid": sheet_grid(s,X+10,Y+12,W-20,H-24)
        else: slide(s,X+10,Y+12,W-20,H*0.62)
        s.add('</g>')
    y2=Y+10; s.add('<g>'); sheet_shadow(s,X,y2,W,H,a=0.3); paper(s,X,y2,W,H,curl=26)
    text_lines(s,X+12,y2+14,W-24,9,head=BLUE[1]); curl(s,X,y2,W,H,26); s.add('</g>')
    return s.out()

def folder(s,x,y,w,h,col,inner):
    # A pocket folder: back panel with a tab, a sheet in it, the front pocket.
    c0,c1=col
    s.add(f'<path d="M{x+4},{y+10} H{x+30} L{x+36},{y} H{x+62} Q{x+66},{y} {x+67},{y+4} L{x+70},{y+10} H{x+w-4} Q{x+w},{y+10} {x+w},{y+14} V{y+h-4} Q{x+w},{y+h} {x+w-4},{y+h} H{x+4} Q{x},{y+h} {x},{y+h-4} V{y+14} Q{x},{y+10} {x+4},{y+10} Z" fill="{s.lin([(0,c1,1),(1,c1,1)])}"/>')
    paper(s,x+8,y+16,w-16,h-30); inner(s,x+16,y+24,w-32,40)
    fy=y+h*0.42
    s.add(f'<path d="M{x},{fy+6} Q{x},{fy} {x+6},{fy} H{x+w-6} Q{x+w},{fy} {x+w},{fy+6} V{y+h-4} Q{x+w},{y+h} {x+w-4},{y+h} H{x+4} Q{x},{y+h} {x},{y+h-4} Z" fill="{s.lin([(0,c0,1),(1,c1,1)])}" stroke="{c1}" stroke-width="1"/>')
    s.add(f'<path d="M{x+6},{fy+1.5} H{x+w-6}" stroke="#fff" stroke-opacity="0.45" stroke-width="1.5"/>')
def concept_folders():
    s=Svg(); plate(s); ground_shadow(s,128,224,96,14)
    W,H=118,150; X,Y=128-W/2,48
    for rot,col,inner in ((-17,GREEN,lambda s,x,y,w,h: sheet_grid(s,x,y,w,h)),(17,ORANGE,lambda s,x,y,w,h: slide(s,x,y,w,h))):
        s.add(f'<g transform="rotate({rot} 128 236)">'); sheet_shadow(s,X,Y,W,H,a=0.16); folder(s,X,Y,W,H,col,inner); s.add('</g>')
    sheet_shadow(s,X,Y+8,W,H,a=0.28); folder(s,X,Y+8,W,H,BLUE,lambda s,x,y,w,h: text_lines(s,x,y,w,3,head=BLUE[1]))
    return s.out()

def concept_tabs():
    s=Svg(); plate(s); ground_shadow(s,124,224,90,13)
    W,H=128,164; X,Y=124-W/2,38
    s.add(f'<g transform="rotate(-5 128 128)">')
    for i,col in enumerate((BLUE,GREEN,ORANGE)):
        ty=Y+22+i*40
        s.add(f'<rect x="{X+W-14}" y="{ty}" width="40" height="32" rx="6" fill="{s.lin([(0,col[0],1),(1,col[1],1)],0,0,1,0)}" stroke="{col[1]}" stroke-width="1"/>')
        gx,gy=X+W+8,ty+9
        if i==0:
            for k in range(3): s.add(f'<rect x="{gx}" y="{gy+k*5.5}" width="{12 if k<2 else 8}" height="3" rx="1.5" fill="#fff"/>')
        elif i==1:
            s.add(f'<rect x="{gx}" y="{gy-1}" width="13" height="14" fill="none" stroke="#fff" stroke-width="2"/><line x1="{gx}" y1="{gy+6}" x2="{gx+13}" y2="{gy+6}" stroke="#fff" stroke-width="2"/><line x1="{gx+6.5}" y1="{gy-1}" x2="{gx+6.5}" y2="{gy+13}" stroke="#fff" stroke-width="2"/>')
        else:
            s.add(f'<circle cx="{gx+6.5}" cy="{gy+6}" r="7" fill="#fff" fill-opacity="0.5"/><path d="M{gx+6.5},{gy+6} V{gy-1} A7,7 0 0 1 {gx+13.2},{gy+8} Z" fill="#fff"/>')
    sheet_shadow(s,X,Y,W,H,a=0.3); paper(s,X,Y,W,H,curl=28)
    text_lines(s,X+14,Y+16,W-28,5,head=BLUE[1])
    sheet_grid(s,X+14,Y+98,W-40,44)
    curl(s,X,Y,W,H,28); s.add('</g>')
    return s.out()

DESCRIPTIONS = {
    "1-papers": "three sheets fanned out: a spreadsheet, a slide and, in front, a text page with its corner turned up",
    "2-folders": "blue, green and orange pocket folders, each holding a sheet of its kind (text, a spreadsheet, a slide)",
    "3-tabs": "a text page with a table, tilted, with blue, green and orange index tabs",
}

if __name__ == "__main__":
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "variants")
    os.makedirs(out, exist_ok=True)
    for n, f in (("1-papers", concept_papers), ("2-folders", concept_folders), ("3-tabs", concept_tabs)):
        header = ("<!-- Documents icon design " + n + ": " + DESCRIPTIONS[n] + ", on the app plate.\n"
                  "     Copyright 2026 Herman van Hazendonk <github.com@herrie.org>\n"
                  "     SPDX-License-Identifier: GPL-3.0-only\n"
                  "     Drawn by tools/icon/make-variants.py. -->\n")
        with open(os.path.join(out, n + ".svg"), "w") as fh:
            fh.write(header + f())
