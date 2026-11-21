"""MOTIF 스마트스토어 상세페이지 도구 (2026-11-21).
쓰는 법:  from ss_kit import *   (같은 폴더에 두거나 sys.path 에 추가)
- GIF: theme('tint'|'ppf') 후 F · ease · txt · ctr · mix · save 로 PIL 프레임을 그린다. 폭 W=860.
- HTML: Page(theme, sha, folder, prefix) 로 블록을 쌓고 .write(path) → (바이트, 글자 수, GIF 수) 반환.
- 확인: render(html, repo_dir) · preview_mp4(gifs, out) · check_urls(html, extra) · tile_last_frames(gifs, out)
"""
from PIL import Image, ImageDraw, ImageFont
import re, subprocess, os
FB='/usr/share/fonts/opentype/noto/NotoSansCJK-Black.ttc'; FBo='/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc'; FR='/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'
def F(s,f=FBo): return ImageFont.truetype(f,s,index=1)
W=860
WH=(255,255,255); LG=(243,244,247); INK=(17,20,28); TX=(26,28,34); LINE=(222,226,234); GR=(150,155,165); HOT=(255,120,90)
# 테마 색: 썬팅 = 네이비 + 블루 / PPF = 검정 + 노랑
THEMES={'tint':dict(DK=(11,18,32),B=(74,140,255),BL=(47,111,228),CARD=(22,32,52),SUB=(175,185,205),dk_hex='#0b1220',ac_hex='#2f6fe4'),
        'ppf': dict(DK=(14,14,14),B=(253,181,33),BL=(253,181,33),CARD=(30,30,30),SUB=(190,190,190),dk_hex='#0e0e0e',ac_hex='#FDB521')}
T=dict(THEMES['tint'])
def theme(name):
    T.clear(); T.update(THEMES[name]); globals().update({k:v for k,v in T.items() if k.isupper()}); return T
theme('tint')
def ease(t): t=max(0,min(1,t)); return 1-(1-t)**3
def mix(a,b,p): p=max(0,min(1,p)); return tuple(int(a[i]+(b[i]-a[i])*p) for i in range(3))
def txt(d,xy,s,f,fill,a=1,bg=None):
    bg=T['DK'] if bg is None else bg
    d.text(xy,s,font=f,fill=mix(bg,fill,a))
def ctr(d,x,w,y,s,f,fill,a=1,bg=None): txt(d,(x+w/2-d.textlength(s,font=f)/2,y),s,f,fill,a,bg)
def save(fr,name,dur=70,hold=32):
    """프레임 목록을 GIF로. 끝 장면을 hold 프레임 멈춘 뒤 반복. 56색 팔레트로 용량을 줄인다."""
    fr=fr+[fr[-1]]*hold
    p=[f.convert('P',palette=Image.ADAPTIVE,colors=56) for f in fr]
    p[0].save(name,save_all=True,append_images=p[1:],duration=dur,loop=0,optimize=True)
def tile_last_frames(gifs,out='gtile.jpg'):
    """GIF들의 마지막 장면을 2열로 모아 한 장으로 — Read 로 겹침 · 넘침을 직접 본다."""
    ims=[]
    for g in gifs:
        im=Image.open(g); im.seek(im.n_frames-1); ims.append(im.convert('RGB').resize((430,im.height//2)))
    h=(len(ims)+1)//2; cols=[ims[:h],ims[h:]]
    H=max(sum(i.height for i in c) for c in cols); t=Image.new('RGB',(860,H),(255,0,255))
    for ci,c in enumerate(cols):
        y=0
        for i in c: t.paste(i,(ci*430,y)); y+=i.height
    t.save(out,quality=85); return out

class Page:
    """스마트스토어 HTML 입력란에 통째로 붙여넣는 상세페이지.
    규칙: max-width 860 · 인라인 스타일만 · 위에서 아래로 쌓기(table · flex · grid · float · border-radius 금지) · 본문 17px."""
    def __init__(s,theme_name,sha,folder,prefix):
        s.t=THEMES[theme_name]; s.U=f'https://cdn.jsdelivr.net/gh/martingarrix1/motif@{sha}/{folder}/{prefix}'
        s.n=1; s.h=['<div style="max-width:860px;margin:0 auto;font-family:\'Pretendard\',\'Noto Sans KR\',\'Apple SD Gothic Neo\',sans-serif;background:#fff;">']
        s.ac=s.t['ac_hex']; s.dk=s.t['dk_hex']; s.btn_tx='#111' if theme_name=='ppf' else '#fff'
    W='#fff'; G='#f3f4f7'
    def add(s,x): s.h.append(x); return s
    def img(s,f,alt): return s.add(f'<img src="{s.U}{f}" alt="{alt}" style="display:block;width:100%;margin:0;">')
    def cap(s,t): return s.add(f'<p style="margin:0;padding:12px 44px 0;font-size:13px;color:#555;background:#fff;">{t}</p>')
    # 섹션 안에 넣는 조각 (문자열 반환)
    def P(s,t): return f'<p style="margin:0 0 14px;font-size:17px;line-height:1.75;color:#2a2d35;font-weight:400;letter-spacing:-0.2px;">{t}</p>'
    def sub(s,t): return f'<p style="margin:22px 0 8px;font-size:19px;color:#11141c;font-weight:800;letter-spacing:-0.4px;">{t}</p>'
    def spec(s,a,b): return f'<p style="margin:0 0 10px;padding:16px 18px;background:#fff;border-left:5px solid {s.ac};font-size:17px;line-height:1.6;color:#2a2d35;"><span style="font-weight:800;color:#11141c;">{a}</span>　{b}</p>'
    def gap(s,h=14): return f'<div style="height:{h}px;"></div>'
    def chk(s,t,ok=True): return f'<p style="margin:0 0 12px;padding:0 0 12px;border-bottom:1px solid #e3e6ee;font-size:17px;line-height:1.65;color:#2a2d35;"><span style="display:inline-block;width:26px;font-weight:900;color:{s.ac if ok else "#d9473f"};">{"✓" if ok else "✕"}</span>{t}</p>'
    def faq(s,q,a): return f'<div style="border-top:1px solid #d9dde6;padding:22px 0;"><p style="margin:0 0 10px;font-size:18px;line-height:1.5;font-weight:800;color:#11141c;"><span style="color:{s.ac};font-weight:900;">Q</span>　{q}</p><p style="margin:0;font-size:17px;line-height:1.75;color:#2a2d35;">{a}</p></div>'
    def faq_end(s): return '<div style="border-top:1px solid #d9dde6;"></div>'
    def h3(s,t,top=0): return f'<p style="margin:{top}px 0 14px;font-size:18px;color:#11141c;font-weight:800;">{t}</p>'
    # 블록
    def sec(s,tag,head,body,bg):
        s.n+=1
        return s.add(f'<div style="background:{bg};padding:64px 44px;"><p style="margin:0 0 6px;font-size:54px;line-height:1;font-weight:900;color:{s.ac};letter-spacing:-2px;">{s.n:02d}</p><p style="margin:0 0 14px;font-size:13px;font-weight:700;color:#777;letter-spacing:4px;">{tag}</p><p style="margin:0 0 30px;font-size:28px;line-height:1.4;font-weight:800;color:#11141c;letter-spacing:-0.8px;">{head}</p>{body}</div>')
    def fit(s,good,other,bg='#fff'):
        return s.sec('IS IT FOR YOU','먼저 확인해 보세요',s.h3('이런 분께 적합합니다')+''.join(s.chk(x) for x in good)+s.h3('이런 분께는 다른 선택을 권합니다',34)+''.join(s.chk(x,False) for x in other),bg)
    def faqs(s,items,bg='#f3f4f7'):
        return s.sec('FAQ','구매 전에 많이 물어보십니다',''.join(s.faq(q,a) for q,a in items)+s.faq_end(),bg)
    def note_dark(s,line1,line2='일정 변경은 입고 전까지 톡톡으로 말씀해 주세요.'):
        return s.add(f'<div style="background:{s.dk};padding:0 44px 40px;"><p style="margin:0 0 6px;font-size:15px;line-height:1.7;color:#fff;">{line1}</p><p style="margin:0;font-size:14px;color:#9aa3b5;">{line2}</p></div>')
    def more(s,post_title,log_no,head='이 상품의 바탕이 된 글'):
        return s.add(f'<div style="background:#fff;padding:56px 44px;"><p style="margin:0 0 10px;font-size:13px;font-weight:700;color:#777;letter-spacing:4px;">READ MORE</p><p style="margin:0 0 8px;font-size:24px;line-height:1.4;font-weight:800;color:#11141c;letter-spacing:-0.6px;">{head}</p><p style="margin:0 0 24px;font-size:16px;line-height:1.7;color:#3a3a3a;">{post_title}</p><a href="https://blog.naver.com/12rrrryn/{log_no}" style="display:block;padding:18px;background:{s.dk};color:#fff;font-size:18px;font-weight:900;text-align:center;text-decoration:none;border-left:6px solid {s.ac};">자세히 알아보기 →</a><p style="margin:10px 0 0;font-size:13px;color:#555;">모티프 네이버 블로그 원문으로 이동합니다</p></div>')
    def close(s,quote,tagline='MOTIF · 100% 예약제 · 대표 직접 시공'):
        return s.add(f'<div style="background:{s.dk};padding:72px 44px;"><p style="margin:0 0 18px;font-size:64px;line-height:0.6;font-weight:900;color:{s.ac};">“</p><p style="margin:0 0 40px;font-size:24px;line-height:1.55;color:#fff;font-weight:800;letter-spacing:-0.8px;">{quote}</p><a href="https://talk.naver.com/WCA4KN" style="display:block;margin:0 0 10px;padding:20px;background:{s.ac};color:{s.btn_tx};font-size:19px;font-weight:900;text-align:center;text-decoration:none;">네이버 톡톡으로 일정 문의 →</a><a href="tel:01048144506" style="display:block;margin:0 0 40px;padding:20px;background:#fff;color:#11141c;font-size:19px;font-weight:900;text-align:center;text-decoration:none;">전화 010-4814-4506</a><div style="border-top:1px solid #2c3140;padding:24px 0 0;"><p style="margin:0 0 4px;font-size:15px;line-height:1.6;color:#fff;font-weight:800;letter-spacing:1px;">{tagline}</p><p style="margin:0 0 6px;font-size:14px;line-height:1.75;color:#9aa3b5;">전북특별자치도 전주시 덕진구 정여립로 1104</p><p style="margin:0;font-size:14px;line-height:1.75;color:#9aa3b5;">네이버 톡톡 문의 <a href="https://talk.naver.com/WCA4KN" style="color:{s.ac};text-decoration:none;">talk.naver.com/WCA4KN</a></p></div></div>')
    def write(s,path):
        t='\n'.join(s.h+['</div>'])+'\n'; open(path,'w',encoding='utf-8').write(t)
        bad=[w for w in ('<table','display:flex','display:grid','float:','border-radius') if w in t]
        assert not bad, f'스마트스토어에서 지워지는 서식: {bad}'
        return len(t.encode()),len(t),t.count('.gif')

def render(html,repo_dir,out='shotm.jpg',height=15000,cols=4):
    """CDN 주소를 file:// 로 바꿔 headless chrome 으로 캡처 → 4단으로 잘라 한 장으로(Read 로 확인). GIF는 첫 프레임만 보인다."""
    t=open(html,encoding='utf-8').read()
    t=re.sub(r'https://cdn\.jsdelivr\.net/gh/martingarrix1/motif@[0-9a-f]+/',f'file://{os.path.abspath(repo_dir)}/',t)
    open('local.html','w',encoding='utf-8').write(t)
    subprocess.run(['google-chrome','--headless=new','--no-sandbox','--hide-scrollbars',f'--window-size=860,{height}','--screenshot=shot.png','file://'+os.path.abspath('local.html')],capture_output=True)
    import numpy as np
    im=Image.open('shot.png').convert('RGB'); a=np.array(im); H=int(np.where((a<250).any(axis=(1,2)))[0].max())+1; im=im.crop((0,0,860,H))
    seg=(H+cols-1)//cols; t2=Image.new('RGB',(cols*430,seg//2),(255,255,255))
    for k in range(cols):
        c=im.crop((0,k*seg,860,min(H,(k+1)*seg))); t2.paste(c.resize((430,c.height//2)),(k*430,0))
    t2.save(out,quality=80); return out,H
def preview_mp4(gifs,out,bg='0x0b1220'):
    """폰에서는 GIF가 열리지 않으므로 GIF들을 이어 붙인 MP4 미리보기를 만든다."""
    hmax=max(Image.open(g).height for g in gifs); hmax+=hmax%2; parts=[]
    for g in gifs:
        m=g[:-4]+'.mp4'; subprocess.run(['ffmpeg','-y','-loglevel','error','-i',g,'-vf',f'fps=15,pad=860:{hmax}:0:({hmax}-ih)/2:color={bg},format=yuv420p','-c:v','libx264','-crf','23',m],check=True); parts.append(m)
    open('list.txt','w').write(''.join(f"file '{os.path.abspath(p)}'\n" for p in parts))
    subprocess.run(['ffmpeg','-y','-loglevel','error','-f','concat','-safe','0','-i','list.txt','-c','copy',out],check=True); return out
def check_urls(html,extra=()):
    """push 가 끝나고 20초 뒤에 부른다(그 전에 열면 CDN이 404를 기억한다). 모든 주소가 200인지 두 번 확인."""
    import time,urllib.request
    urls=sorted(set(re.findall(r'https://cdn\.jsdelivr[^"]*',open(html,encoding='utf-8').read())))+list(extra); ok=True
    for r in range(2):
        codes=[]
        for u in urls:
            try: codes.append(urllib.request.urlopen(urllib.request.Request(u,method='HEAD'),timeout=30).status)
            except Exception as e: codes.append(getattr(e,'code',0))
        print(codes); ok=ok and all(c==200 for c in codes); time.sleep(5)
    return ok
