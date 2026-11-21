# MOTIF 썬팅 무빙 영상 템플릿 (화이트 스튜디오 · 피아노 · 구역별 수치 카드)
# 사용: tint_moving.py stills|render config.json
import os as _os
_FD=_os.path.join(_os.path.dirname(_os.path.abspath(__file__)),"..","fonts")   # 묶음에 든 폰트를 먼저 쓴다(맥 · 리눅스 공통)
def _font(name,fallback):
    p=_os.path.join(_FD,name); return p if _os.path.exists(p) else fallback
import subprocess, numpy as np, sys, json, math, glob
from PIL import Image, ImageDraw, ImageFont
CFG=json.load(open(sys.argv[2])); W,H,FPS=1080,1920,30; XF=CFG.get("xf",1.2)   # 컷 사이 디졸브 길이(초)
# 솔라가드 THE LX 수치 (TSER, VLT, VLT/SHGC) — 확인된 값만. 없는 농도는 추가하지 말고 마틴에게 확인
FILM={"LX 15":("71","12","0.41"),"LX 25":("65","23","0.66"),"LX 40":("61","42","1.08")}
ZONE={"front":("FRONT","전면 유리"),"side":("SIDE & REAR","측면 · 후면 유리")}
SEQ=[(c["file"],c["dur"],c.get("zone")) for c in CFG["clips"]]
st=[];t=0
for n,d,z in SEQ: st.append(t); t+=d-XF
CLIP_END=t+XF; DIM0=CLIP_END-4.3; TXT0=DIM0+0.4; END0=CLIP_END-0.5; TOTAL=END0+3.6
B=CFG.get("brand","brand/")
logo=Image.open(B+'motif_logo_dark.png').convert('RGBA').crop((4,3,796,187))
def lg(w,a=1.0):
    im=logo.resize((w,int(logo.height*w/logo.width)),Image.LANCZOS)
    if a<1: im.putalpha(im.getchannel('A').point(lambda v:int(v*a)))
    return im
PF=_font("Poppins-Bold.ttf","/usr/share/fonts/truetype/google-fonts/Poppins-Bold.ttf").replace("Poppins-Bold","Poppins-%s"); _f={}
def P(w,s):
    if (w,s) not in _f: _f[(w,s)]=ImageFont.truetype(PF%w,s)
    return _f[(w,s)]
EN=_font('InterDisplay-Medium.otf','/usr/share/fonts/opentype/inter/InterDisplay-Medium.otf'); ENB=_font('InterDisplay-Bold.otf','/usr/share/fonts/opentype/inter/InterDisplay-Bold.otf')
KRM=_font('NotoSansCJK-Medium.ttc',([f for f in glob.glob('/usr/share/fonts/opentype/noto/NotoSansCJK-*.ttc') if 'Medium' in f or 'Regular' in f] or [''])[0])
def F(p,s):
    if (p,s) not in _f: _f[(p,s)]=ImageFont.truetype(p,s)
    return _f[(p,s)]
INK=(22,24,30); SUB=(96,100,112)
def expo(p): p=max(0,min(1,p)); return 1-2**(-10*p) if p<1 else 1.0
def ease(u): u=max(0,min(1,u)); return 1-(1-u)**3
S=2; BH=400; BY=1230
_g=np.zeros((H,W,4),np.uint8); _g[...,:3]=250; _yy=np.clip((np.arange(H)-1120)/190,0,1); _g[...,3]=(255*0.9*(_yy**1.5))[:,None].astype(np.uint8); SCR=Image.fromarray(_g)
def sptext(d,t_,f,cx,y,sp,fill):
    tw=sum(d.textlength(ch,font=f)+sp for ch in t_)-sp; x=cx-tw/2
    for ch in t_: d.text((x,y),ch,font=f,fill=fill,anchor="lm"); x+=d.textlength(ch,font=f)+sp
def block(zone,ts,dur):
    film=CFG["film"][zone]; v=FILM[film]; cap=f"{ZONE[zone][0]}  —  SOLAR GARD {film}"; cols=[("TSER",v[0]),("VLT",v[1]),("VLT/SHGC",v[2])]
    im=Image.new("RGBA",(W*S,BH*S),(0,0,0,0)); d=ImageDraw.Draw(im)
    out=max(0,min(1,(ts-(dur-0.45))/0.45)); a_all=1-out; A=lambda v_:int(255*max(0,min(1,v_))*a_all)
    e=expo(ts/0.6); sptext(d,cap,P("Bold",28*S),W*S/2,(40+14*(1-e))*S,6*S,INK+(A(ts/0.3),))
    hw=440*expo((ts-0.1)/0.8); d.line([(W/2-hw)*S,84*S,(W/2+hw)*S,84*S],fill=(120,124,136,A(1)),width=3)
    n=len(cols); cw=880/3; x0=W/2-cw*n/2; fn=P("Bold",128*S); RH=176*S; cy=262*S
    row=Image.new("RGBA",(W*S,RH),(0,0,0,0)); r=ImageDraw.Draw(row); yo=-RH*out*out
    for c,(lab,val) in enumerate(cols):
        cx=(x0+cw*(c+0.5))*S; t0=0.2+c*0.15
        sptext(d,lab,P("Medium",30*S),cx,(138+10*(1-expo((ts-t0)/0.5)))*S,6*S,SUB+(A((ts-t0)/0.3),))
        if c: xl=(x0+cw*c)*S; d.line([xl,124*S,xl,(124+210*expo((ts-t0)/0.7))*S],fill=(150,154,166,A(1)),width=2)
        ws=[r.textlength(ch,font=fn) if ch=="." else max(r.textlength(str(k),font=fn) for k in range(10))*0.94 for ch in val]; x=cx-sum(ws)/2; k=0
        for ch,w_ in zip(val,ws):
            xc=x+w_/2; x+=w_
            if ch==".":
                if ts>t0+0.3: r.text((xc,RH/2+yo),ch,font=fn,fill=INK+(255,),anchor="mm")
                continue
            T0=t0+k*0.1; D=0.85+k*0.12; k+=1
            if ts<T0: continue
            p=max(0,min(1,(ts-T0)/D)); pe=1-(1-p)**4; turns=14+6*k; pos=(int(ch)-turns)+turns*pe; base=math.floor(pos); fr=pos-base
            for j in ((0,1) if fr>0.002 else (0,)): r.text((xc,RH/2+(j-fr)*RH*0.92+yo),str((base+j)%10),font=fn,fill=INK+(255,),anchor="mm")
    m=np.ones((RH,1),np.float32); ed=int(RH*0.14); m[:ed,0]=np.linspace(0,1,ed); m[-ed:,0]=np.linspace(1,0,ed)
    arr=np.array(row); arr[...,3]=(arr[...,3]*m*a_all**0.5).astype(np.uint8); im.alpha_composite(Image.fromarray(arr),(0,int(cy-RH/2))); return im.resize((W,BH),Image.LANCZOS)
def spx(d,xy,txt,font,fill,track):
    x,y=xy
    for ch in txt: d.text((x,y),ch,font=font,fill=fill); x+=d.textlength(ch,font=font)+track
def spw(d,txt,font,track): return sum(d.textlength(c,font=font)+track for c in txt)-track
LINES=[("small",CFG["car"],0.0),("rule",None,0.25)]
dl=0.45
for z in ("front","side"):
    if z in CFG["film"]: LINES+= [("lab",ZONE[z][1],dl),("big","THE "+CFG["film"][z],dl+0.15)]; dl+=0.55
LINES.append(("small2","SOLAR GARD THE LX",dl+0.15))
def endtext(t):
    ov=Image.new('RGBA',(W,H),(0,0,0,0)); d=ImageDraw.Draw(ov); y=640
    for kind,txt,dl in LINES:
        u=ease((t-TXT0-dl)/0.9); a=int(255*u); dy=int(46*(1-u))
        if kind=="small": f=F(EN,40); w=spw(d,txt,f,7); spx(d,((W-w)/2,y+dy),txt,f,(255,255,255,int(a*.82)),7); y+=96
        elif kind=="rule": w=int(150*u); d.rectangle((W/2-w,y,W/2+w,y+2),fill=(255,255,255,int(a*.7))); y+=84
        elif kind=="lab": f=F(KRM,40); w=spw(d,txt,f,5); spx(d,((W-w)/2,y+dy),txt,f,(255,255,255,int(a*.78)),5); y+=70
        elif kind=="big": f=F(ENB,118); w=spw(d,txt,f,4); spx(d,((W-w)/2,y+dy),txt,f,(255,255,255,a),4); y+=205
        else: f=F(EN,32); w=spw(d,txt,f,10); y+=10; spx(d,((W-w)/2,y+dy),txt,f,(255,255,255,int(a*.7)),10)
    return ov
import cv2
def zoom(a,s_):
    if s_<=1.0005: return a
    return cv2.warpAffine(a,np.float32([[s_,0,W/2*(1-s_)],[0,s_,H/2*(1-s_)]]),(W,H),flags=cv2.INTER_LINEAR,borderMode=cv2.BORDER_REPLICATE)
def rd(n): return subprocess.Popen(["ffmpeg","-loglevel","error","-i",n,"-f","rawvideo","-pix_fmt","rgb24","-"],stdout=subprocess.PIPE)
def render(out,stills=None):
    procs=[None]*len(SEQ); cur={}; N=int(TOTAL*FPS); enc=None; last=None; outs=[]
    if not stills: enc=subprocess.Popen(["ffmpeg","-loglevel","error","-y","-f","rawvideo","-pix_fmt","rgb24","-s",f"{W}x{H}","-r",str(FPS),"-i","-","-c:v","libx264","-crf","16","-preset","medium","-pix_fmt","yuv420p","-movflags","+faststart",out],stdin=subprocess.PIPE)
    for k in range(N):
        t=k/FPS; fr=None; bloom=0.0
        for i,(n,d,z) in enumerate(SEQ):
            if st[i]<=t<st[i]+d:
                if procs[i] is None: procs[i]=rd(n)
                b=procs[i].stdout.read(W*H*3)
                if len(b)==W*H*3: cur[i]=np.frombuffer(b,np.uint8).reshape(H,W,3).astype(np.float32)
                a_=cur.get(i)
                if a_ is None: continue
                w=1.0
                if i>0 and t<st[i]+XF:
                    u=(t-st[i])/XF; w=u*u*u*(u*(u*6-15)+10)      # smootherstep: 시작과 끝이 멈칫하지 않게
                    a_=zoom(a_,1.05-0.05*w)                      # 들어오는 컷은 살짝 당겨진 상태에서 풀리고
                    if fr is not None: fr=zoom(fr,1+0.05*w)      # 나가는 컷은 살짝 밀고 들어가며 사라진다
                    bloom=max(bloom,math.sin(math.pi*u)*0.16)
                fr=a_ if fr is None else fr*(1-w)+a_*w
        if fr is None: fr=last
        last=fr
        if bloom>0: fr=fr+(255-fr)*bloom       # 디졸브 중간에 흰빛이 살짝 번진다(화이트 스튜디오와 이어지게)
        if stills and not any(abs(t-s)<1/60 for s in stills): continue
        im=Image.fromarray(np.clip(fr,0,255).astype(np.uint8)).convert('RGBA')
        # 구역 카드: 그 구역이 처음 나오는 컷에서 한 번만
        for i,(n,d,z) in enumerate(SEQ):
            if z and t<DIM0 and z not in [q[2] for q in SEQ[:i]]:
                c0=st[i]+(XF if i else 0.3); c1=min(st[i]+d-XF,DIM0); ts=t-c0; dur=c1-c0
                if 0<=ts<dur:
                    sa=min(1,ts/0.5)*min(1,max(0,(dur-ts)/0.4)); sc_=SCR.copy(); sc_.putalpha(sc_.split()[3].point(lambda v:int(v*sa))); im.alpha_composite(sc_); im.alpha_composite(block(z,ts,dur),(0,BY))
        la=1-ease((t-DIM0)/0.8)
        if la>0.01: L=lg(210,la*0.92); im.alpha_composite(L,((W-L.width)//2,1722))
        dm=ease((t-DIM0)/1.2)*0.76
        if dm>0: im=Image.blend(im,Image.new('RGBA',(W,H),(0,0,0,255)),dm)
        if t>TXT0-0.05: im.alpha_composite(endtext(t))
        e=ease((t-END0)/0.9)
        if e>0:
            card=Image.new('RGBA',(W,H),(246,246,244,255)); la2=ease((t-END0-0.5)/0.9); sc=1.04-0.04*ease((t-END0)/3.0)
            L=lg(int(330*sc),la2); card.alpha_composite(L,((W-L.width)//2,(H-L.height)//2)); im=Image.blend(im,card,e)
        a=np.asarray(im.convert('RGB')).astype(np.float32); fo=min(1,(TOTAL-t)/0.7)
        if fo<1: a=a*fo
        a=a.astype(np.uint8)
        if stills: outs.append(Image.fromarray(a).resize((270,480)))
        else: enc.stdin.write(a.tobytes())
    if stills:
        sh=Image.new('RGB',(270*len(outs),480))
        for i,o in enumerate(outs): sh.paste(o,(270*i,0))
        sh.save(out)
    else: enc.stdin.close(); enc.wait()
    print("TOTAL",round(TOTAL,2))
if sys.argv[1]=='stills':
    pts=[]
    for i,(n,d,z) in enumerate(SEQ): pts+= [st[i]+1.2, st[i]+d-1.2]
    render('st.jpg',pts+[CLIP_END-1.0,END0+2])
else: render(CFG.get("out","video.mp4"))
