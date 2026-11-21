# MOTIF 차량 무빙 영상 "레드 버전" 템플릿 (레드 스튜디오 · 같은 길이 컷 · 슬로우 · 웅장한 피아노)
# 사용: red_moving.py stills|render config.json
import os as _os
_FD=_os.path.join(_os.path.dirname(_os.path.abspath(__file__)),"..","fonts")   # 묶음에 든 폰트를 먼저 쓴다(맥 · 리눅스 공통)
def _font(name,fallback):
    p=_os.path.join(_FD,name); return p if _os.path.exists(p) else fallback
import subprocess, numpy as np, sys, json, glob
from PIL import Image, ImageDraw, ImageFont
CFG=json.load(open(sys.argv[2])); W,H,FPS=1080,1920,30
SEQ=CFG["cuts"]                     # [{"file","ss","dur","speed"}] 하드 컷으로 이어 붙인다
END=CFG["end"]                      # {"file","ss","dur","speed"} 마지막 정면 컷(어두워지며 글씨 → 로고)
B=CFG.get("brand","brand/")
logoL=Image.open(B+'motif_logo_light.png').convert('RGBA'); logoL=logoL.crop(logoL.getbbox())
def lg(w,a=1.0):
    im=logoL.resize((w,int(logoL.height*w/logoL.width)),Image.LANCZOS)
    if a<1: im.putalpha(im.getchannel('A').point(lambda v:int(v*a)))
    return im
EN=_font('InterDisplay-Medium.otf','/usr/share/fonts/opentype/inter/InterDisplay-Medium.otf'); ENB=_font('InterDisplay-Bold.otf','/usr/share/fonts/opentype/inter/InterDisplay-Bold.otf')
KRM=_font('NotoSansCJK-Medium.ttc',([f for f in glob.glob('/usr/share/fonts/opentype/noto/NotoSansCJK-*.ttc') if 'Medium' in f or 'Regular' in f] or [''])[0]); _f={}
def F(p,s):
    if (p,s) not in _f: _f[(p,s)]=ImageFont.truetype(p,s)
    return _f[(p,s)]
def ease(u): u=max(0,min(1,u)); return 1-(1-u)**3
def spx(d,xy,txt,font,fill,track):
    x,y=xy
    for ch in txt: d.text((x,y),ch,font=font,fill=fill); x+=d.textlength(ch,font=font)+track
def spw(d,txt,font,track): return sum(d.textlength(c,font=font)+track for c in txt)-track
st=[];t=0
for c in SEQ: st.append(t); t+=c["dur"]
T_END=t; TXT0=T_END+0.9; LOGO0=T_END+END["dur"]-0.4; TOTAL=LOGO0+2.6
yy,xx=np.mgrid[0:H,0:W].astype(np.float32); r=np.sqrt(((xx-W/2)/(W*0.8))**2+((yy-H/2)/(H*0.8))**2)
VIG=(1-0.28*np.clip(r-0.4,0,1)**1.5)[...,None]
def rd(c,dur):
    sp=c.get("speed",1.0)
    mi="minterpolate=fps=30:mi_mode=mci:mc_mode=aobmc:me_mode=bidir:vsbmc=1:scd=none," if sp<0.99 else "fps=30,"
    vf=f"setpts=(PTS-STARTPTS)/{sp},{mi}scale=1080:1920:flags=lanczos,unsharp=5:5:0.5,eq=contrast=1.06:saturation=1.05"
    p=subprocess.run(["ffmpeg","-loglevel","error","-ss",str(c["ss"]),"-t",str(dur*sp+0.2),"-i",c["file"],"-vf",vf,"-frames:v",str(int(round(dur*FPS))+1),"-f","rawvideo","-pix_fmt","rgb24","-"],stdout=subprocess.PIPE).stdout
    return np.frombuffer(p,np.uint8).reshape(-1,H,W,3)
LINES=[("small",CFG["car"]),("rule",None)]+[("big",b) for b in CFG["big"]]+[("small2",s) for s in CFG.get("small",[])]
def endtext(t):
    ov=Image.new('RGBA',(W,H),(0,0,0,0)); d=ImageDraw.Draw(ov); y=CFG.get("text_y",700); dl=0
    for kind,txt in LINES:
        u=ease((t-TXT0-dl)/0.8); a=int(255*u); dy=int(40*(1-u)); dl+=0.22
        if kind=="small": f=F(EN,40); w=spw(d,txt,f,7); spx(d,((W-w)/2,y+dy),txt,f,(255,255,255,int(a*.85)),7); y+=92
        elif kind=="rule": w=int(150*u); d.rectangle((W/2-w,y,W/2+w,y+2),fill=(255,255,255,int(a*.7))); y+=70
        elif kind=="big": f=F(ENB,104); w=spw(d,txt,f,4); spx(d,((W-w)/2,y+dy),txt,f,(255,255,255,a),4); y+=150
        else: f=F(KRM if any(ord(ch)>0x3000 for ch in txt) else EN,34); w=spw(d,txt,f,8); y+=14; spx(d,((W-w)/2,y+dy),txt,f,(255,255,255,int(a*.75)),8); y+=60
    return ov
def render(out,stills=False):
    enc=None; outs=[]
    if not stills: enc=subprocess.Popen(["ffmpeg","-loglevel","error","-y","-f","rawvideo","-pix_fmt","rgb24","-s",f"{W}x{H}","-r",str(FPS),"-i","-","-c:v","libx264","-crf","16","-preset","medium","-pix_fmt","yuv420p","-movflags","+faststart",out],stdin=subprocess.PIPE)
    def emit(fr,t):
        f=fr.astype(np.float32)*VIG
        dm=ease((t-T_END-0.3)/1.0)*0.58 if t>T_END else 0
        if dm>0: f*=(1-dm)
        im=Image.fromarray(np.clip(f,0,255).astype(np.uint8)).convert('RGBA')
        la=1-ease((t-T_END-0.2)/0.6) if t>T_END else 1
        if la>0.01: L=lg(200,la*0.9); im.alpha_composite(L,((W-L.width)//2,1730))
        if t>TXT0-0.05: im.alpha_composite(endtext(t))
        e=ease((t-LOGO0)/0.8)
        if e>0:
            card=Image.new('RGBA',(W,H),(8,6,7,255)); la2=ease((t-LOGO0-0.4)/0.8); L=lg(int(330*(1.04-0.04*ease((t-LOGO0)/2.6))),la2)
            card.alpha_composite(L,((W-L.width)//2,(H-L.height)//2)); im=Image.blend(im,card,e)
        a=np.asarray(im.convert('RGB')).astype(np.float32); fo=min(1,(TOTAL-t)/0.6)
        if fo<1: a=a*max(fo,0)
        return a.astype(np.uint8)
    k=0
    for i,c in enumerate(SEQ):
        n1=int(round((st[i]+c["dur"])*FPS)); fr=rd(c,c["dur"]); j=0
        while k<n1:
            f=fr[min(j,len(fr)-1)]
            if stills:
                if j==int(c["dur"]*FPS*0.5): outs.append(Image.fromarray(emit(f,k/FPS)).resize((180,320)))
            else: enc.stdin.write(emit(f,k/FPS).tobytes())
            k+=1;j+=1
    fr=rd(END,END["dur"]+0.5); N=int(round(TOTAL*FPS)); j=0
    while k<N:
        f=fr[min(j,len(fr)-1)]
        if stills:
            if j in (int(1.0*FPS),int((END["dur"]-0.8)*FPS),int((END["dur"]+1.4)*FPS)): outs.append(Image.fromarray(emit(f,k/FPS)).resize((180,320)))
        else: enc.stdin.write(emit(f,k/FPS).tobytes())
        k+=1;j+=1
    if stills:
        cols=12; rows=(len(outs)+cols-1)//cols; sh=Image.new('RGB',(180*cols,320*rows))
        for i,o in enumerate(outs): sh.paste(o,(180*(i%cols),320*(i//cols)))
        sh.save(out)
    else: enc.stdin.close(); enc.wait()
    print("T_END",round(T_END,3),"TOTAL",round(TOTAL,2))
if sys.argv[1]=='stills': render('st.jpg',True)
else: render(CFG.get("out","video.mp4"))
