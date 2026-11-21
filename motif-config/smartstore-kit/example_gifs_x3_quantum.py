from PIL import Image, ImageDraw, ImageFont
FB='/usr/share/fonts/opentype/noto/NotoSansCJK-Black.ttc'; FBo='/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc'; FR='/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'
def F(s,f=FBo): return ImageFont.truetype(f,s,index=1)
DK=(11,18,32); B=(74,140,255); WH=(255,255,255); LG=(243,244,247); INK=(17,20,28); SUB=(175,185,205); HOT=(255,120,90)
def ease(t): t=max(0,min(1,t)); return 1-(1-t)**3
def save(fr,name,dur,hold=32):
    fr=fr+[fr[-1]]*hold
    p=[f.convert('P',palette=Image.ADAPTIVE,colors=56) for f in fr]
    p[0].save(name,save_all=True,append_images=p[1:],duration=dur,loop=0,optimize=True)
def txt(d,xy,s,f,fill,a=1,bg=DK):
    a=max(0,min(1,a)); c=tuple(int(bg[i]+(fill[i]-bg[i])*a) for i in range(3)); d.text(xy,s,font=f,fill=c)
W=860
import math
GR=(150,155,165); CARD=(22,32,52); LINE=(222,226,234); TX=(26,28,34); BL=(47,111,228); SUN=(255,190,90)
def mix(a,b,p): p=max(0,min(1,p)); return tuple(int(a[i]+(b[i]-a[i])*p) for i in range(3))
def ctr(d,x,w,y,s,f,fill,a,bg): txt(d,(x+w/2-d.textlength(s,font=f)/2,y),s,f,fill,a,bg)
# 1 hero
H=660; fr=[]
for i in range(48):
    im=Image.new('RGB',(W,H),DK); d=ImageDraw.Draw(im)
    txt(d,(44,44),'BMW X3  /  SOLAR GARD QUANTUM',F(17),B,ease(i/8))
    p=ease((i-3)/8); txt(d,(44-int((1-p)*60),92),'밖에서는 가려지고',F(54,FB),WH,p)
    p=ease((i-9)/8); txt(d,(44,166+int((1-p)*26)),'안에서는 맑게 보이는 유리',F(54,FB),B,p)
    d.rectangle([44,262,44+int(772*ease((i-15)/8)),266],fill=B)
    for k,s in enumerate(['풀메탈 스퍼터링','전면 19%','측후면 14%','보증 7년']):
        a=ease((i-18-k*3)/6); x=44+k*196
        d.rectangle([x,292,x+184,346],outline=mix(DK,(70,80,100),a),width=2); ctr(d,x,184,307,s,F(19),WH,a,DK)
    a=ease((i-29)/6); f=F(30); s='정상가 1,600,000원'; txt(d,(44,392),s,f,(130,140,160),a); w=d.textlength(s,font=f)
    d.line([44,412,44+int(w*ease((i-32)/6)),412],fill=HOT,width=3)
    q=ease((i-33)/10); v=int(round((1600000-320000*q)/1000)*1000); s2=f'{v:,}'; f2=F(88,FB)
    txt(d,(44,450),s2,f2,WH,ease((i-32)/5)); txt(d,(44+d.textlength(s2,font=f2)+12,510),'원',F(34),B,ease((i-32)/5))
    a=ease((i-40)/5); d.rectangle([44,582,44+150,626],fill=mix(DK,BL,a)); ctr(d,44,150,591,'20% 할인',F(20),WH,a,mix(DK,BL,a))
    txt(d,(214,592),'전면 + 측후면 · 선루프 제외',F(20,FR),SUB,ease((i-42)/5))
    fr.append(im)
save(fr,'mo-hero.gif',70)
# 2 dilemma
H=560; fr=[]
for i in range(50):
    im=Image.new('RGB',(W,H),LG); d=ImageDraw.Draw(im)
    txt(d,(44,34),'농도를 낮추면 생기는 일',F(30),INK,ease(i/6),LG)
    q=ease((i-8)/26)
    for k,(nm,v0,v1,c) in enumerate([('외부에서 본 프라이버시',30,88,BL),('운전자의 시야',88,34,(217,71,63))]):
        a=ease((i-3-k*3)/6); y=120+k*150; v=v0+(v1-v0)*q
        txt(d,(44,y),nm,F(24),TX,a,LG)
        d.rectangle([44,y+44,816,y+96],fill=(226,229,236)); d.rectangle([44,y+44,44+int(772*v/100*a),y+96],fill=c)
        s='상승' if k==0 else '저하'; f=F(24); txt(d,(816-d.textlength(s,font=f),y),s,f,c,ease((i-30)/6),LG)
    a=ease((i-34)/7); bg=mix(LG,WH,a); d.rectangle([44,430,816,510],fill=bg,outline=LINE); d.rectangle([44,430,50,510],fill=BL)
    txt(d,(74,454),'투과율의 저하는 방향을 가리지 않습니다',F(26),TX,a,bg)
    txt(d,(44,H-32),'개념도 · 이해를 돕기 위한 연출',F(15,FR),(130,135,145),ease((i-38)/6),LG)
    fr.append(im)
save(fr,'mo-dilemma.gif',70)
# 3 absorb vs reflect
H=640; fr=[]
def ray(d,x0,y0,x1,y1,p,col,wd=5):
    p=max(0,min(1,p)); 
    if p<=0: return
    d.line([x0,y0,x0+(x1-x0)*p,y0+(y1-y0)*p],fill=col,width=wd)
for i in range(54):
    im=Image.new('RGB',(W,H),DK); d=ImageDraw.Draw(im)
    txt(d,(44,34),'흡수로 가리는 필름 · 반사로 가리는 필름',F(28),WH,ease(i/6))
    for k,(nm,sub) in enumerate([('세라믹 필름','빛을 흡수 · 프라이버시는 농도에 의존'),('퀀텀','금속 박막이 빛을 반사')]):
        a=ease((i-4-k*4)/7); x=44+k*396; bg=mix(DK,CARD,a)
        d.rectangle([x,100,x+376,560],fill=bg); d.rectangle([x,100,x+376,104],fill=B if k else GR)
        txt(d,(x+24,122),nm,F(30),WH,a,bg); txt(d,(x+24,166),sub,F(17,FR),SUB,ease((i-8-k*4)/6),bg)
        gx=x+210
        d.rectangle([gx,220,gx+14,500],fill=mix(bg,(120,170,220) if k==0 else (190,205,225),a))
        txt(d,(x+24,512),'외부',F(16),SUB,a,bg); txt(d,(x+316,512),'실내',F(16),SUB,a,bg)
        t=i-14-k*4
        for j in range(3):
            y0=250+j*80; pp=ease((t-j*2)/10)
            ray(d,x+30,y0,gx,y0+30,pp,SUN)
            if k==0:
                rr=ease((t-j*2-10)/8)
                if rr>0: r=int(6+14*rr); d.ellipse([gx+7-r,y0+30-r,gx+7+r,y0+30+r],outline=mix(bg,(255,140,90),1-rr*0.5),width=3)
                ray(d,gx+14,y0+30,x+290,y0+44,ease((t-j*2-12)/10),mix(bg,SUN,0.3),2)
            else:
                ray(d,gx,y0+30,x+60,y0+62,ease((t-j*2-10)/10),(120,200,255))
                ray(d,gx+14,y0+30,x+290,y0+44,ease((t-j*2-12)/10),mix(bg,SUN,0.3),2)
    txt(d,(44,H-56),'동일한 농도라면 반사 구조가 프라이버시에 유리합니다',F(22),WH,ease((i-40)/7))
    txt(d,(44,H-24),'개념도 · 이해를 돕기 위한 연출',F(14,FR),(120,130,150),ease((i-44)/6))
    fr.append(im)
save(fr,'mo-reflect.gif',70)
# 4 day/night
H=520; fr=[]
for i in range(44):
    im=Image.new('RGB',(W,H),LG); d=ImageDraw.Draw(im)
    txt(d,(44,34),'반사 필름의 프라이버시가 작동하는 조건',F(28),INK,ease(i/6),LG)
    for k,(t1,t2,t3,v,c) in enumerate([('DAY','주간','외부가 실내보다 밝을 때',92,BL),('NIGHT','야간 · 실내등 점등','조도의 방향이 역전될 때',30,GR)]):
        p=ease((i-4-k*10)/8); y=100+k*190; x=44+int((1-p)*90)
        if p<=0: continue
        bg=mix(LG,WH,p); d.rectangle([x,y,x+772,y+170],fill=bg,outline=LINE); d.rectangle([x,y,x+6,y+170],fill=c)
        txt(d,(x+30,y+18),t1,F(16),c,p,bg); txt(d,(x+30,y+44),t2,F(30),TX,p,bg); txt(d,(x+30,y+90),t3,F(19,FR),(90,95,105),ease((i-8-k*10)/6),bg)
        q=ease((i-10-k*10)/12); d.rectangle([x+30,y+130,x+742,y+146],fill=(232,235,241)); d.rectangle([x+30,y+130,x+30+int(712*v/100*q),y+146],fill=c)
        s='효과 뚜렷' if k==0 else '효과 감소'; f=F(22); txt(d,(x+742-d.textlength(s,font=f),y+48),s,f,c,ease((i-14-k*10)/6),bg)
    txt(d,(44,H-32),'개념도 · 이해를 돕기 위한 연출',F(15,FR),(130,135,145),ease((i-34)/6),LG)
    fr.append(im)
save(fr,'mo-daynight.gif',70)
# 5 sputtering
H=600; fr=[]
import random
random.seed(7); pts=[(random.random(),random.random(),random.randint(3,9)) for _ in range(46)]
for i in range(50):
    im=Image.new('RGB',(W,H),DK); d=ImageDraw.Draw(im)
    txt(d,(44,34),'시인성은 박막의 균질도에서 결정됩니다',F(28),WH,ease(i/6))
    for k,(nm,sub) in enumerate([('염료 필름','불균일한 입자 · 빛의 산란'),('마그네트론 스퍼터링','나노미터 단위의 균일 증착')]):
        a=ease((i-4-k*5)/7); x=44+k*396; bg=mix(DK,CARD,a)
        d.rectangle([x,100,x+376,520],fill=bg); d.rectangle([x,100,x+376,104],fill=B if k else GR)
        txt(d,(x+24,122),nm,F(28),WH,a,bg); txt(d,(x+24,164),sub,F(17,FR),SUB,ease((i-8-k*5)/6),bg)
        d.rectangle([x+24,220,x+352,420],outline=mix(bg,(70,80,100),a),width=2)
        if k==0:
            for j,(px,py,r) in enumerate(pts):
                aa=ease((i-12-j*0.4)/6)
                if aa>0: cx=x+34+px*308; cy=230+py*180; rr=r*aa; d.ellipse([cx-rr,cy-rr,cx+rr,cy+rr],fill=mix(bg,(150,120,170),aa))
        else:
            n=0
            for ry in range(6):
                for rx in range(12):
                    aa=ease((i-16-(ry*12+rx)*0.3)/5)
                    if aa>0: cx=x+46+rx*26; cy=250+ry*28; d.ellipse([cx-7,cy-7,cx+7,cy+7],fill=mix(bg,(170,200,235),aa))
        s='풍경이 흐려짐' if k==0 else '왜곡 없이 전달'; ctr(d,x,376,450,s,F(24),GR if k==0 else B,ease((i-34)/6),bg)
    txt(d,(44,H-34),'개념도 · 이해를 돕기 위한 연출',F(14,FR),(120,130,150),ease((i-40)/6))
    fr.append(im)
save(fr,'mo-sputter.gif',70)
# 6 shade
H=560; fr=[]
sh=[52,37,28,19,14,8]
for i in range(48):
    im=Image.new('RGB',(W,H),LG); d=ImageDraw.Draw(im)
    txt(d,(44,34),'퀀텀 농도 6단계',F(30),INK,ease(i/6),LG); txt(d,(44,82),'숫자가 낮을수록 어둡습니다',F(19,FR),(90,95,105),ease((i-2)/6),LG)
    for k,v in enumerate(sh):
        a=ease((i-5-k*3)/7); x=44+k*130; y=150+int((1-a)*24)
        if a<=0: continue
        g=int(40+v*3.6); col=mix(LG,(g,int(g*1.04),int(g*1.02)),a)
        d.rectangle([x,y,x+118,y+200],fill=col); f=F(40,FB); ctr(d,x,118,y+70,str(v),f,WH if v<40 else TX,a,col)
        hl={19:'전면',14:'측후면'}.get(v)
        if hl:
            b=ease((i-28-(0 if v==19 else 4))/7)
            if b>0:
                d.rectangle([x-4,y-4,x+122,y+204],outline=mix(LG,BL,b),width=5)
                bg=mix(LG,BL,b); d.rectangle([x,y+216,x+118,y+260],fill=bg); ctr(d,x,118,y+225,hl,F(21),WH,b,bg)
    a=ease((i-38)/7); bg=mix(LG,WH,a); d.rectangle([44,440,816,516],fill=bg,outline=LINE); d.rectangle([44,440,50,516],fill=BL)
    txt(d,(74,462),'이 차량: 전면 19% · 측후면 14%',F(26),TX,a,bg)
    fr.append(im)
save(fr,'mo-shade.gif',70)
# 7 process
H=680; fr=[]
ps=[('구매','스마트스토어 결제'),('톡톡으로 일정 예약','100% 예약제'),('입고와 농도 상담','주행 환경 기준으로 결정'),('유리 세정과 시공','대표 직접 시공'),('점검과 출고','끝단 마감과 시야 확인')]
for i in range(44):
    im=Image.new('RGB',(W,H),DK); d=ImageDraw.Draw(im)
    txt(d,(44,34),'진행 순서',F(30),WH,ease(i/6))
    d.rectangle([73,120,75,120+int(480*ease((i-4)/26))],fill=(40,60,100))
    for k,(a,b) in enumerate(ps):
        p=ease((i-5-k*5)/7); y=104+k*112
        if p<=0: continue
        d.ellipse([54,y+6,94,y+46],fill=mix(DK,B,p)); ctr(d,54,40,y+11,str(k+1),F(22),WH,p,mix(DK,B,p))
        txt(d,(124+int((1-p)*40),y),a,F(28),WH,p); txt(d,(124,y+44),b,F(19,FR),SUB,ease((i-8-k*5)/6))
    fr.append(im)
save(fr,'mo-process.gif',70)
