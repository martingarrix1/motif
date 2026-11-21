# 웅장한 피아노곡 (D단조 · 4/4 · 88bpm) — 실제 어쿠스틱 피아노 샘플 연주
# 사용: piano_grand.py <길이초> <out.wav> <샘플폴더>   (한 마디 = 2.727초 = 한 컷)
import numpy as np,sys,wave,subprocess
from scipy.signal import lfilter
SR=44100; DUR=float(sys.argv[1]); OUT=sys.argv[2]; D=sys.argv[3] if len(sys.argv)>3 else "pn"; rng=np.random.default_rng(1121)
NM=["C","Cs","D","Ds","E","F","Fs","G","Gs","A","As","B"]; _c={}
def smp(m):
    if m not in _c:
        raw=subprocess.run(["ffmpeg","-loglevel","error","-i",f"{D}/{NM[m%12]}{m//12-1}.mp3","-f","f32le","-ac","2","-ar",str(SR),"-"],stdout=subprocess.PIPE).stdout
        a=np.frombuffer(raw,np.float32).reshape(-1,2).astype(np.float64); pk=np.abs(a).max(); a=a/pk if pk>0 else a
        on=np.argmax(np.abs(a).max(1)>0.02); _c[m]=a[max(0,on-60):]
    return _c[m]
tot=np.zeros((int(SR*(DUR+8)),2))
def note(t0,m,dur,vel):
    m=int(np.clip(m,24,107)); a=smp(m); n=min(len(a),int(SR*(dur+0.3)))
    if n<=0 or t0<0: return
    s=a[:n].copy(); rel=int(SR*0.3); s[-rel:]*=np.linspace(1,0,rel)[:,None]**1.5
    k=0.3+0.7*min(1,vel/0.95)
    if k<0.98: s=lfilter([k],[1,-(1-k)],s,axis=0)
    pan=np.clip((m-60)/60,-0.3,0.3); g=np.array([1-pan,1+pan])*vel
    i=int(t0*SR); e=min(len(tot),i+n); tot[i:e]+=s[:e-i]*g
beat=60/88; bar=beat*4
# i – VI – III – VII – i – iv – V – i (Dm Bb F C | Dm Gm A Dm)
P=[(26,[62,65,69]),(22,[58,62,65]),(29,[60,65,69]),(24,[60,64,67]),(26,[62,65,69]),(31,[58,62,67]),(33,[61,64,69]),(26,[62,65,69])]
MEL=[[(0,74,2),(2,77,1),(3,76,1)],[(0,74,3),(3,70,1)],[(0,72,2),(2,77,2)],[(0,76,2),(2,72,1),(3,76,1)],
     [(0,81,2),(2,77,1),(3,76,1)],[(0,74,2),(2,79,2)],[(0,76,1.5),(1.5,73,.5),(2,81,2)],[(0,74,4)]]
END=DUR-4.6; t=0.0; k=0
while t+bar<=END+0.01:
    i=k%8; r,ch=P[i]; lvl=min(1.0,0.55+0.09*k)          # 마디가 갈수록 커진다
    note(t,r,bar,0.78*lvl); note(t,r+12,bar,0.62*lvl); note(t+2*beat,r,beat*2,0.55*lvl); note(t+2*beat,r+12,beat*2,0.45*lvl)
    for j in range(8):
        if j in (0,4): continue
        note(t+j*beat/2+rng.normal(0,0.004),r+12+(7 if j%2 else 12),beat*0.9,(0.30 if j%2 else 0.36)*lvl)
    for b in range(4):
        v=(0.50 if b in (0,2) else 0.34)*lvl
        if k>=1:
            for m in ch: note(t+b*beat+rng.normal(0,0.004),m,beat*1.1,v)
    if k>=1:
        for (o,m,l) in MEL[i]:
            note(t+o*beat+0.008,m,beat*l+0.5,0.74*lvl)
            if k>=4: note(t+o*beat+0.008,m+12,beat*l+0.5,0.50*lvl)
    else:
        for j,m in enumerate([62,65,69,74]): note(t+j*beat,m,bar-j*beat,0.42)      # 첫 마디는 조용히 여는 분산화음
    t+=bar;k+=1
for j,m in enumerate([26,38,50,57,62,65,69,74,81]): note(t+j*0.03,m,6.0,0.85 if j<2 else 0.62)   # 마침: 넓게 쌓은 으뜸화음
L=int(SR*2.6); ir=rng.normal(0,1,L)*np.exp(-np.arange(L)/SR*2.4); ir[:int(SR*0.015)]=0; ir/=np.abs(ir).sum()/3.6
N=len(tot)+L; wet=np.stack([np.fft.irfft(np.fft.rfft(tot[:,c],N)*np.fft.rfft(np.roll(ir,c*41),N))[:len(tot)] for c in (0,1)],1)
mix=(tot+0.22*wet)[:int(SR*DUR)]; f=int(SR*2.2); mix[-f:]*=(np.linspace(1,0,f)**1.5)[:,None]; mix=np.tanh(mix/np.abs(mix).max()*1.6)/np.tanh(1.6)*0.92
with wave.open(OUT,"wb") as w: w.setnchannels(2); w.setsampwidth(2); w.setframerate(SR); w.writeframes((np.clip(mix,-1,1)*32767).astype("<i2").tobytes())
print("saved",OUT,"bars",k,"end chord at",round(t,2))
