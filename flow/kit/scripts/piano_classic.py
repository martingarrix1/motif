# 실제 피아노 샘플(음마다 녹음된 어쿠스틱 피아노)로 연주하는 클래식풍 소품 "유리의 결" (D장조 · 4/4 · 약 60bpm)
# 사용: piano_classic.py <길이초> <out.wav> <샘플폴더>
import numpy as np,sys,wave,subprocess,os
SR=44100; DUR=float(sys.argv[1]); OUT=sys.argv[2]; rng=np.random.default_rng(2211)
D=sys.argv[3] if len(sys.argv)>3 else "pn"; NM=["C","Cs","D","Ds","E","F","Fs","G","Gs","A","As","B"]; _c={}
def smp(m):
    if m not in _c:
        f=f"{D}/{NM[m%12]}{m//12-1}.mp3"
        raw=subprocess.run(["ffmpeg","-loglevel","error","-i",f,"-f","f32le","-ac","2","-ar",str(SR),"-"],stdout=subprocess.PIPE).stdout
        a=np.frombuffer(raw,np.float32).reshape(-1,2).astype(np.float64); pk=np.abs(a).max(); a=a/pk if pk>0 else a
        on=np.argmax(np.abs(a).max(1)>0.02); _c[m]=a[max(0,on-60):]
    return _c[m]
tot=np.zeros((int(SR*(DUR+8)),2))
def note(t0,m,dur,vel):
    a=smp(m); n=min(len(a),int(SR*(dur+0.35)))
    if n<=0: return
    s=a[:n].copy(); rel=int(SR*0.35); s[-rel:]*=np.linspace(1,0,rel)[:,None]**1.6
    # 여린 음은 배음을 덜어 부드럽게(해머가 약하게 칠 때의 음색)
    k=0.25+0.75*min(1,vel/0.9); y=np.empty_like(s); acc=np.zeros(2)
    if k<0.98:
        from scipy.signal import lfilter
        s=lfilter([k],[1,-(1-k)],s,axis=0)
    pan=np.clip((m-64)/60,-0.35,0.35); g=np.array([1-pan,1+pan])*vel
    i=int(max(0,t0)*SR); e=min(len(tot),i+n); tot[i:e]+=s[:e-i]*g
beat=60/60; bar=beat*4
P=[(38,[50,57,62,66]),(35,[47,54,59,62]),(31,[43,50,55,59]),(33,[45,52,57,61]),
   (42,[54,57,62,66]),(35,[47,54,59,62]),(40,[52,55,59,64]),(33,[45,52,55,61]),
   (38,[50,57,62,66]),(43,[55,59,62,67]),(40,[52,55,59,64]),(33,[45,52,57,61]),
   (35,[47,54,59,62]),(31,[43,50,55,59]),(33,[45,52,57,61]),(38,[50,57,62,66])]
M=[[(0,78,2),(2,81,1),(3,78,1)],[(0,76,1.5),(1.5,74,.5),(2,78,2)],[(0,74,3),(3,71,1)],[(0,73,3),(3,76,1)],
   [(0,81,1.5),(1.5,78,.5),(2,76,2)],[(0,74,2),(2,78,2)],[(0,79,1),(1,76,1),(2,71,2)],[(0,73,4)],
   [(0,78,1.5),(1.5,81,.5),(2,86,2)],[(0,83,2),(2,79,2)],[(0,78,1),(1,76,1),(2,71,2)],[(0,73,3),(3,76,1)],
   [(0,74,2),(2,71,2)],[(0,71,1),(1,74,1),(2,78,1),(3,74,1)],[(0,73,2),(2,69,1),(3,73,1)],[(0,74,4)]]
END=DUR-5.0          # 마지막 마침 화음이 놓이는 시각
t=0.4;k=0
while t+bar<END:
    i=k%16; b_,ch=P[i]; left=END-t; slow=1+0.18*max(0,1-(left-bar)/(2*bar))   # 끝으로 갈수록 살짝 늦춘다(리타르단도)
    bt=beat*slow; phr=0.85+0.15*np.sin(np.pi*((k%4)+0.5)/4)                 # 4마디 프레이즈의 셈여림
    note(t,b_,bt*4,0.50*phr)
    arp=[ch[0],ch[1],ch[2],ch[3],ch[2],ch[1],ch[2],ch[1]]
    for j,m in enumerate(arp):
        note(t+j*bt/2+rng.normal(0,0.006),m,bt*4-j*bt/2,(0.26 if j%2 else 0.32)*phr*(0.9 if j else 1.1))
    if k>=1:
        for (o,m,l) in M[i]:
            note(t+o*bt+rng.normal(0,0.007)+0.012,m,bt*l+0.9,(0.62+0.10*(m-74)/12+rng.normal(0,0.02))*phr)
            if k>=8 and o==0: note(t+0.02,m-12,bt*l+0.5,0.24*phr)
    t+=bt*4;k+=1
for j,m in enumerate([26,38,50,57,62,66,74]): note(t+j*0.045,m,6.0,0.46 if j<2 else 0.38)   # 마침: D장조 으뜸화음을 굴려서
L=int(SR*2.2); ir=rng.normal(0,1,L)*np.exp(-np.arange(L)/SR*3.0); ir[:int(SR*0.012)]=0; ir/=np.abs(ir).sum()/3.2
N=len(tot)+L; wet=np.stack([np.fft.irfft(np.fft.rfft(tot[:,c],N)*np.fft.rfft(np.roll(ir,c*37),N))[:len(tot)] for c in (0,1)],1)
mix=(tot+0.16*wet)[:int(SR*DUR)]
f=int(SR*2.5); mix[-f:]*=(np.linspace(1,0,f)**1.5)[:,None]; mix/=np.abs(mix).max()*1.1
with wave.open(OUT,"wb") as w: w.setnchannels(2); w.setsampwidth(2); w.setframerate(SR); w.writeframes((np.clip(mix,-1,1)*32767).astype("<i2").tobytes())
print("saved",OUT,"bars",k)
