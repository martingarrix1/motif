#!/usr/bin/env python3
"""reel.json 하나로 MOTIF 릴스 mp4 만들기 (제미나이 채팅 → 깃허브 자동 렌더용)

사용: python3 run_reel.py <작업폴더>            # 폴더 안 reel.json
      python3 run_reel.py <작업폴더> --fake     # 녹음 없이 화면만 시험(무음)
순서: script.txt 생성 → make_reel.py voice(일레븐랩스 + 제미나이 검수) → spec.py 생성 → hfkit 렌더 → finish(피아노 · -16 LUFS) → 끝 검정 페이드
결과: <작업폴더>/<name>.mp4

reel.json 형식은 flow/gemini/GEM_지침.md 참고.
"""
import json, os, re, subprocess, sys

MK = os.path.expanduser(os.environ.get("MOTIF_MK", "~/mk"))
TK = os.path.join(MK, "textkit")
HF = os.path.join(TK, "hfkit")
SEX = re.compile(r"^(s(\d+))?([+-]\d+(?:\.\d+)?)?$")


def sh(c, **k):
    print("$", " ".join(c)); subprocess.run(c, check=True, **k)


def dur(f):
    return float(subprocess.check_output(["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "csv=p=0", f]))


def wrap(text, n=13):
    """자막 자동 줄바꿈: 낱말 단위로 한 줄 n자 안팎"""
    rows, cur = [], ""
    for w in text.split():
        if cur and len(cur) + 1 + len(w) > n: rows.append(cur); cur = w
        else: cur = (cur + " " + w).strip()
    if cur: rows.append(cur)
    return "|".join(rows)


def cap_rows(s):
    """'신차 패키지는 [[품목이 많을수록]]|((오히려)) 손해' → caption rows.  | 줄바꿈 · [[ ]] 그라데이션 강조 · (( )) 회색"""
    rows = []
    for line in s.split("|"):
        segs, pos = [], 0
        for m in re.finditer(r"\[\[(.+?)\]\]|\(\((.+?)\)\)", line):
            pre = line[pos:m.start()]
            if pre.strip(): segs.append([pre.strip(), None])
            segs.append([(m.group(1) or m.group(2)).strip(), "b" if m.group(1) else "g"])
            if pre and not pre.endswith(" ") and segs[-2:-1]: segs[-1][0] = "~" + segs[-1][0]
            pos = m.end()
        rest = line[pos:]
        if rest.strip():
            segs.append([("~" if not rest.startswith(" ") and segs else "") + rest.strip(), None])
        rows.append(segs)
    return rows


def resolve(v, sc0, sstarts):
    """시간 표기: 숫자 = 장면 시작 기준 초 / "s2+0.3" = 이 장면 2번째 문장 시작 + 0.3초 / "+1.2" = 장면 시작 + 1.2초"""
    if isinstance(v, str):
        m = SEX.match(v.strip())
        if m and (m.group(1) or m.group(3)):
            base = sstarts[int(m.group(2)) - 1] if m.group(1) else sc0
            return round(base + float(m.group(3) or 0), 3)
        return v
    if isinstance(v, list): return [resolve(x, sc0, sstarts) for x in v]
    if isinstance(v, dict): return {k: resolve(x, sc0, sstarts) for k, x in v.items()}
    return v


TIME_KEYS = {"at", "fix_at", "done_at"}


def main():
    D = os.path.abspath(sys.argv[1]); fake = "--fake" in sys.argv
    R = json.load(open(os.path.join(D, "reel.json"), encoding="utf-8"))
    name = R.get("name") or os.path.basename(D); mode = R.get("mode", "thelx")
    scenes = R["scenes"]
    open(os.path.join(D, "script.txt"), "w", encoding="utf-8").write("\n\n".join("\n".join(s["lines"]) for s in scenes) + "\n")

    # 1) 녹음
    if fake:
        st, t = [], 0.0
        for k, s in enumerate(scenes):
            for j, l in enumerate(s["lines"]):
                st.append(round(t, 3)); t += 0.2 * len(l.replace(" ", "")) + (0.75 if j == len(s["lines"]) - 1 else 0.4)
        sh(["ffmpeg", "-y", "-loglevel", "error", "-f", "lavfi", "-i", "anullsrc=r=44100:cl=mono", "-t", str(t), os.path.join(D, "voice.wav")])
        json.dump({"sent_starts": st, "voice_len": t}, open(os.path.join(D, "timing.json"), "w"))
    else:
        env = dict(os.environ)
        if R.get("voice") == "f":
            env["MOTIF_VOICE_ID"] = "m7Ra9MY2nOIpz2AQPg0z"
            env["MOTIF_VOICE_SETTINGS"] = json.dumps({"stability": 0.6, "similarity_boost": 1.0, "style": 0, "use_speaker_boost": True, "speed": 1.08})
        sh(["python3", os.path.join(TK, "make_reel.py"), "voice", D], env=env)
    T = json.load(open(os.path.join(D, "timing.json")))
    ss, VL = T["sent_starts"], T["voice_len"]

    # 2) spec
    tail = 5.5 if mode == "motif" else 4.0
    DUR = round(VL + tail, 2)
    items, k = [], 0
    bounds = []
    for s in scenes:
        n = len(s["lines"]); bounds.append(ss[k:k + n]); k += n
    sign = mode == "motif" and re.search("모티프.*였습니다", scenes[-1]["lines"][-1])
    end_t = round(bounds[-1][-1] - 0.85, 3) if sign else DUR
    for i, s in enumerate(scenes):
        st = bounds[i]; t0 = max(0.0, st[0] - 0.12) if i else 0.15
        t1 = (bounds[i + 1][0] - 0.05) if i + 1 < len(scenes) else (end_t if sign else DUR - 1.2)
        g = s.get("graphic")
        if g:
            g = dict(g); p = g.pop("p")
            gd = {kk: (resolve(vv, st[0], st) if kk in TIME_KEYS or isinstance(vv, (list, dict, str)) else vv) for kk, vv in g.items()}
            for kk in TIME_KEYS:
                if kk in g and isinstance(g[kk], (int, float)): gd[kk] = round(st[0] + g[kk], 3)
            items.append(dict(p=p, t=round(t0, 3), d=round(t1 - t0, 3), **gd))
        if s.get("bg"): items.append(dict(p="bgword", t=round(t0, 3), d=round(t1 - t0, 3), word=s["bg"]))
        caps = s.get("captions") or [wrap(l) for l in s["lines"]]
        y = 1000
        for j, c in enumerate(caps[:len(st)]):
            if sign and i == len(scenes) - 1 and j == len(st) - 1: continue   # 마무리 멘트는 엔드카드가 대신
            rows = cap_rows(c); ct = max(t0, st[j] - 0.1)
            items.append(dict(p="caption", t=round(ct, 3), d=round(t1 - ct, 3), rows=rows, y=y, size=60)); y += 84 * len(rows) + 26
    hd = end_t if sign else DUR - 1.2
    if R.get("header"): items.insert(0, dict(p="header", t=0.0, d=round(hd, 3), title=R["header"]))
    if mode in ("thelx", "motif"): items.append(dict(p="logo", t=0.0, d=round(hd, 3), name="thelx" if mode == "thelx" else "motif", w=180 if mode == "thelx" else 140))
    if sign: items.append(dict(p="endcard", t=end_t, d=round(DUR - end_t, 3), logo_="motif"))
    files = [os.path.join(D, f) for f in R.get("photos", []) if os.path.exists(os.path.join(D, f))]
    spec = dict(dur=DUR, dir=f"/tmp/hf/{name}", files=files, items=items)
    open(os.path.join(D, "spec.py"), "w", encoding="utf-8").write("SPEC = " + json.dumps(spec, ensure_ascii=False, indent=1).replace("null", "None").replace("true", "True").replace("false", "False") + "\n")

    # 3) 렌더 → 소리 → 검정 페이드
    vid = os.path.join(D, "video.mp4"); mix = os.path.join(D, "mix.mp4"); out = os.path.join(D, f"{name}.mp4")
    sh(["python3", os.path.join(HF, "hfkit.py"), os.path.join(D, "spec.py"), vid] + (["--draft"] if "--draft" in sys.argv else []), cwd=HF)
    sh(["python3", os.path.join(HF, "finish.py"), "reel", vid, os.path.join(D, "voice.wav"), mix, R.get("piano", "calm")], cwd=HF)
    L = dur(mix)
    sh(["ffmpeg", "-y", "-loglevel", "error", "-i", mix, "-vf", f"fade=t=out:st={L - 1.5:.2f}:d=1.5", "-af", f"afade=t=out:st={L - 2.5:.2f}:d=2.5",
        "-c:v", "libx264", "-crf", "18", "-preset", "veryfast", "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "256k", "-movflags", "+faststart", out])
    print("완성:", out, f"{L:.1f}초")


if __name__ == "__main__":
    main()
