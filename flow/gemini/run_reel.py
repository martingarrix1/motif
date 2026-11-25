#!/usr/bin/env python3
"""reel.json 하나로 MOTIF 릴스 mp4 생성 (기본형 + editorial + onetake + fx 겸용)"""
import json, os, re, subprocess, sys

MK = os.path.expanduser(os.environ.get("MOTIF_MK", "~/mk"))
TK = os.path.join(MK, "textkit")
HF = os.path.join(TK, "hfkit")
SEX = re.compile(r"^(s(\d+))?([+-]\d+(?:\.\d+)?)?$")
TIME_KEYS = {"at", "fix_at", "done_at", "collapse_at"}


def sh(c, **k):
    print("$", " ".join(c)); subprocess.run(c, check=True, **k)


def dur(f):
    return float(subprocess.check_output(["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "csv=p=0", f]))


def wrap(text, n=12):
    rows, cur = [], ""
    for w in text.split():
        if cur and len(cur) + 1 + len(w) > n: rows.append(cur); cur = w
        else: cur = (cur + " " + w).strip()
    if cur: rows.append(cur)
    return "|".join(rows)


def cap_rows(s):
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
    if isinstance(v, str):
        m = SEX.match(v.strip())
        if m and (m.group(1) or m.group(3)):
            base = sstarts[int(m.group(2)) - 1] if m.group(1) else sc0
            return round(base + float(m.group(3) or 0), 3)
        return v
    if isinstance(v, list): return [resolve(x, sc0, sstarts) for x in v]
    if isinstance(v, dict): return {k: resolve(x, sc0, sstarts) for k, x in v.items()}
    return v


def main():
    D = os.path.abspath(sys.argv[1]); fake = "--fake" in sys.argv
    R = json.load(open(os.path.join(D, "reel.json"), encoding="utf-8"))
    name = R.get("name") or os.path.basename(D)
    mode = R.get("mode", "thelx")
    style = R.get("style", "")
    scenes = R["scenes"]
    open(os.path.join(D, "script.txt"), "w", encoding="utf-8").write("\n\n".join("\n".join(s["lines"]) for s in scenes) + "\n")

    if fake:
        st, t = [], 0.0
        for s in scenes:
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

    tail = 5.5 if mode == "motif" else 4.2
    DUR = round(VL + tail, 2)
    bounds, k = [], 0
    for s in scenes:
        n = len(s["lines"]); bounds.append(ss[k:k + n]); k += n
    motif_sign = bool(mode == "motif" and re.search("모티프.*였습니다", scenes[-1]["lines"][-1]))
    qend = None if motif_sign else R.get("end")
    end_t = round(bounds[-1][-1] - 0.85, 3) if motif_sign else (round(bounds[-1][-1] - 0.3, 3) if qend else DUR)
    sign = motif_sign or bool(qend)
    hd = end_t if sign else DUR
    files = [os.path.join(D, f) for f in R.get("photos", []) if os.path.exists(os.path.join(D, f))]

    if style == "onetake":
        stations, hud_items = [], []
        if R.get("header"): hud_items.append(dict(p="header", t=0.0, d=round(hd, 3), title=R["header"]))
        if mode in ("thelx", "motif"): hud_items.append(dict(p="logo", t=0.0, d=round(hd, 3), name="thelx" if mode == "thelx" else "motif", w=180 if mode == "thelx" else 140))
        for i, s in enumerate(scenes):
            st = bounds[i]
            cam_t = max(0.0, st[0] - 0.88) if i else 0.0
            act_t = max(0.0, st[0] - 0.12) if i else 0.35
            t1 = (bounds[i + 1][0] + 0.02) if i + 1 < len(scenes) else (end_t if sign else DUR)
            stay_d = round(DUR - act_t, 3)
            g_items, c_items, sbox = [], [], None

            g = s.get("graphic")
            if g:
                g = dict(g); p = g.pop("p"); sbox = g.get("box", [84, 300, 912, 600])
                gd = {kk: (resolve(vv, st[0], st) if kk in TIME_KEYS or isinstance(vv, (list, dict, str)) else vv) for kk, vv in g.items()}
                for kk in TIME_KEYS:
                    if kk in g and isinstance(g[kk], (int, float)): gd[kk] = round(st[0] + g[kk], 3)
                g_items.append(dict(p=p, t=round(act_t, 3), d=stay_d, **gd))

            fxs = s.get("fx") or []
            if isinstance(fxs, dict): fxs = [fxs]
            for fx in fxs:
                fx = dict(fx); fp = fx.pop("p")
                fxd = {kk: (resolve(vv, st[0], st) if kk in TIME_KEYS or isinstance(vv, (list, dict, str)) else vv) for kk, vv in fx.items()}
                c_items.append(dict(p=fp, t=round(act_t, 3), d=round(max(1.0, t1 - act_t), 3), **fxd))

            caps = s.get("captions") or [wrap(l) for l in s["lines"]]
            vis_n = (len(st) - 1) if (sign and i == len(scenes) - 1) else len(st)
            cy_cap = 1020
            for j, c in enumerate(caps[:vis_n]):
                rows = cap_rows(c); ct = max(act_t, st[j] - 0.10)
                all_sub = all(all(st_ == "g" for _, st_ in r) for r in rows)
                c_items.append(dict(p="caption", t=round(ct, 3), d=round(DUR - ct, 3), rows=rows, y=cy_cap, size=58))
                cy_cap += (66 if all_sub else 78) * len(rows) + 20
            stations.append(dict(cam_t=round(cam_t, 3), shell_box=sbox, g_items=g_items, c_items=c_items))

        if motif_sign:
            ov_start = max(0.2, round(end_t - 0.35, 3))
            ec_start = round(ov_start + 0.85, 3)
            stations.append(dict(cam_t=ov_start, is_end=True, overview=True, g_items=[], c_items=[dict(p="endcard", t=ec_start, d=round(DUR - ec_start, 3), logo_="motif")]))
        elif qend:
            q = (qend.get("question") or [])[:3]; q += [""] * (3 - len(q))
            end_c = [dict(p="centertext", t=end_t, d=round(DUR - end_t, 3), y=620, size=46, lines=[[q[0], "g"], [q[1], "k"], [q[2], "b"]])]
            if qend.get("product"): end_c.append(dict(p="centertext", t=round(end_t + 0.5, 3), d=round(DUR - end_t - 0.5, 3), y=1080, size=34, lines=[[qend["product"], "g"]]))
            stations.append(dict(cam_t=max(0.2, round(end_t - 0.55, 3)), is_end=True, overview=False, g_items=[], c_items=end_c))

        spec = dict(dur=DUR, style="onetake", dir=f"/tmp/hf/{name}", files=files, stations=stations, hud_items=hud_items)
    else:
        items = []
        editorial = (style == "editorial")
        for i, s in enumerate(scenes):
            st = bounds[i]; t0 = max(0.0, st[0] - 0.14) if i else 0.40
            t1 = (bounds[i + 1][0] + 0.02) if i + 1 < len(scenes) else (end_t if sign else DUR)
            caps = s.get("captions") or [wrap(l) for l in s["lines"]]
            vis_n = (len(st) - 1) if (sign and i == len(scenes) - 1) else len(st)
            g = s.get("graphic")
            is_ed = editorial or (g and str(g.get("p", "")).startswith("ed_"))

            if g:
                g = dict(g); p = g.pop("p")
                gd = {kk: (resolve(vv, st[0], st) if kk in TIME_KEYS or isinstance(vv, (list, dict, str)) else vv) for kk, vv in g.items()}
                for kk in TIME_KEYS:
                    if kk in g and isinstance(g[kk], (int, float)): gd[kk] = round(st[0] + g[kk], 3)
                if is_ed:
                    gd["caps"] = [cap_rows(c) for c in caps[:vis_n]]
                    gd["stimes"] = [round(max(t0, x - 0.10), 3) for x in st[:vis_n]]
                items.append(dict(p=p, t=round(t0, 3), d=round(t1 - t0, 3), **gd))

            fxs = s.get("fx") or []
            if isinstance(fxs, dict): fxs = [fxs]
            for fx in fxs:
                fx = dict(fx); fp = fx.pop("p")
                fxd = {kk: (resolve(vv, st[0], st) if kk in TIME_KEYS or isinstance(vv, (list, dict, str)) else vv) for kk, vv in fx.items()}
                items.append(dict(p=fp, t=round(t0, 3), d=round(t1 - t0, 3), **fxd))

            if not is_ed:
                if s.get("bg"): items.append(dict(p="bgword", t=round(t0, 3), d=round(t1 - t0, 3), word=s["bg"]))
                y = 1020
                for j, c in enumerate(caps[:vis_n]):
                    rows = cap_rows(c); ct = max(t0, st[j] - 0.12)
                    all_sub = all(all(st_ == "g" for _, st_ in r) for r in rows)
                    items.append(dict(p="caption", t=round(ct, 3), d=round(t1 - ct, 3), rows=rows, y=y, size=58))
                    y += (66 if all_sub else 78) * len(rows) + 20

        if R.get("header") and not editorial: items.insert(0, dict(p="header", t=0.0, d=round(hd, 3), title=R["header"]))
        if mode in ("thelx", "motif"): items.append(dict(p="logo", t=0.0, d=round(hd, 3), name="thelx" if mode == "thelx" else "motif", w=180 if mode == "thelx" else 140))
        if motif_sign: items.append(dict(p="endcard", t=end_t, d=round(DUR - end_t, 3), logo_="motif"))
        if qend:
            q = (qend.get("question") or [])[:3]; q += [""] * (3 - len(q))
            items.append(dict(p="centertext", t=end_t, d=round(DUR - end_t, 3), y=620, size=46, lines=[[q[0], "g"], [q[1], "k"], [q[2], "b"]]))
            if qend.get("product"): items.append(dict(p="centertext", t=round(end_t + 0.5, 3), d=round(DUR - end_t - 0.5, 3), y=1080, size=34, lines=[[qend["product"], "g"]]))
        spec = dict(dur=DUR, style=style, zoomcut=bool(R.get("zoomcut", False)), dir=f"/tmp/hf/{name}", files=files, items=items)

    open(os.path.join(D, "spec.py"), "w", encoding="utf-8").write("SPEC = " + json.dumps(spec, ensure_ascii=False, indent=1).replace("null", "None").replace("true", "True").replace("false", "False") + "\n")

    vid = os.path.join(D, "video.mp4"); mix = os.path.join(D, "mix.mp4"); out = os.path.join(D, f"{name}.mp4")
    sh(["python3", os.path.join(HF, "hfkit.py"), os.path.join(D, "spec.py"), vid] + (["--draft"] if "--draft" in sys.argv else []), cwd=HF)
    sh(["python3", os.path.join(HF, "finish.py"), "reel", vid, os.path.join(D, "voice.wav"), mix, R.get("piano", "calm")], cwd=HF)
    L = dur(mix)
    sh(["ffmpeg", "-y", "-loglevel", "error", "-i", mix,
        "-vf", f"fade=t=out:st={max(0, L - 2.0):.2f}:d=2.0:color=black",
        "-af", f"afade=t=out:st={max(0, L - 4.0):.2f}:d=4.0",
        "-c:v", "libx264", "-crf", "18", "-preset", "veryfast", "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "256k", "-movflags", "+faststart", out])
    print("완성:", out, f"{L:.1f}초")


if __name__ == "__main__":
    main()
