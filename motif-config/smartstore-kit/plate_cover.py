#!/usr/bin/env python3
import cv2, numpy as np, argparse, sys
ap = argparse.ArgumentParser()
ap.add_argument("src"); ap.add_argument("dst")
ap.add_argument("--corners"); ap.add_argument("--logo"); ap.add_argument("--wm")
a = ap.parse_args()
img = cv2.imread(a.src)
if img is None: sys.exit("사진을 열 수 없음")
h, w = img.shape[:2]
def logo_from_file(p):
    im = cv2.imread(p, cv2.IMREAD_UNCHANGED)
    if im.ndim == 3 and im.shape[2] == 4 and im[..., 3].min() < 250:
        m = im[..., 3].astype(np.float32) / 255
    else:
        g = cv2.cvtColor(im[..., :3], cv2.COLOR_BGR2GRAY).astype(np.float32) / 255 if im.ndim == 3 else im.astype(np.float32) / 255
        m = 1 - g if g.mean() > 0.5 else g
    ys, xs = np.where(m > 0.4)
    return m[ys.min():ys.max() + 1, xs.min():xs.max() + 1]
def logo_from_watermark():
    if a.wm: x1, y1, x2, y2 = map(int, a.wm.split(","))
    else: x1, y1, x2, y2 = int(w * .02), int(h * .918), int(w * .147), int(h * .986)
    g = cv2.cvtColor(img[y1:y2, x1:x2], cv2.COLOR_BGR2GRAY).astype(np.float32)
    bg = cv2.GaussianBlur(g, (0, 0), 25)
    d = np.clip(np.abs(bg - g) / 70.0, 0, 1); d[d < 0.25] = 0
    ys, xs = np.where(d > 0.4)
    if len(xs) < 200: sys.exit("워터마크에서 로고를 찾지 못함: --logo 또는 --wm 지정 필요")
    d = d[max(ys.min() - 2, 0):ys.max() + 3, max(xs.min() - 2, 0):xs.max() + 3]
    S = 6
    big = cv2.GaussianBlur(cv2.resize(d, None, fx=S, fy=S, interpolation=cv2.INTER_CUBIC), (0, 0), S * 0.6)
    return np.clip((big - 0.45) * 8, 0, 1)
def detect():
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    quads = []
    for lo, hi in [((90, 60, 120), (115, 255, 255))]:
        m = cv2.morphologyEx(cv2.inRange(hsv, lo, hi), cv2.MORPH_CLOSE, np.ones((25, 25), np.uint8))
        for c in cv2.findContours(m, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)[0]:
            (cx, cy), (rw, rh), ang = cv2.minAreaRect(c)
            if rw < rh: rw, rh, ang = rh, rw, ang + 90
            if rh == 0 or cv2.contourArea(c) < 0.003 * w * h or not 1.8 < rw / rh < 6: continue
            if cv2.contourArea(c) / (rw * rh) < 0.75: continue
            b = sorted(cv2.boxPoints(((cx, cy), (rw * 1.05, rh * 1.03), ang)).tolist(), key=lambda p: p[0])
            L = sorted(b[:2], key=lambda p: p[1]); R = sorted(b[2:], key=lambda p: p[1])
            quads.append([L[0], R[0], R[1], L[1]])
    return quads
if a.corners:
    v = list(map(float, a.corners.split(",")))
    quads = [[v[i:i + 2] for i in range(j, j + 8, 2)] for j in range(0, len(v), 8)]
else:
    quads = detect()
if not quads: sys.exit("번호판 자동 검출 실패: 사진을 보고 --corners 로 네 모서리 지정")
logo = logo_from_file(a.logo) if a.logo else logo_from_watermark()
out = img.astype(np.float32)
for q in quads:
    q = np.float32(q)
    pw = (np.linalg.norm(q[1] - q[0]) + np.linalg.norm(q[2] - q[3])) / 2
    ph = (np.linalg.norm(q[3] - q[0]) + np.linalg.norm(q[2] - q[1])) / 2
    PH = 330; PW = int(PH * pw / ph)
    plate = np.zeros((PH, PW, 3), np.float32) + 240
    al = np.zeros((PH, PW), np.float32); r = 34
    cv2.rectangle(al, (r, 0), (PW - r, PH), 1, -1); cv2.rectangle(al, (0, r), (PW, PH - r), 1, -1)
    for c in [(r, r), (PW - r, r), (r, PH - r), (PW - r, PH - r)]: cv2.circle(al, c, r, 1, -1, cv2.LINE_AA)
    lh = int(PH * 0.42); lw = int(logo.shape[1] * lh / logo.shape[0])
    if lw > PW * 0.8: lw = int(PW * 0.8); lh = int(logo.shape[0] * lw / logo.shape[1])
    lg = cv2.resize(logo, (lw, lh), interpolation=cv2.INTER_AREA)[..., None]
    y0, x0 = (PH - lh) // 2, (PW - lw) // 2
    plate[y0:y0 + lh, x0:x0 + lw] = plate[y0:y0 + lh, x0:x0 + lw] * (1 - lg) + 20 * lg
    M = cv2.getPerspectiveTransform(np.float32([[0, 0], [PW, 0], [PW, PH], [0, PH]]), q)
    wp = cv2.warpPerspective(plate, M, (w, h), flags=cv2.INTER_AREA)
    wa = cv2.warpPerspective(al, M, (w, h), flags=cv2.INTER_AREA)[..., None]
    out = out * (1 - wa) + wp * wa
    print("가린 위치:", q.astype(int).tolist())
cv2.imwrite(a.dst, out.astype(np.uint8), [cv2.IMWRITE_JPEG_QUALITY, 95])
