#!/usr/bin/env python3
"""Claude 계정 스킬(SKILL.md)을 읽어 MOTIF 발행 스튜디오의 채널 프롬프트 JSON으로 만든다.
사용: python3 sync_skills.py OUT.json  → ArtifactData set collection "settings" doc_id "skills" file_path OUT.json
"""
import glob, json, os, re, sys, datetime, hashlib
MAP = {  # 앱 프롬프트 키 → (스킬 이름, 잘라낼 구간: 시작 제목 정규식, 끝 제목 정규식)
    'naver': ('motif-blog-writer', None, None),
    'kakao': ('motif-kakao', None, None),
    'threads': ('motif-threads', None, None),
    'instagram': ('motif-insta-master', None, None),
    'reels': ('motif-reels', r'^## ', r'^## .*(렌더|녹음|설치|키트|체크리스트|ffmpeg|OmniVoice)'),
    'tistory': ('motif-blog-autopublish', r'^### 티스토리', r'^### 워드프레스'),
    'wordpress': ('motif-blog-autopublish', r'^### 워드프레스', r'^## 3\.'),
}
COMMON = ('motif-blog-autopublish', r'^## 고정 규칙', r'^## 속도 규칙')
def find(name):
    for root in ['/root/.claude/skills', os.path.expanduser('~/.claude/skills'), '/mnt/skills']:
        hits = glob.glob(f'{root}/**/{name}/SKILL.md', recursive=True)
        if hits: return sorted(hits)[0]
    return None
def body(path):
    t = open(path, encoding='utf-8').read()
    return re.sub(r'^---\n.*?\n---\n', '', t, count=1, flags=re.S).strip()
def section(t, start, end):
    if not start: return t
    lines, out, on = t.split('\n'), [], False
    for ln in lines:
        if not on and re.match(start, ln): on = True
        elif on and end and re.match(end, ln): break
        if on: out.append(ln)
    return '\n'.join(out).strip() or t
res, src = {}, {}
for key, (name, s, e) in MAP.items():
    p = find(name)
    if not p: continue
    txt = section(body(p), s, e)
    if key in ('tistory', 'wordpress'):
        cp = find(COMMON[0]); txt = section(body(cp), COMMON[1], COMMON[2]) + '\n\n' + txt
    res[key] = txt[:24000]; src[key] = name
res['sources'] = src
res['hash'] = hashlib.md5(json.dumps({k: res[k] for k in src}, ensure_ascii=False).encode()).hexdigest()
res['updatedAt'] = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000Z')
json.dump(res, open(sys.argv[1] if len(sys.argv) > 1 else 'skills.json', 'w', encoding='utf-8'), ensure_ascii=False)
print({k: len(v) for k, v in res.items() if k in src}, res['hash'])
