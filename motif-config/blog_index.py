#!/usr/bin/env python3
"""모티프 블로그 링크 인덱스 만들기 · 갱신 (네이버 · 티스토리 · 워드프레스)

사용: python3 motif-config/blog_index.py            # 새 글만 받아서 갱신 (매일 아침 예약 작업이 실행)
      python3 motif-config/blog_index.py --full     # 모든 글 본문을 다시 받아 처음부터 다시 만들기

만드는 파일 (이 스크립트와 같은 폴더):
  blog-index.json                  전체 자료 (다음 실행이 이 파일과 비교해 새 글을 찾는다)
  naver-link-index.md              네이버 압축판: 규칙 · 허브 글 · 주제별 한 줄 목록 (글 쓸 때 읽는 파일)
  naver-link-index-full.md         네이버 상세판: 글마다 요약 · 소제목 · 차종 · 제품 · 주제
  tistory-wordpress-link-index.md  티스토리 · 워드프레스 목록 (같은 블로그끼리만 링크)

표준 출력 마지막 줄: RESULT {"new": [...], "removed": [...], "retitled": [...], "counts": {...}}
"""
import concurrent.futures as cf
import datetime
import html
import json
import os
import re
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
BLOG = '12rrrryn'
START = '2023-11-13'  # 이 날짜 이후 글만 (2026-11-13 기준 3년)
KST = datetime.timezone(datetime.timedelta(hours=9))
UA_M = 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 Mobile/15E148 Safari/604.1'
UA_D = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'

HUBS = [
    ('라인업 총정리', '224356903283'), ('솔라가드 등급제', '224350266461'), ('시공 전 5가지 질문', '224363316982'),
    ('LX 성분·구조', '224353637267'), ('TSER 한계·구조', '224359631357'), ('견적 차이=등급', '224372373858'),
    ('신차패키지 일정', '224373040446'), ('PPF 재시공 비용', '224350899809'), ('가죽코팅 시트별', '224369249895'),
]
RULES = """## 링크 삽입 규칙
1. 한 편당 2~5개. 오프닝에는 넣지 않고 본문 중간과 마무리 직전에 나눠 넣기
2. 문맥이 실제로 맞는 문장 뒤에만 넣기. 키워드만 겹친다고 넣지 않기
3. 형태를 섞어 쓰기
   A 권유형: 👉 함께 읽어보시면 좋은 글 + 다음 줄 제목 + 다음 줄 URL
   B 근거형: "자세한 수치는 아래 글에 정리해 두었습니다." + 다음 줄 URL
   C 사례형: "비슷한 고민으로 찾아주셨던 차주분 이야기입니다." + 다음 줄 URL
4. 같은 주제에서 2개 이상 고르지 않기 (사례 1 + 기술 1 + 가격·선택 1 조합이 기본)
5. 같은 차종 사례가 있으면 우선. 없으면 같은 세그먼트(전기차·수입 플래그십·국산 SUV)
6. 같은 주제면 최신 글 우선. 단 허브 글은 날짜와 관계없이 우선
7. URL은 https://blog.naver.com/12rrrryn/글번호 형식 그대로. 파라미터 붙이지 않기
8. 글을 다 쓴 뒤 코드블록 밖에 [삽입한 링크] 목록(제목·위치·형태) 따로 표시
9. 2025년 이전 글은 가격·행사·보증 내용이 지금과 다를 수 있으니 사례·생각을 잇는 용도로만 쓰기
"""


def curl(url, ua=UA_D, referer=None, tries=3):
    cmd = ['curl', '-s', '-L', '--max-time', '30', '-A', ua]
    if referer:
        cmd += ['-H', 'Referer: ' + referer]
    for i in range(tries):
        r = subprocess.run(cmd + [url], capture_output=True)
        if r.returncode == 0 and len(r.stdout) > 200:
            return r.stdout.decode('utf-8', 'ignore')
        time.sleep(2 + i * 3)
    return ''


def clean(t):
    t = re.sub(r'<(script|style)[^>]*>.*?</\1>', '', t, flags=re.S)
    t = re.sub(r'</p>|<br\s*/?>|</div>|</h\d>|</li>', '\n', t)
    t = html.unescape(html.unescape(re.sub(r'<[^>]+>', '', t))).replace('​', '').replace('\xa0', ' ')
    return re.sub(r'[ \t]+', ' ', re.sub(r'\n\s*\n+', '\n', t)).strip()


# ---------------- 분류 사전 ----------------
WORK = [('썬팅', r'썬팅|틴팅|선팅'), ('PPF', r'PPF|ppf|보호필름'), ('유리막코팅', r'유리막|세라믹\s?코팅|커버플릭'),
        ('가죽코팅', r'가죽\s?코팅|야와라'), ('신차패키지', r'신차\s?패키지'), ('랩핑', r'랩핑'), ('신차검수', r'신차\s?검수'),
        ('재시공·제거', r'재시공|필름 제거|제거 작업'), ('블랙박스·하이패스', r'블랙박스|하이패스'),
        ('생활보호 PPF', r'생활\s?보호|도어\s?엣지|도어\s?컵')]
PROD = [('솔라가드 LX', r'LX\d*|엘엑스'), ('제니스 IR', r'제니스'), ('퀀텀', r'퀀텀'), ('버텍스 1100', r'버텍스\s?1100'),
        ('버텍스 900', r'버텍스\s?900'), ('버텍스', r'버텍스(?!\s?(?:1100|900))'), ('티타늄', r'티타늄'), ('S라인', r'S\s?라인'),
        ('보그', r'보그'), ('G라인', r'G\s?라인|지라인'), ('엑스펠', r'엑스펠|XPEL|xpel'), ('얼티메이트 플러스', r'얼티메이트'),
        ('스텔스·무광', r'스텔스|무광'), ('익스트림', r'익스트림'), ('울트라핏', r'울트라핏'), ('컬러 PPF', r'컬러\s?PPF'),
        ('야와라', r'야와라'), ('커버플릭', r'커버플릭'), ('브이쿨', r'브이쿨|V-?KOOL'), ('루마', r'루마'), ('레이노', r'레이노'),
        ('후퍼옵틱', r'후퍼')]
CARS = ['모델Y', '모델 Y', '주니퍼', '모델3', '모델 3', '모델X', '모델S', '사이버트럭', 'GV60', 'GV70', 'GV80', 'G70', 'G80', 'G90',
        '아이오닉5', '아이오닉 5', '아이오닉6', '아이오닉 6', '아이오닉9', '아이오닉 9', 'EV3', 'EV4', 'EV5', 'EV6', 'EV9', '카니발',
        '쏘렌토', '싼타페', '팰리세이드', '그랜저', '아반떼', '쏘나타', 'K5', 'K8', 'K9', '스포티지', '투싼', '셀토스', '코나', '니로',
        '캐스퍼', '스타리아', '넥쏘', '토레스', '액티언', '콜레오스', '타스만', 'X3', 'X4', 'X5', 'X6', 'X7', 'iX', 'i4', 'i5', 'i7',
        '5시리즈', '3시리즈', '7시리즈', 'M2', 'M3', 'M4', 'M5', 'E클래스', 'S클래스', 'C클래스', 'GLC', 'GLE', 'GLS', 'G바겐', '지바겐',
        'EQE', 'EQS', '911', '타이칸', '카이엔', '마칸', '파나메라', '박스터', 'Q5', 'Q7', 'Q8', 'A6', 'A7', 'XC40', 'XC60', 'XC90',
        'EX30', 'S90', '렉서스', '프리우스', '알파드', '캠리', '디펜더', '레인지로버', '폴스타', '랭글러', '에스컬레이드', '머스탱',
        '브롱코', '우루스', '벤틀리', '마세라티', '페라리', '람보르기니', 'BYD', '콜로라도', '트랙스', '무쏘']
BR = [('테슬라', r'테슬라'), ('제네시스', r'제네시스|GV\d0|G[789]0'), ('현대', r'현대|아이오닉|팰리세이드|싼타페|그랜저'),
      ('기아', r'기아|EV\d|카니발|쏘렌토'), ('BMW', r'BMW|비엠'), ('벤츠', r'벤츠|메르세데스'), ('포르쉐', r'포르쉐'),
      ('아우디', r'아우디'), ('볼보', r'볼보'), ('랜드로버', r'랜드로버|디펜더|레인지로버')]
TOPIC = [('가격·비용', r'가격|비용|만 원|만원|견적'), ('농도 선택', r'농도|투과율|VLT'), ('열차단·TSER', r'열차단|TSER|IRER|적외선'),
         ('법규', r'법규|도로교통법|단속|과태료'), ('전파·하이패스', r'전파|하이패스|GPS|수신'), ('내구성·보증', r'내구성|보증|변색|수명'),
         ('시공 후 관리', r'건조|관리|세차|주의사항'), ('야간 시야', r'야간|밤 운전|시야'), ('전기차', r'전기차|전비|주행거리'),
         ('비교', r'비교|차이|vs|VS'), ('구조·기술', r'스퍼터링|구조|금속|세라믹|나노'), ('중고차·재시공', r'중고차|재시공|제거'),
         ('반사', r'반사율|반사 필름|비반사|무반사'), ('시공 품질', r'열성형|재단|먼지|들뜸|마감')]
IMPORT = r'BMW|비엠|벤츠|메르세데스|포르쉐|아우디|볼보|랜드로버|디펜더|레인지로버|렉서스|폴스타|미니|벤틀리|마세라티|페라리|람보르기니|캐딜락|에스컬레이드|지프|포드|토요타|BYD|(?<![A-Za-z0-9])[XxQ][3-8](?![A-Za-z0-9])|[357]시리즈|911|타이칸|카이엔|마칸|파나메라|[ESC]클래스|GL[CES]|XC\d0'


def tags(title, body, lst, th):
    first = [n for n, p in lst if re.search(p, title)]   # 제목에 나온 것이 먼저
    rest = [n for n, p in lst if n not in first and th and len(re.findall(p, body)) >= th]
    return first + rest


def cars(title, body):
    out = []
    for c in CARS:
        k = c.replace(' ', '')
        pat = r'(?<![A-Za-z0-9])' + re.escape(c) + r'(?![A-Za-z0-9])'
        hit = re.search(pat, title) or len(re.findall(pat, body)) >= 3
        if hit and k not in out:
            out.append(k)
    return out


def section(title, category=''):
    """한 글을 하나의 주제 칸에 넣는다 (제목 기준)."""
    t = title
    if category == '자동차 소식':
        return '자동차 소식'
    if category == '공지·오시는 길' or t.startswith('['):
        return '매장·철학·기타'
    if re.search(r'PPF|ppf|보호필름|랩핑|엑스펠|XPEL', t) or category == 'PPF':
        return 'PPF·랩핑'
    if re.search(r'유리막|가죽|커버플릭|야와라|세라믹\s?코팅', t):
        return '유리막·가죽코팅'
    if re.search(r'신차\s?패키지|신차\s?검수|출고', t):
        return '신차패키지·신차검수'
    if re.search(r'재시공|제거|관리|건조|세차|기포|들뜸|수명|변색|년 뒤|년 후|주의사항', t):
        return '썬팅-재시공·관리·내구성'
    if re.search(r'가격|비용|견적|등급|가성비|비싼|비싸|저렴|할인|이벤트|행사|값어치|돈 ?값|엔트리', t):
        return '썬팅-가격·등급'
    if re.search(r'농도|밝은|어둡|어두운|투과율|법|단속|야간|시야|블랙박스|하이패스|HUD', t):
        return '썬팅-농도·시야·법규'
    if re.search(r'TSER|TESR|IRER|수치|측정|열차단율|반사율|ISO|스펙|그래프|%', t):
        return '썬팅-수치·측정'
    if re.search(r'테슬라|모델\s?[Y3XS]|주니퍼|사이버트럭', t):
        return '시공사례-테슬라'
    if re.search(IMPORT, t):
        return '시공사례-수입차'
    if cars(t, ''):
        return '시공사례-국산차'
    if re.search(r'LX|엘엑스', t):
        return '썬팅-솔라가드 LX'
    if re.search(r'제니스|퀀텀|버텍스|티타늄|S\s?라인|보그|G\s?라인|브이쿨|루마|레이노|후퍼', t):
        return '썬팅-제니스 IR·퀀텀·다른 제품'
    if re.search(r'솔라가드|스퍼터링|열반사|세라믹|필름|전기차|열차단|썬팅|틴팅|햇살|적외선|자외선', t):
        return '썬팅-원리·일반'
    return '매장·철학·기타'


SECTIONS = ['썬팅-솔라가드 LX', '썬팅-제니스 IR·퀀텀·다른 제품', '썬팅-원리·일반', '썬팅-수치·측정', '썬팅-농도·시야·법규', '썬팅-가격·등급', '썬팅-재시공·관리·내구성',
            '신차패키지·신차검수', 'PPF·랩핑', '유리막·가죽코팅', '시공사례-테슬라', '시공사례-수입차', '시공사례-국산차',
            '매장·철학·기타', '자동차 소식']


def summarize(src, limit=170):
    src = re.sub(r'\s+', ' ', html.unescape(html.unescape(src))).replace('📌 빠른 답변', '').strip()
    out = ''
    for x in re.split(r'(?<=[.?!])\s+', src):
        if len(out) + len(x) > limit:
            break
        out += x + ' '
    return out.strip() or src[:limit]


def classify(title, txt, category=''):
    return dict(section=section(title, category), work=tags(title, txt, WORK, 4), product=tags(title, txt, PROD, 3),
                car=cars(title, txt), brand=tags(title, txt, BR, 0) or tags(title, txt, BR, 4),
                topic=tags(title, txt, TOPIC, 5))


def find_heads(lines):
    heads = []
    for l in lines:
        l = l.strip()
        if 6 < len(l) < 45 and re.search(r'(까\??|\?)$', l) and l not in heads:
            heads.append(l)
    return heads[:8]


# ---------------- 네이버 ----------------
def naver_list():
    """전체 글 목록 (START 이후). 실패하면 None."""
    items, cut = {}, datetime.datetime.strptime(START, '%Y-%m-%d').replace(tzinfo=KST).timestamp() * 1000
    for page in range(1, 60):
        s = curl(f'https://m.blog.naver.com/api/blogs/{BLOG}/post-list?categoryNo=0&itemCount=30&page={page}', UA_M,
                 f'https://m.blog.naver.com/{BLOG}')
        try:
            its = json.loads(s)['result']['items']
        except Exception:
            return None
        if not its:
            break
        for it in its:
            items[str(it['logNo'])] = it
        if min(it['addDate'] for it in its) < cut:
            break
        time.sleep(0.3)
    out = {}
    for n, it in items.items():
        if it['addDate'] < cut or it.get('notOpen') or it.get('allOpenPost') is False:
            continue
        if '마이 블로그 리포트' in (it.get('titleWithInspectMessage') or ''):
            continue
        out[n] = dict(title=html.unescape(it.get('titleWithInspectMessage') or ''), category=it.get('categoryName') or '',
                      date=datetime.datetime.fromtimestamp(it['addDate'] / 1000, KST).strftime('%Y-%m-%d'),
                      brief=it.get('briefContents') or '')
    return out


def naver_post(n, meta):
    s = curl(f'https://m.blog.naver.com/PostView.naver?blogId={BLOG}&logNo={n}', UA_M)
    i = s.find('se-main-container')
    body = s[i:] if i >= 0 else s
    comps = body.split('<div class="se-component ')
    b = ''.join(c for c in comps[1:] if c.startswith(('se-text', 'se-quotation', 'se-sectionTitle', 'se-table')))
    txt = clean(b)
    links = sorted(set(re.findall(BLOG + r'/(\d{9,})', body)) - {n})
    hs = re.findall(r'#([0-9A-Za-z가-힣]+)', txt[-700:])
    body2 = re.sub(r'#[0-9A-Za-z가-힣]+', '', txt)
    t = meta['title']
    brief = html.unescape(meta['brief'])
    if brief.startswith(t):
        brief = brief[len(t):]
    o = dict(ch='naver', url=f'https://blog.naver.com/{BLOG}/{n}', id=n, date=meta['date'], title=t,
             category=meta['category'], heads=find_heads(txt.split('\n')), summary=summarize(brief or body2[:400]),
             chars=len(txt), links=links, hashtags=hs[:12])
    o.update(classify(t, txt, meta['category']))
    return o


# ---------------- 티스토리 ----------------
def tistory_list():
    s = curl('https://motif-jeonju.tistory.com/sitemap.xml')
    ids = re.findall(r'<loc>https://motif-jeonju\.tistory\.com/(\d+)</loc>', s)
    return ids if s and ids else None


def tistory_post(n):
    s = curl(f'https://motif-jeonju.tistory.com/{n}')

    def g(p):
        m = re.search(p, s)
        return m.group(1) if m else ''
    t = html.unescape(g(r'<meta property="og:title" content="([^"]*)"'))
    if not t:
        return None
    desc = g(r'<meta property="og:description" content="([^"]*)"')
    date = g(r'<meta property="article:published_time" content="([^"]*)"')[:10]
    m = re.search(r'class="[^"]*(?:tt_article_useless_p_margin|contents_style)[^"]*"(.*?)(?:<div class="container_postbtn|<!-- System - START)', s, re.S)
    inner = m.group(1) if m else ''
    txt = clean(inner) if inner else html.unescape(desc)
    h2 = [clean(x) for x in re.findall(r'<h[23][^>]*>(.*?)</h[23]>', inner, re.S)]
    tg = [html.unescape(x) for x in re.findall(r'rel="tag">([^<]+)</a>', s)]
    wp = sorted(set(re.findall(r'motif-jeonju\.com/([a-z0-9][a-z0-9-]{3,})/?["\'<]', inner)))
    o = dict(ch='tistory', url=f'https://motif-jeonju.tistory.com/{n}', id=n, date=date, title=t, category='',
             heads=[h for h in h2 if 3 < len(h) < 60 and not h.startswith('Q.')][:8], summary=summarize(desc),
             chars=len(txt), links=wp, hashtags=tg[:12])
    o.update(classify(t, txt))
    return o


# ---------------- 워드프레스 ----------------
def wordpress_all():
    out = []
    for page in range(1, 20):
        s = curl(f'https://motif-jeonju.com/wp-json/wp/v2/posts?per_page=100&page={page}&_fields=id,date,slug,link,title,excerpt')
        try:
            ps = json.loads(s)
        except Exception:
            return None if page == 1 else out
        if not isinstance(ps, list):
            break
        for p in ps:
            t, ex = clean(p['title']['rendered']), clean(p['excerpt']['rendered'])
            o = dict(ch='wordpress', url=p['link'], id=str(p['id']), slug=p['slug'], date=p['date'][:10], title=t,
                     category='', heads=[], summary=summarize(ex), chars=0, links=[], hashtags=[])
            o.update(classify(t, ex))
            out.append(o)
        if len(ps) < 100:
            break
    return out or None


# ---------------- 출력 ----------------
def tagline(o):
    car = o['car'][:2] or o['brand'][:1]
    bits = car + o['product'][:2] + o['topic'][:2]
    seen, out = set(), []
    for b in bits:
        if b not in seen:
            seen.add(b)
            out.append(b)
    return ' · '.join(out)


def line(o):
    tg = tagline(o)
    return f"- {o['date'][2:]} | {o['title']} | {o['url']}" + (f" | {tg}" if tg else '')


def grouped(posts):
    out = []
    for sec in SECTIONS:
        ps = [o for o in posts if o['section'] == sec]
        if not ps:
            continue
        out.append(f'\n### {sec} ({len(ps)}편)')
        out += [line(o) for o in ps]
    return out


def write_all(data, today):
    nv, ts, wp = data['naver'], data['tistory'], data['wordpress']
    by = {o['id']: o for o in nv}
    L = [f'# 모티프 네이버 블로그 내부 링크 인덱스', '',
         f'범위: blog.naver.com/{BLOG} {nv[-1]["date"]} ~ {nv[0]["date"]} 발행 {len(nv)}편 · 마지막 갱신 {today} (매일 아침 자동 갱신)',
         '형식: 날짜 | 제목 | URL | 차종 · 제품 · 주제(본문 분석)',
         '주제 칸은 제목 기준 자동 분류입니다. 글을 고를 때 제목과 꼬리표로 문맥을 확인하세요.',
         '글의 요약 · 소제목까지 보려면 같은 폴더의 naver-link-index-full.md 에서 글번호로 찾으세요.', '', RULES,
         '## 허브 글 (우선 참조)']
    for name, n in HUBS:
        if n in by:
            L.append(f"- {name} | {by[n]['title']} | {by[n]['url']}")
    L += ['', '## 최근 30일 새 글 (아직 링크가 적은 글 · 문맥이 맞으면 먼저 검토)']
    lim = (datetime.datetime.strptime(today, '%Y-%m-%d') - datetime.timedelta(days=30)).strftime('%Y-%m-%d')
    L += [line(o) for o in nv if o['date'] >= lim]
    L += ['', '## 전체 글 인덱스 (주제별 · 최신순)'] + grouped(nv)
    open(os.path.join(HERE, 'naver-link-index.md'), 'w', encoding='utf-8').write('\n'.join(L) + '\n')

    F = ['# 모티프 네이버 블로그 링크 인덱스 상세판', '',
         f'{len(nv)}편 · 마지막 갱신 {today}. 요약은 글 도입부이고 소제목 · 차종 · 제품 · 주제는 본문에서 자동으로 뽑은 것입니다.', '']
    for o in nv:
        F.append(f"## {o['title']}")
        F.append(f"- 주소: {o['url']}")
        F.append(f"- 날짜: {o['date']} · 블로그 분류: {o['category']} · 주제 칸: {o['section']} · 본문 {o['chars']}자")
        for k, lab in (('work', '작업'), ('car', '차종'), ('brand', '브랜드'), ('product', '제품'), ('topic', '다루는 주제')):
            if o[k]:
                F.append(f"- {lab}: {', '.join(o[k])}")
        if o['heads']:
            F.append('- 소제목: ' + ' / '.join(o['heads']))
        F.append(f"- 도입부: {o['summary']}")
        if o['links']:
            F.append('- 이 글에 이미 넣은 링크: ' + ', '.join(o['links']))
        if o['hashtags']:
            F.append('- 해시태그: ' + ' '.join('#' + h for h in o['hashtags']))
        F.append('')
    open(os.path.join(HERE, 'naver-link-index-full.md'), 'w', encoding='utf-8').write('\n'.join(F))

    T = ['# 모티프 티스토리 · 워드프레스 내부 링크 인덱스', '',
         f'티스토리 {len(ts)}편 · 워드프레스 {len(wp)}편 · 마지막 갱신 {today} (매일 아침 자동 갱신)',
         '링크는 같은 블로그 글끼리만 넣습니다 (티스토리↔티스토리 · 워드프레스↔워드프레스). 짝 글끼리는 서로 링크합니다.',
         '형식: 날짜 | 제목 | URL | 차종 · 제품 · 주제', '',
         '## 티스토리 (motif-jeonju.tistory.com)'] + grouped(ts) + ['', '## 워드프레스 (motif-jeonju.com)'] + grouped(wp)
    open(os.path.join(HERE, 'tistory-wordpress-link-index.md'), 'w', encoding='utf-8').write('\n'.join(T) + '\n')
    json.dump(data, open(os.path.join(HERE, 'blog-index.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=0)


def main():
    full = '--full' in sys.argv
    today = datetime.datetime.now(KST).strftime('%Y-%m-%d')
    path = os.path.join(HERE, 'blog-index.json')
    old = json.load(open(path, encoding='utf-8')) if os.path.exists(path) and not full else {}
    res = dict(new=[], removed=[], retitled=[], warnings=[])
    data = {}

    # 네이버
    prev = {o['id']: o for o in old.get('naver', [])}
    lst = naver_list()
    if lst is None or (prev and len(lst) < len(prev) * 0.8):
        res['warnings'].append('네이버 목록을 받지 못해 네이버는 이전 내용 유지')
        data['naver'] = old.get('naver', [])
    else:
        need = [n for n in lst if n not in prev or prev[n]['title'] != lst[n]['title']]
        with cf.ThreadPoolExecutor(4) as ex:
            got = dict(zip(need, ex.map(lambda n: naver_post(n, lst[n]), need)))
        posts = []
        for n, meta in lst.items():
            if n in got:
                o = got[n]
                if o['chars'] < 100 and n in prev:      # 본문을 못 받았으면 이전 것 유지
                    o = prev[n]
                elif n not in prev:
                    res['new'].append(dict(ch='naver', date=o['date'], title=o['title'], url=o['url']))
                else:
                    res['retitled'].append(dict(ch='naver', title=o['title'], url=o['url']))
                if o['chars'] < 100:
                    res['warnings'].append(f"네이버 {n} 본문을 읽지 못함(제목만 반영)")
                posts.append(o)
            else:
                o = prev[n]
                o['date'], o['category'] = meta['date'], meta['category']
                o['section'] = section(o['title'], o['category'])
                posts.append(o)
        res['removed'] += [dict(ch='naver', title=o['title'], url=o['url']) for n, o in prev.items() if n not in lst]
        posts.sort(key=lambda o: (o['date'], o['id']), reverse=True)
        data['naver'] = posts

    # 티스토리
    prev = {o['id']: o for o in old.get('tistory', [])}
    ids = tistory_list()
    if ids is None or (prev and len(ids) < len(prev) * 0.8):
        res['warnings'].append('티스토리 목록을 받지 못해 티스토리는 이전 내용 유지')
        data['tistory'] = old.get('tistory', [])
    else:
        need = [n for n in ids if n not in prev]
        with cf.ThreadPoolExecutor(4) as ex:
            got = dict(zip(need, ex.map(tistory_post, need)))
        posts = []
        for n in ids:
            o = got.get(n) or prev.get(n)
            if not o:
                continue
            if n not in prev:
                res['new'].append(dict(ch='tistory', date=o['date'], title=o['title'], url=o['url']))
            posts.append(o)
        res['removed'] += [dict(ch='tistory', title=o['title'], url=o['url']) for n, o in prev.items() if n not in ids]
        posts.sort(key=lambda o: int(o['id']), reverse=True)
        data['tistory'] = posts

    # 워드프레스
    prev = {o['id']: o for o in old.get('wordpress', [])}
    wp = wordpress_all()
    if wp is None or (prev and len(wp) < len(prev) * 0.8):
        res['warnings'].append('워드프레스 목록을 받지 못해 워드프레스는 이전 내용 유지')
        data['wordpress'] = old.get('wordpress', [])
    else:
        now = {o['id'] for o in wp}
        res['new'] += [dict(ch='wordpress', date=o['date'], title=o['title'], url=o['url']) for o in wp if o['id'] not in prev]
        res['removed'] += [dict(ch='wordpress', title=o['title'], url=o['url']) for n, o in prev.items() if n not in now]
        wp.sort(key=lambda o: (o['date'], int(o['id'])), reverse=True)
        data['wordpress'] = wp

    if not data['naver']:
        print('RESULT ' + json.dumps(dict(error='네이버 글을 하나도 받지 못함 · 파일을 바꾸지 않음'), ensure_ascii=False))
        sys.exit(1)
    if full or not old:
        res['new'] = []  # 처음 만들 때는 전부 새 글이므로 목록을 비운다
    write_all(data, today)
    res['counts'] = {k: len(v) for k, v in data.items()}
    print('RESULT ' + json.dumps(res, ensure_ascii=False))


if __name__ == '__main__':
    main()
