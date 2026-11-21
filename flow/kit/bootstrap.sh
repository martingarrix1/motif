#!/bin/sh
# MOTIF 영상 작업 준비 한 줄: curl -sSL https://raw.githubusercontent.com/martingarrix1/motif/smartstore-assets/flow/kit/bootstrap.sh | sh -s ~/mk
# 결과: <폴더>/scripts (편집 · 피아노 코드), fonts, brand (로고 원본), pn (피아노 샘플 84개), textkit (텍스트 버전 릴스 한 줄 제작 · 클래식 피아노 20곡 · 깜비 캐릭터)
D=${1:-$HOME/mk}; R=https://raw.githubusercontent.com/martingarrix1/motif/smartstore-assets/flow
mkdir -p "$D/scripts" "$D/fonts" "$D/brand" "$D/pn" && cd "$D" || exit 1
{
 for f in red_moving tint_moving piano_grand piano_classic; do echo "$R/kit/scripts/$f.py scripts/$f.py"; done
 for f in InterDisplay-Bold.otf InterDisplay-Medium.otf NotoSansCJK-Medium.ttc Poppins-Bold.ttf Poppins-Medium.ttf; do echo "$R/kit/fonts/$f fonts/$f"; done
 for f in motif_logo_dark motif_logo_light thelx_ink_dark thelx_ink_light lxsg_ink_dark lxsg_ink_light; do echo "$R/brand/$f.png brand/$f.png"; done
 for n in C Cs D Ds E F Fs G Gs A As B; do for o in 1 2 3 4 5 6 7; do echo "https://raw.githubusercontent.com/nbrosowsky/tonejs-instruments/master/samples/piano/$n$o.mp3 pn/$n$o.mp3"; done; done
} | while read u p; do [ -s "$p" ] || echo "$u $p"; done | xargs -P 16 -n 2 sh -c 'curl -sSL --retry 2 -o "$1" "$0"'
python3 -c "import numpy,PIL,scipy,cv2" 2>/dev/null || pip install -q --break-system-packages numpy pillow scipy opencv-python-headless 2>/dev/null || pip install -q numpy pillow scipy opencv-python-headless
# 텍스트 버전 릴스 키트(textkit): 명령 한 줄 제작 — python3 textkit/textv/make_text.py all <폴더>
curl -sSL --retry 2 -o textkit.tgz "$R/kit/textkit.tgz" && rm -rf textkit && tar xzf textkit.tgz && rm -f textkit.tgz
[ -f /usr/share/fonts/opentype/noto/NotoSansCJK-Black.ttc ] || (apt-get install -y -q fonts-noto-cjk >/dev/null 2>&1 || sudo apt-get install -y -q fonts-noto-cjk >/dev/null 2>&1)
echo "준비 완료 $D : textkit $( [ -d textkit ] && echo OK || echo 없음)"
echo "준비 완료 $D : scripts $(ls scripts|wc -l) · fonts $(ls fonts|wc -l) · brand $(ls brand|wc -l) · pn $(ls pn|wc -l)"
