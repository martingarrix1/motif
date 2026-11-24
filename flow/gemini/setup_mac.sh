#!/bin/bash
# MOTIF 릴스 · 제미나이 CLI 작업장 설치 (아이맥 터미널에 한 줄로 실행)
#   curl -sSL https://raw.githubusercontent.com/martingarrix1/motif/smartstore-assets/flow/gemini/setup_mac.sh | bash
# 다시 실행해도 안전하다(이미 있는 것은 건너뛰고 키트는 최신으로 바꾼다).
set -e
RAW=https://raw.githubusercontent.com/martingarrix1/motif/smartstore-assets
say(){ printf "\n▶ %s\n" "$1"; }

say "1/6 홈브루 확인"
if ! command -v brew >/dev/null; then
  echo "홈브루가 없습니다. https://brew.sh 의 설치 명령을 먼저 실행한 뒤 이 명령을 다시 실행하세요."; exit 1; fi

say "2/6 ffmpeg · Node 22 · 파이썬 도구"
command -v ffmpeg >/dev/null || brew install ffmpeg
NODE_MAJOR=$(node -v 2>/dev/null | sed 's/v\([0-9]*\).*/\1/' || echo 0)
if [ "${NODE_MAJOR:-0}" -lt 22 ]; then
  brew install node@22
  NB="$(brew --prefix node@22)/bin"
  grep -q "node@22" ~/.zshrc 2>/dev/null || echo "export PATH=\"$NB:\$PATH\"" >> ~/.zshrc
  export PATH="$NB:$PATH"
fi
python3 -m pip install -q --user fonttools brotli numpy 2>/dev/null || python3 -m pip install -q --user --break-system-packages fonttools brotli numpy

say "3/6 제미나이 CLI"
command -v gemini >/dev/null || npm install -g @google/gemini-cli

say "4/6 MOTIF 키트(최신)"
curl -sSL $RAW/flow/kit/bootstrap.sh | sh -s ~/mk
mkdir -p ~/mk/fonts
for W in Regular Bold Black; do
  [ -f ~/mk/fonts/NotoSansCJKkr-$W.otf ] || curl -sSL -o ~/mk/fonts/NotoSansCJKkr-$W.otf "https://github.com/notofonts/noto-cjk/raw/main/Sans/OTF/Korean/NotoSansCJKkr-$W.otf"
done

say "5/6 API 키(구글 드라이브의 MOTIF_API_KEYS.env에서)"
KEYF=$(find "$HOME/Library/CloudStorage" -maxdepth 4 -name "MOTIF_API_KEYS.env" 2>/dev/null | head -n 1)
if [ -n "$KEYF" ]; then
  mkdir -p ~/.gemini ~/.eleven; umask 077
  G=$(grep -E '^GEMINI_API_KEY' "$KEYF" | head -n1 | cut -d= -f2- | tr -d ' \r"')
  L=$(grep -E '^ELEVENLABS_API_KEY' "$KEYF" | head -n1 | cut -d= -f2- | tr -d ' \r"')
  [ -n "$G" ] && printf '%s' "$G" > ~/.gemini/key
  [ -n "$L" ] && printf '%s' "$L" > ~/.eleven/key
  echo "키를 넣었습니다(값은 화면에 표시하지 않음)."
else
  echo "MOTIF_API_KEYS.env를 찾지 못했습니다. 구글 드라이브 데스크톱 앱이 켜져 있는지 확인하세요."
fi

say "6/6 작업장 ~/motif-reels"
mkdir -p ~/motif-reels
curl -sSL -o ~/motif-reels/GEMINI.md $RAW/flow/gemini/GEMINI.md
echo
echo "완료. 앞으로는 터미널에서 아래 두 줄이면 됩니다."
echo "  cd ~/motif-reels"
echo "  gemini"
echo "처음 실행하면 구글 계정 로그인 창이 뜹니다(12rrrryn 계정으로)."
