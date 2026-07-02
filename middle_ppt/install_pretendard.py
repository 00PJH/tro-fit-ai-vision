"""
Pretendard 폰트 설치 스크립트 (Windows)
PPTX에서 Pretendard가 올바르게 표시되려면 Windows에 폰트가 설치되어 있어야 합니다.
관리자 권한으로 실행하거나, 사용자 폰트 디렉토리에 복사합니다.
"""

import os
import shutil
import ctypes
import sys

FONT_SRC = os.path.join(
    os.path.dirname(__file__),
    "Pretendard-1.3.9", "public", "static"
)

# 사용자 폰트 디렉토리 (관리자 권한 불필요)
USER_FONT_DIR = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")

FONTS_TO_INSTALL = [
    "Pretendard-Bold.otf",
    "Pretendard-SemiBold.otf",
    "Pretendard-Regular.otf",
    "Pretendard-Light.otf",
    "Pretendard-ExtraLight.otf",
    "Pretendard-Medium.otf",
]

def install_fonts():
    os.makedirs(USER_FONT_DIR, exist_ok=True)
    installed = []
    for fname in FONTS_TO_INSTALL:
        src = os.path.join(FONT_SRC, fname)
        dst = os.path.join(USER_FONT_DIR, fname)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            installed.append(fname)
            print(f"  [OK] {fname}")
        else:
            print(f"  [SKIP] 파일 없음: {src}")

    print(f"\n총 {len(installed)}개 폰트를 '{USER_FONT_DIR}'에 설치했습니다.")
    print("PowerPoint를 재시작하면 Pretendard 폰트가 표시됩니다.")

if __name__ == "__main__":
    print("Pretendard 폰트 설치 중...\n")
    install_fonts()
