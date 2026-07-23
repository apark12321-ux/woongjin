#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
hwp2hwpx.py — HWP 파일을 HWPX 파일로 변환

[동작 방식]
    설치된 한글(Hancom Office)을 win32com(OLE 자동화)으로 제어하여,
    한글 자체 저장 필터로 HWP -> HWPX 변환을 수행한다.
    한글이 직접 변환하므로 수식/이미지/서식 재현율이 가장 높다.

[전제 조건]
    - Windows
    - 한글(HWP) 2014 이상 설치 (HWPX 저장 필터 지원)
    - pip install pywin32

[사용 예]
    단일 파일          : python hwp2hwpx.py "input.hwp"
    출력 파일 지정      : python hwp2hwpx.py "input.hwp" -o "output.hwpx"
    폴더 일괄 변환      : python hwp2hwpx.py "D:\\Users\\jmvh\\문서\\웅진"
    하위 폴더까지 재귀   : python hwp2hwpx.py "D:\\...\\웅진" -r
    기존 파일 덮어쓰기   : python hwp2hwpx.py "input.hwp" --overwrite
    한글 창 보이기      : python hwp2hwpx.py "input.hwp" --visible
"""

import argparse
import sys
from pathlib import Path

try:
    import pythoncom
    import win32com.client as win32
except ImportError:
    print("[오류] pywin32가 필요합니다.  ->  pip install pywin32")
    sys.exit(1)


# 한글 SaveAs 저장 필터 이름 (한글 2014+ 기준). 두 번째 인자가 포맷명이다.
HWPX_FORMAT = "HWPX"


class HwpConverter:
    """설치된 한글을 자동화하여 HWP -> HWPX 변환을 수행하는 래퍼."""

    def __init__(self, visible: bool = False):
        pythoncom.CoInitialize()
        # HWPFrame.HwpObject : 한글 오토메이션 진입점
        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        self._register_security_module()
        try:
            self.hwp.XHwpWindows.Item(0).Visible = visible
        except Exception:
            pass  # 버전에 따라 창 객체 접근이 안 될 수 있음 — 무시

    def _register_security_module(self):
        """
        파일 열기/저장 시 뜨는 보안 경고 팝업을 자동 처리한다.
        (보안 모듈 DLL이 레지스트리에 등록돼 있어야 완전 무팝업.
         등록이 안 돼 있어도 변환은 되지만 팝업이 뜰 수 있음.
         팝업이 계속 성가시면 pyhwpx 라이브러리 사용을 권장 — 자동 등록해 줌.)
        """
        try:
            self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass

    def convert(self, src: Path, dst: Path) -> None:
        # 한글 API는 절대경로 문자열을 요구
        src_str = str(src.resolve())
        dst_str = str(dst.resolve())

        # 열기: forceopen 으로 잠김/충돌 무시하고 강제 오픈
        if not self.hwp.Open(src_str, "HWP", "forceopen:true"):
            raise RuntimeError(f"열기 실패: {src_str}")

        # 저장: HWPX 필터로 저장
        if not self.hwp.SaveAs(dst_str, HWPX_FORMAT, ""):
            raise RuntimeError(f"저장 실패: {dst_str}")

        # 현재 문서만 닫고 한글 프로세스는 유지 (다음 파일에서 재사용)
        self._close_current_doc()

    def _close_current_doc(self):
        try:
            self.hwp.XHwpDocuments.Item(0).Close(isDirty=False)
        except Exception:
            try:
                self.hwp.Clear(1)  # 1 = 저장하지 않고 버림
            except Exception:
                pass

    def quit(self):
        try:
            self.hwp.Quit()
        except Exception:
            pass
        finally:
            pythoncom.CoUninitialize()


def collect_targets(path: Path, recursive: bool) -> list[Path]:
    """입력 경로에서 변환 대상 .hwp 파일 목록을 수집."""
    if path.is_file():
        return [path] if path.suffix.lower() == ".hwp" else []
    pattern = "**/*.hwp" if recursive else "*.hwp"
    # 대소문자 무관하게 수집 (윈도우는 기본적으로 무관하지만 명시적으로 정리)
    return sorted(p for p in path.glob(pattern) if p.suffix.lower() == ".hwp")


def main():
    parser = argparse.ArgumentParser(
        description="HWP -> HWPX 변환기 (설치된 한글을 자동화)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", help="변환할 .hwp 파일 또는 폴더 경로")
    parser.add_argument("-o", "--output",
                        help="출력 .hwpx 경로 (단일 파일 변환 시에만 유효). "
                             "생략하면 원본과 같은 위치에 같은 이름으로 저장")
    parser.add_argument("-r", "--recursive", action="store_true",
                        help="폴더 입력 시 하위 폴더까지 재귀 탐색")
    parser.add_argument("--overwrite", action="store_true",
                        help="이미 존재하는 .hwpx도 다시 변환(덮어쓰기)")
    parser.add_argument("--visible", action="store_true",
                        help="변환 중 한글 창을 화면에 표시(디버깅용)")
    args = parser.parse_args()

    in_path = Path(args.input)
    if not in_path.exists():
        print(f"[오류] 경로를 찾을 수 없습니다: {in_path}")
        sys.exit(1)

    targets = collect_targets(in_path, args.recursive)
    if not targets:
        print("[알림] 변환할 .hwp 파일이 없습니다.")
        sys.exit(0)

    # 출력 경로 매핑 계산
    if args.output and in_path.is_file():
        out_map = {targets[0]: Path(args.output)}
    else:
        if args.output:
            print("[알림] -o 는 단일 파일 변환에만 적용됩니다. 폴더 입력이므로 무시합니다.")
        out_map = {src: src.with_suffix(".hwpx") for src in targets}

    print(f"[시작] 대상 {len(targets)}개 파일")
    converter = HwpConverter(visible=args.visible)

    ok, skipped, failed = 0, 0, 0
    try:
        for i, src in enumerate(targets, 1):
            dst = out_map[src]
            tag = f"({i}/{len(targets)})"

            if dst.exists() and not args.overwrite:
                print(f"  {tag} 건너뜀 (이미 존재): {dst.name}")
                skipped += 1
                continue

            try:
                converter.convert(src, dst)
                print(f"  {tag} 완료: {src.name}  ->  {dst.name}")
                ok += 1
            except Exception as e:
                print(f"  {tag} 실패: {src.name}  ({e})")
                failed += 1
    finally:
        converter.quit()

    print(f"[종료] 성공 {ok} · 건너뜀 {skipped} · 실패 {failed}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
