# -*- coding: utf-8 -*-
r"""
수식_자동재계산.py  (한컴오피스 한글 2014 대응)
==============================================================
이 스크립트 하나가 두 가지를 동시에 해결합니다:
  (1) 한글 2014 버전 파일(.hwp)로 저장
  (2) 모든 수식의 간격을 한글이 직접 재계산 → 간격 완벽

--------------------------------------------------------------
[실행 조건]  한컴오피스 한글 2014 이상이 설치된 Windows PC
[설치] 명령 프롬프트(cmd):  pip install pywin32

[사용법 — 3가지 중 편한 것]
  (A) 변환할 hwpx 파일을 이 .py 아이콘 위로 드래그&드롭   ← 가장 쉬움
  (B) cmd 에서 파일 지정:  python 수식_자동재계산.py 파일.hwpx
  (C) 그냥 더블클릭 → 폴더의 hwpx 목록이 뜨면 번호로 선택

  ※ 그냥 실행해도 더 이상 "모든 hwpx 자동 변환"을 하지 않습니다.
    반드시 목록에서 선택하므로 샘플 파일을 건드릴 일이 없습니다.
  ※ 이미 만든 _완성.hwp 나 단원체크 샘플은 목록에서 자동 제외됩니다.
--------------------------------------------------------------
"""
import os
import sys
import glob


# 변환 대상에서 자동 제외할 파일 패턴 (샘플/결과물)
EXCLUDE_KEYWORDS = ["_완성", "단원체크", "검증", "샘플", "sample", "test"]


def get_hwp():
    import win32com.client as win32
    hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    except Exception:
        pass
    hwp.XHwpWindows.Item(0).Visible = True
    return hwp


def process(hwp, in_path):
    in_path = os.path.abspath(in_path)
    out_path = os.path.splitext(in_path)[0] + "_완성.hwp"
    print(f"  열기: {os.path.basename(in_path)}")
    hwp.Open(in_path)

    n = 0
    ctrl = hwp.HeadCtrl
    while ctrl is not None:
        nxt = ctrl.Next
        if ctrl.UserDesc == "수식":
            try:
                props = ctrl.Properties
                props.SetItem("String", props.Item("String"))
                ctrl.Properties = props
                n += 1
            except Exception:
                pass
        ctrl = nxt
    try:
        hwp.HAction.Run("Recalc")
    except Exception:
        pass
    print(f"  수식 {n}개 재계산(간격 조정)")

    hwp.SaveAs(os.path.abspath(out_path), "HWP", "")
    print(f"  저장(2014용 hwp): {os.path.basename(out_path)}")
    return out_path, n


def is_excluded(path):
    name = os.path.basename(path)
    return any(k in name for k in EXCLUDE_KEYWORDS)


def pick_from_folder():
    """폴더의 hwpx 목록을 보여주고 번호로 선택하게 한다 (샘플/결과물 제외)."""
    cands = [f for f in glob.glob("*.hwpx") if not is_excluded(f)]
    if not cands:
        print("[안내] 변환할 .hwpx 가 폴더에 없습니다.")
        print("       (단원체크/_완성/검증 등은 자동 제외됩니다)")
        print("       파일을 이 스크립트 아이콘 위로 드래그&드롭 하거나,")
        print("       python 수식_자동재계산.py 파일.hwpx 로 실행하세요.")
        return []
    print("\n변환할 파일을 고르세요 (샘플·결과물은 제외했습니다):")
    for i, f in enumerate(cands, 1):
        print(f"  {i}. {f}")
    print("  0. 전체")
    try:
        sel = input("번호 입력 (기본 1): ").strip() or "1"
    except EOFError:
        sel = "1"
    if sel == "0":
        return cands
    try:
        idx = int(sel) - 1
        if 0 <= idx < len(cands):
            return [cands[idx]]
    except ValueError:
        pass
    print("[안내] 잘못된 입력. 1번 파일로 진행합니다.")
    return [cands[0]]


def main():
    try:
        import win32com.client  # noqa
    except ImportError:
        print("[오류] pywin32 가 없습니다.  cmd 에서:  pip install pywin32")
        input("엔터를 누르면 종료...")
        sys.exit(1)

    if len(sys.argv) >= 2:
        # 드래그&드롭 또는 파일 지정 → 그 파일만
        targets = sys.argv[1:]
    else:
        # 그냥 실행 → 목록에서 선택 (자동 일괄 변환 안 함)
        targets = pick_from_folder()

    if not targets:
        input("엔터를 누르면 종료...")
        sys.exit(0)

    print(f"\n대상 {len(targets)}개: {[os.path.basename(t) for t in targets]}")
    hwp = get_hwp()
    try:
        for t in targets:
            if not os.path.exists(t):
                print(f"[건너뜀] 없음: {t}"); continue
            print(f"\n처리: {os.path.basename(t)}")
            process(hwp, t)
    finally:
        try:
            hwp.Clear(1); hwp.Quit()
        except Exception:
            pass
    print("\n전부 완료. (_완성.hwp 파일이 생성되었습니다)")
    input("엔터를 누르면 종료...")


if __name__ == "__main__":
    main()
