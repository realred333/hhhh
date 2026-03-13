import win32gui
import win32con
import pyautogui
import time
import threading
import json
import sys
import os
import subprocess
from pyhwpx import Hwp

PROGRESS_FILE = "progress.json"
RESTART_EVERY = 300

def save_progress(idx):
    with open(PROGRESS_FILE, "w") as f:
        json.dump({"start_idx": idx}, f)

def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE) as f:
            return json.load(f).get("start_idx", 0)
    return 0

def handle_equation_window():
    """수식 편집기 창 제어: Delete 키 누르고 초고속 종료"""
    hwnd = 0
    timeout = 100
    while hwnd == 0 and timeout > 0:
        hwnd = win32gui.FindWindow(None, "수식 편집기")
        if hwnd == 0:
            time.sleep(0.01)
            timeout -= 1

    if hwnd:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.01)
        pyautogui.press('delete')
        time.sleep(0.01)
        pyautogui.hotkey('shift', 'esc')

def process_massive_equations(filepath, start_idx=0):
    print(f"문서를 열고 수식을 스캔합니다... (start_idx={start_idx})")
    hwp = Hwp()
    hwp.open(filepath)

    eq_ctrls = [c for c in hwp.ctrl_list if c.CtrlID == "eqed"]
    total_eq = len(eq_ctrls)

    if total_eq == 0 or start_idx >= total_eq:
        print("완료 또는 수식 없음.")
        if os.path.exists(PROGRESS_FILE):
            os.remove(PROGRESS_FILE)
        hwp.quit()
        print("🎉 모든 수식 변환 완료!")
        return

    print(f"🎯 총 {total_eq}개 중 {start_idx}번부터 재개")

    end_idx = min(start_idx + RESTART_EVERY, total_eq)
    current_idx = start_idx

    try:
        for i in range(start_idx, end_idx):
            current_idx = i
            ctrl = eq_ctrls[i]

            pos = ctrl.GetAnchorPos(0)
            hwp.hwp.SetPos(pos.Item("List"), pos.Item("Para"), pos.Item("Pos"))
            hwp.hwp.FindCtrl()

            t = threading.Thread(target=handle_equation_window)
            t.start()
            hwp.hwp.Run("EquationModify")
            t.join()

            if (i + 1) % 50 == 0 or i == total_eq - 1:
                print(f"  -> {i + 1} / {total_eq} 완료")

        # 이번 배치 완료 → 저장 후 재시작
        hwp.Save()
        hwp.quit()
        print(f"✅ {end_idx} / {total_eq} 까지 저장 완료. HWP 재시작...")

        save_progress(end_idx)
        subprocess.Popen([sys.executable] + sys.argv)
        sys.exit(0)

    except KeyboardInterrupt:
        print(f"\n🛑 긴급 정지! {current_idx + 1}번까지 저장 중...")
        hwp.Save()
        hwp.quit()
        save_progress(current_idx + 1)
        print(f"✅ 저장 완료. 다음 실행 시 자동으로 {current_idx + 1}번부터 재개됩니다.")


os.chdir(os.path.dirname(os.path.abspath(__file__)))

start = load_progress()
process_massive_equations(r"math1.hwpx", start_idx=start)
