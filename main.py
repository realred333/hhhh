import win32gui
import win32con
import pyautogui
import time
import threading
from pyhwpx import Hwp

def handle_equation_window():
    """수식 편집기 창 제어 및 닫기 스레드"""
    hwnd = 0
    timeout = 50
    while hwnd == 0 and timeout > 0:
        hwnd = win32gui.FindWindow(None, "수식 편집기")
        if hwnd == 0:
            time.sleep(0.1)
            timeout -= 1
            
    if hwnd:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.3)
        
        rect = win32gui.GetWindowRect(hwnd)
        click_x = rect[0] + ((rect[2] - rect[0]) // 2)
        click_y = rect[1] + ((rect[3] - rect[1]) // 3)
        
        pyautogui.click(click_x, click_y)
        time.sleep(0.1)
        
        pyautogui.press('right')
        time.sleep(0.05)
        pyautogui.press('space')
        time.sleep(0.05)
        pyautogui.press('backspace')
        time.sleep(0.1)
        
        pyautogui.hotkey('shift', 'esc')

def process_massive_equations(filepath, start_idx=0):
    print("문서를 열고 수식을 스캔합니다...")
    hwp = Hwp()
    hwp.open(filepath)
    
    eq_ctrls = [c for c in hwp.ctrl_list if c.CtrlID == "eqed"]
    total_eq = len(eq_ctrls)
    
    if total_eq == 0:
        print("문서에 수식이 없습니다.")
        hwp.quit()
        return
        
    print(f"🎯 총 {total_eq}개의 수식을 발견했습니다!")
    if start_idx > 0:
        print(f"🚀 {start_idx + 1}번째 수식부터 이어서 작업을 시작합니다.\n")
    else:
        print("작업을 시작합니다. 200개마다 자동 저장됩니다.\n")
        
    for i in range(start_idx, total_eq):
        ctrl = eq_ctrls[i]
        
        pos = ctrl.GetAnchorPos(0)
        hwp.hwp.SetPos(pos.Item("List"), pos.Item("Para"), pos.Item("Pos"))
        hwp.hwp.FindCtrl()
        
        t = threading.Thread(target=handle_equation_window)
        t.start()
        
        hwp.hwp.Run("EquationModify")
        
        t.join()
        time.sleep(0.2)
        
        # 200개마다 디스크에 덮어쓰기 저장 (메모리 안정화)
        if (i + 1) % 200 == 0:
            hwp.Save()
            print(f"✅ {i + 1} / {total_eq} 완료 -> 중간 저장 완료")
        # 50개 단위 출력
        elif (i + 1) % 50 == 0 or i == total_eq - 1:
            print(f"  -> {i + 1} / {total_eq} 완료")
            
    # 모든 작업이 끝나면 최종 저장 후 닫기
    hwp.Save()
    hwp.quit()
    print("🎉 6684개 대장정 완료! 파일이 모두 정규화되었습니다.")

# ==========================================
# 실행 부분
# 만약 3200번에서 뻗었다면 start_idx=3200 으로 바꾸고 다시 돌리시면 됩니다.
process_massive_equations(r"math1.hwpx", start_idx=0)