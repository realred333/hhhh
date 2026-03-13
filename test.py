import win32gui
import win32con
import pyautogui
import time
import threading
from pyhwpx import Hwp

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
        time.sleep(0.01) # 창 뜨자마자 아주 살짝만 대기
        
        # 1. 딜리트(Delete) 키 입력!
        pyautogui.press('delete')
        time.sleep(0.01) # 🌟 HWP 엔진이 정규화 렌더링을 할 시간 살짝 부여
        
        # 2. 저장 및 닫기
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
        
    print(f"🎯 총 {total_eq}개의 수식 변환 시작 (Delete 트리거 + 안전 종료 모드)")
    
    # 이 부분은 키 씹힘 방지를 위해 다시 주석 처리하거나 빼는 걸 추천합니다.
    # pyautogui.PAUSE = 0.01 
    
    current_idx = start_idx
    
    try:
        for i in range(start_idx, total_eq):
            current_idx = i # 현재 진행 중인 번호 기억
            ctrl = eq_ctrls[i]
            
            pos = ctrl.GetAnchorPos(0)
            hwp.hwp.SetPos(pos.Item("List"), pos.Item("Para"), pos.Item("Pos"))
            hwp.hwp.FindCtrl()
            
            t = threading.Thread(target=handle_equation_window)
            t.start()
            
            hwp.hwp.Run("EquationModify")
            t.join()
            
            # 200개마다 정상 자동 저장
            if (i + 1) % 200 == 0:
                hwp.Save()
                print(f"✅ {i + 1} / {total_eq} 완료 -> 중간 저장 완료")
            elif (i + 1) % 50 == 0 or i == total_eq - 1:
                print(f"  -> {i + 1} / {total_eq} 완료")
                
        # 무사히 끝까지 다 돌았을 때 최종 저장
        hwp.Save()
        hwp.quit()
        print("🎉 모든 수식 변환 대장정 완료!")
        
    except KeyboardInterrupt:
        # 🚨 여기서 Ctrl + C 를 감지합니다! 🚨
        print(f"\n\n🛑 [긴급 정지] Ctrl+C가 눌렸습니다!")
        print(f"💾 지금까지 바뀐 {current_idx + 1}번째 수식까지 안전하게 파일에 저장합니다...")
        
        hwp.Save()  # 무조건 디스크에 박제!
        hwp.quit()  # 유령 프로세스 찌꺼기 방지
        
        print(f"✅ 저장 완료 및 한글 정상 종료됨.")
        print(f"👉 다음에 다시 돌릴 때는 start_idx={current_idx + 1} 로 설정하고 시작하세요!\n")

# 실행 (경로 맞춰주세요)
process_massive_equations(r"math1.hwpx", start_idx=0)