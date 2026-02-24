import pyautogui
import pyperclip
import pygetwindow as gw
import time
from datetime import datetime
import threading
import queue

class TeamsSender:
    def __init__(self, chat_room_name="관제SO팀_TEAMS공유"):
        self.chat_room_name = chat_room_name
        self.msg_queue = queue.Queue()
        self.worker_thread = threading.Thread(target=self._process_queue, daemon=True)
        self.worker_thread.start()

    def send_alert(self, item_data):
        self.msg_queue.put(item_data)

    def _process_queue(self):
        while True:
            first_item = self.msg_queue.get()
            teams_window = None

            try:
                print(f"🚀 알람 그룹 전송 시작: 첫 번째 타겟 - {first_item.get('name')}")
                teams_window = self._enter_room_initial(first_item)
                
                if teams_window:
                    self._send_message_only(teams_window, first_item)
                    self.msg_queue.task_done()
                    
                    while True:
                        try:
                            print("⏳ 추가 알람 대기 중... (최대 3초 대기)")
                            next_item = self.msg_queue.get(timeout=3.0)
                            
                            print(f"🔗 연속 전송 모드 감지: {next_item.get('name')}")
                            
                            time.sleep(2.0)
                            
                            self._send_message_only(teams_window, next_item)
                            self.msg_queue.task_done()
                        
                        except queue.Empty:
                            print("🚫 더 이상 연속된 알람 없음. 배치 전송 종료.")
                            break
                else:
                    print("❌ Teams 창을 열지 못해 전송 실패")
                    self.msg_queue.task_done()

            except Exception as e:
                print(f"❌ Teams 전송 중 오류 발생: {e}")
                try: self.msg_queue.task_done()
                except: pass
            
            print("🏁 대기열 처리 완료. DAS로 복귀합니다.")
            self._return_focus_to_das(teams_window)
            time.sleep(1)

    def _return_focus_to_das(self, teams_window):
        if teams_window:
            try:
                if not teams_window.isMinimized:
                    teams_window.minimize()
                    time.sleep(0.5)
            except Exception as e:
                print(f"⚠️ Teams 최소화 실패: {e}")

        try:
            target_title = "DAS -" 
            das_windows = [w for w in gw.getAllWindows() if target_title in (w.title or "")]
            
            if das_windows:
                das_win = das_windows[0]
                if das_win.isMinimized:
                    das_win.restore()
                    time.sleep(0.2)
                das_win.activate()
                print("✅ DAS 프로그램 활성화 완료")
            else:
                print("⚠️ DAS 메인 창을 찾을 수 없습니다.")
        except Exception as e:
            print(f"⚠️ DAS 포커스 복귀 중 오류: {e}")

    def _activate_teams_window(self):
        print("🔍 Teams 창 찾는 중…")
        candidates = []

        for window in gw.getAllWindows():
            title = (window.title or "").strip()
            if not title: continue
            if "Microsoft Teams" in title or "teams" in title:
                candidates.append(window)

        if not candidates:
            print("❌ Teams 창 후보가 없습니다.")
            return None

        target_window = max(candidates, key=lambda w: w.width * w.height)

        start_time = time.time()
        while time.time() - start_time < 5.0:
            try:
                if target_window.isActive:
                    break
                
                if target_window.isMinimized:
                    target_window.restore()
                    time.sleep(0.5) 
                
                target_window.activate()
                time.sleep(0.5)

                if not target_window.isActive:
                    target_window.maximize()
                    time.sleep(0.2)
                    target_window.restore()

            except Exception as e:
                time.sleep(0.5)
        
        return target_window

    def _kill_popups(self):
        pyautogui.press('esc')
        time.sleep(0.1)
        pyautogui.press('esc')
        time.sleep(0.1)

    def _focus_chat_input_after_open(self, teams_window):
        time.sleep(0.5) 
        if teams_window is not None:
            try:
                if teams_window.width < 500 or teams_window.height < 400:
                     teams_window.maximize()
                     time.sleep(0.5)
            except: pass

            x = teams_window.left + (teams_window.width // 2)
            y = teams_window.top + teams_window.height - 100 
            
            pyautogui.click(x, y)
            time.sleep(0.2)
            pyautogui.click(x, y + 40)
            time.sleep(0.2)

            pyautogui.typewrite(' ')
            time.sleep(0.05)
            pyautogui.press('backspace')
            time.sleep(0.1)

    def _enter_room_initial(self, item_data):
        print(f"🔄 [초기화] Teams 방 입장 시도: {self.chat_room_name}")

        teams_window = self._activate_teams_window()
        if not teams_window:
            return None

        self._kill_popups()
        time.sleep(0.2)

        pyautogui.hotkey('ctrl', 'shift', 'f')
        time.sleep(1.5)

        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyautogui.press('delete') 
        time.sleep(0.1)
        pyautogui.press('backspace')
        time.sleep(0.1)
        
        pyperclip.copy(self.chat_room_name)
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1.5) 

        pyautogui.press('tab')
        time.sleep(0.2)
        pyautogui.press('tab')
        time.sleep(0.2)
        pyautogui.press('tab')
        time.sleep(0.2)
        
        pyautogui.press('down')
        time.sleep(0.2)
        
        pyautogui.press('enter')
        time.sleep(1.5)

        return teams_window

    def _send_message_only(self, teams_window, item_data):
        service_name = item_data.get('name', '알 수 없음')
        print(f"📨 메시지 작성 중... ({service_name})")

        current_time_str = datetime.now().strftime('%Y-%m-%d %H:%M')
        site_url = item_data.get('url', '-')
        group_raw = item_data.get('group', '')
        
        category_tag = ""
        if "US" in group_raw or "미국" in group_raw:
            category_tag = "미국"
        elif "JP" in group_raw or "일본" in group_raw:
            category_tag = "일본"
        elif "Roaming" in group_raw or "로밍" in group_raw:
            category_tag = "로밍"
        else:
            category_tag = group_raw if group_raw else ""

        tag_prefix = f"{category_tag} " if category_tag else ""

        message = (
            f"[관제SO팀][Downdetector] {tag_prefix}{service_name} User Report 증가 발생\n"
            f"ㅇ일시 : {current_time_str} ~\n"
            f"ㅇ대상 : {service_name}\n"
            f"ㅇ건수 : 00건\n"
            f"ㅇ원인 : 확인중\n"
            f"ㅇ서비스영향 : 확인중\n"
            f"ㅇ고객문의 : 확인중\n"
            f"    - Site : {site_url}\n"
            f"    - Status : https://3spage.netlify.app/\n"
            f"ㅇ조치/확인내역 : 확인중\n"
            f"ㅇ기타 : "
        )

        if not teams_window.isActive:
            try: teams_window.activate()
            except: pass
            time.sleep(0.5)
            
        self._kill_popups() 
        self._focus_chat_input_after_open(teams_window)
        
        pyperclip.copy(message)
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.8)
        pyautogui.press('enter')
        time.sleep(0.2)

        print(f"✅ 전송 완료: {service_name}")