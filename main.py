import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import json
import os
import pygame
import webbrowser
import sys
import subprocess
from datetime import datetime, timedelta
from monitor_engine import MonitorEngine
from teams_send import TeamsSender

# --- 경로 설정 ---
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
    INTERNAL_DIR = BASE_DIR 
else:
    INTERNAL_DIR = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = INTERNAL_DIR

CONFIG_FILE = os.path.join(BASE_DIR, "save", "config.json")
if not os.path.exists(CONFIG_FILE): 
    CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

KEYWORD_FILE = os.path.join(BASE_DIR, "save", "keywords.json")
POINT_FILE = os.path.join(BASE_DIR, "save", "point.json")

AUDIO_US = os.path.join(INTERNAL_DIR, "audio", "uscaution.mp3")
AUDIO_JP = os.path.join(INTERNAL_DIR, "audio", "jpcaution.mp3")
AUDIO_ROM = os.path.join(INTERNAL_DIR, "audio", "romcaution.mp3")

# --- 색상 테마 ---
COLOR_BG = "#1E1E1E"        
COLOR_HEADER = "#252526"    
COLOR_TEXT_MAIN = "#FFFFFF" 
COLOR_TEXT_SUB = "#AAAAAA"  
COLOR_ACCENT = "#4EC9B0"
COLOR_BLUE_HIGHLIGHT = "#569CD6"
COLOR_SUCCESS = "#4CAF50"   
COLOR_WARNING = "#FFC107"   
COLOR_DANGER = "#F44336"    
COLOR_GRAY = "#3E3E42"      
COLOR_NODATA = "#000000"    

class DASApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DAS - Downdetector Alarm System")
        self.geometry("1920x1080")
        self.configure(bg=COLOR_BG)
        self.state('zoomed') 

        icon_path = None
        potential_paths = [
            os.path.join(BASE_DIR, "DAS.ico"),
            os.path.join(BASE_DIR, "DAS.png"),
            os.path.join(INTERNAL_DIR, "DAS.ico")
        ]
            
        for p in potential_paths:
            if os.path.exists(p):
                icon_path = p
                break

        if icon_path:
            try:
                if icon_path.endswith('.ico'):
                    self.iconbitmap(icon_path)
                elif icon_path.endswith('.png'):
                    from tkinter import PhotoImage
                    img = PhotoImage(file=icon_path)
                    self.iconphoto(True, img)
            except Exception as e:
                print(f"⚠️ 아이콘 설정 실패: {e}")

        self.bind("<F11>", self.toggle_fullscreen)
        self.bind("<Escape>", self.quit_app)

        self.load_config()
        self.check_point_file()
        
        room_name = self.config.get('chat_room', "관제SO팀_TEAMS공유")
        self.teams_bot = TeamsSender(room_name)
        self.is_teams_enabled = True
        self.alert_history = {}

        self.engine = MonitorEngine()
        
        self.is_monitoring = False
        self.is_alarm_active = False
        self.is_sound_enabled = True 
        self.stop_event = threading.Event()
        self.monitor_thread = None 
        self.current_popup = None
        self.is_restarting = False

        self.last_collection_time = None
        self.next_run_time = None  
        
        self.create_ui()
        
        pygame.mixer.init()
        self.update_clock()

        self.after(1000, self.start_monitoring)

    def load_config(self):
        try:
            content = self.read_file_safe(CONFIG_FILE)
            if content:
                self.config = json.loads(content)
            else:
                self.config = {
                    "regions": [], 
                    "fixed_targets": [], 
                    "check_interval_seconds": 60,
                    "alert_cool_down_seconds": 60
                }
        except Exception as e:
            print(f"Config Load Error: {e}")
            self.config = {
                "regions": [], 
                "fixed_targets": [], 
                "check_interval_seconds": 60,
                "alert_cool_down_seconds": 60
            }

    def check_point_file(self):
        """좌표 파일이 없으면 기본값으로 생성"""
        if not os.path.exists(POINT_FILE):
            default_points = {
                "click1": {"x": 535, "y": 375},
                "click2": {"x": 535, "y": 425}
            }
            try:
                save_dir = os.path.dirname(POINT_FILE)
                if not os.path.exists(save_dir):
                    os.makedirs(save_dir)
                    
                with open(POINT_FILE, 'w', encoding='utf-8') as f:
                    json.dump(default_points, f, indent=4)
            except: pass

    def read_file_safe(self, filepath):
        if not os.path.exists(filepath): return None
        encodings = ['utf-8', 'cp949', 'euc-kr', 'latin-1']
        for enc in encodings:
            try:
                with open(filepath, 'r', encoding=enc) as f:
                    return f.read()
            except UnicodeDecodeError: continue
            except Exception: return None
        return None

    def create_ui(self):
        header = tk.Frame(self, bg=COLOR_HEADER, height=110)
        header.pack(fill="x", side="top", pady=(0, 0))
        header.pack_propagate(False)

        btn_frame = tk.Frame(header, bg=COLOR_HEADER)
        btn_frame.pack(side="right", padx=60, fill="y", pady=30) 

        def mk_btn(parent, text, cmd, color=COLOR_GRAY, width=12):
            btn = tk.Button(parent, text=text, command=cmd, bg=color, fg="white", 
                            font=("Malgun Gothic", 11, "bold"), relief="flat", width=width, cursor="hand2",
                            activebackground="#555", activeforeground="white")
            btn.pack(side="left", padx=5, fill="y") 
            return btn

        self.btn_sound = mk_btn(btn_frame, "알람 ON", self.toggle_sound, COLOR_SUCCESS, width=10)
        self.btn_teams = mk_btn(btn_frame, "Teams ON", self.toggle_teams, COLOR_SUCCESS, width=10)
        self.btn_monitor = mk_btn(btn_frame, "감시 시작", self.toggle_monitoring, COLOR_SUCCESS, width=12)
        
        tk.Frame(btn_frame, width=15, bg=COLOR_HEADER).pack(side="left")
        
        mk_btn(btn_frame, "좌표 설정", self.open_point_popup, "#444", width=9)
        mk_btn(btn_frame, "키워드 설정", self.open_keyword_popup, "#444", width=9)
        mk_btn(btn_frame, "시스템 설정", self.open_config_popup, "#444", width=9)
        mk_btn(btn_frame, "종료", self.quit_app, COLOR_DANGER, width=9)

        info_frame = tk.Frame(header, bg=COLOR_HEADER)
        info_frame.pack(side="left", padx=40, fill="y", pady=25) 

        for i in [0, 2, 4, 6]:
            info_frame.columnconfigure(i, minsize=160, weight=0)
        
        for i in [1, 3, 5]:
            info_frame.columnconfigure(i, minsize=40, weight=0)

        def add_separator(parent, col):
            sep = tk.Frame(parent, width=1, bg="#555")
            sep.grid(row=0, column=col, rowspan=2, sticky="ns", padx=30)

        self.lbl_cur_time = self.create_time_widget(info_frame, "현재 시간", "00:00:00", 0)
        add_separator(info_frame, 1)
        
        self.lbl_last_time = self.create_time_widget(info_frame, "최근 수집 시간", "-", 2)
        add_separator(info_frame, 3)
        
        self.lbl_next_time = self.create_time_widget(info_frame, "다음 수집 시간", "-", 4)
        add_separator(info_frame, 5)

        self.lbl_sys_status = self.create_time_widget(info_frame, "시스템 상태", "수집 중지 상태", 6)
        self.lbl_sys_status.config(fg="#888")

        container = tk.Frame(self, bg=COLOR_BG)
        container.pack(fill="both", expand=True, padx=20, pady=20)

        self.canvas = tk.Canvas(container, bg=COLOR_BG, highlightthickness=0)
        self.scroll_frame = tk.Frame(self.canvas, bg=COLOR_BG)

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        
        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self.on_canvas_configure)

        self.canvas.pack(side="left", fill="both", expand=True)
        
        self.lbl_guide = tk.Label(self.scroll_frame, text="\n상단의 '감시 시작' 버튼을 눌러주세요.", 
                                  fg="#666", bg=COLOR_BG, font=("Malgun Gothic", 20))
        self.lbl_guide.pack(fill="x", pady=200)

        self.is_monitoring = False
        self.btn_monitor.config(text="감시 시작", bg=COLOR_SUCCESS, fg="White")

    def create_time_widget(self, parent, title, init_val, col):
        frame = tk.Frame(parent, bg=COLOR_HEADER)
        frame.grid(row=0, column=col, sticky="w") 
        
        tk.Label(frame, text=title, fg="#CCC", bg=COLOR_HEADER, font=("Malgun Gothic", 10)).pack(anchor="w")
        lbl = tk.Label(frame, text=init_val, fg="white", bg=COLOR_HEADER, font=("Verdana", 20, "bold"))
        lbl.pack(anchor="w")
        return lbl

    def update_system_status(self, text, color):
        self.lbl_sys_status.config(text=text, fg=color)

    def on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def update_clock(self):
        if self.is_restarting: return

        now = datetime.now()
        self.lbl_cur_time.config(text=now.strftime("%H:%M:%S"))
        
        # --- 자동 재시작 (00시, 12시) ---
        if now.hour in [0, 12] and now.minute == 0 and now.second == 0:
            print(f"⏰ {now.strftime('%H:%M:%S')} - 정기 점검을 위한 자동 재시작 시도")
            
            if not self.is_restarting:
                self.restart_application()
            return

        if self.is_monitoring:
            if self.next_run_time:
                remaining = (self.next_run_time - now).total_seconds()
                if remaining > 0:
                    self.lbl_next_time.config(text=f"{int(remaining)}초", fg="white")
                else:
                    self.lbl_next_time.config(text="0초", fg=COLOR_WARNING)
            else:
                self.lbl_next_time.config(text="0초", fg=COLOR_WARNING)
        else:
            self.lbl_next_time.config(text="-", fg="#888")

        self.after(1000, self.update_clock)

    def restart_application(self):
        if self.is_restarting: return 
        
        self.is_restarting = True
        self.update_system_status("시스템 재시작 중...", COLOR_DANGER)
        
        if self.is_monitoring:
            self.stop_monitoring()
        
        self.after(2000, self._perform_restart)

    def _perform_restart(self):
        print("♻️ onedir 모드 재시작 프로세스 진입...")

        try:
            self.stop_alarm()
            if pygame.mixer.get_init(): pygame.mixer.quit()
            if pygame.get_init(): pygame.quit()
        except: pass

        try:
            self.engine.close_browser()
        except: pass

        try:
            new_env = os.environ.copy()
            if '_MEIPASS2' in new_env:
                del new_env['_MEIPASS2']
            
            executable = sys.executable
            
            creation_flags = subprocess.CREATE_NEW_CONSOLE if sys.platform == 'win32' else 0
            
            subprocess.Popen(
                [executable] + sys.argv[1:],
                env=new_env,
                close_fds=True,
                creationflags=creation_flags
            )
            
            print("♻️ 새 프로세스 호출 완료. 현재 프로세스 종료.")
            
            self.quit()
            self.destroy()
            os._exit(0)

        except Exception as e:
            print(f"❌ 재시작 실패: {e}")
            self.is_restarting = False
            self.update_clock()

    def toggle_fullscreen(self, event=None):
        self.attributes("-fullscreen", not self.attributes("-fullscreen"))

    def open_file_editor(self, title, filepath):
        if not os.path.exists(filepath):
            try:
                d = os.path.dirname(filepath)
                if d and not os.path.exists(d): os.makedirs(d)
                with open(filepath, 'w', encoding='utf-8') as f: f.write("{}")
            except: pass

        popup = tk.Toplevel(self)
        popup.title(title)
        popup.geometry("800x600")
        popup.configure(bg=COLOR_HEADER)
        
        tk.Label(popup, text=f"{title} 편집", fg="white", bg=COLOR_HEADER, font=("Malgun Gothic", 12, "bold")).pack(pady=10)
        text_area = tk.Text(popup, bg="#333", fg="white", font=("Consolas", 11), insertbackground="white", padx=10, pady=10)
        text_area.pack(expand=True, fill="both", padx=10, pady=(0, 10))

        content = self.read_file_safe(filepath)
        if content: text_area.insert("1.0", content)

        def save_content():
            new_content = text_area.get("1.0", "end-1c")
            try:
                if filepath.endswith(".json"):
                    json.loads(new_content)
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(new_content)
                if filepath == CONFIG_FILE:
                    self.load_config()
                messagebox.showinfo("저장", "저장되었습니다.")
                popup.destroy()
            except Exception as e:
                messagebox.showerror("오류", str(e))

        btn_frame = tk.Frame(popup, bg=COLOR_HEADER)
        btn_frame.pack(fill="x", pady=10, padx=10)
        tk.Button(btn_frame, text="취소", command=popup.destroy, bg="#555", fg="white", width=10, font=("Malgun Gothic", 10)).pack(side="right", padx=5)
        tk.Button(btn_frame, text="저장", command=save_content, bg=COLOR_SUCCESS, fg="white", width=10, font=("Malgun Gothic", 10)).pack(side="right", padx=5)

    # =========================================================================
    # 좌표 설정 팝업 기능
    # =========================================================================
    def open_point_popup(self):
        popup = tk.Toplevel(self)
        popup.title("로봇 우회 좌표 설정")
        popup.geometry("450x300")
        popup.configure(bg=COLOR_HEADER)
        
        current_data = {"click1": {"x": 535, "y": 375}, "click2": {"x": 535, "y": 425}}
        try:
            if os.path.exists(POINT_FILE):
                with open(POINT_FILE, 'r', encoding='utf-8') as f:
                    current_data = json.load(f)
        except: pass

        tk.Label(popup, text="로봇 우회 클릭 좌표 설정", fg="white", bg=COLOR_HEADER, font=("Malgun Gothic", 14, "bold")).pack(pady=20)

        input_frame = tk.Frame(popup, bg=COLOR_HEADER)
        input_frame.pack(pady=10)

        entries = {}
        def create_entry_row(row, label_text, key_prefix):
            tk.Label(input_frame, text=label_text, fg="#CCC", bg=COLOR_HEADER, font=("Malgun Gothic", 11)).grid(row=row, column=0, columnspan=2, pady=(10, 5))
            
            tk.Label(input_frame, text="X:", fg="white", bg=COLOR_HEADER).grid(row=row+1, column=0, padx=5)
            e_x = tk.Entry(input_frame, width=8, bg="#333", fg="white", insertbackground="white")
            e_x.insert(0, str(current_data[key_prefix]["x"]))
            e_x.grid(row=row+1, column=1, padx=5)
            entries[f"{key_prefix}_x"] = e_x

            tk.Label(input_frame, text="Y:", fg="white", bg=COLOR_HEADER).grid(row=row+1, column=2, padx=5)
            e_y = tk.Entry(input_frame, width=8, bg="#333", fg="white", insertbackground="white")
            e_y.insert(0, str(current_data[key_prefix]["y"]))
            e_y.grid(row=row+1, column=3, padx=5)
            entries[f"{key_prefix}_y"] = e_y

        create_entry_row(0, "1차 시도 클릭 좌표", "click1")
        create_entry_row(2, "2차 시도 클릭 좌표", "click2")

        def save_points():
            try:
                new_data = {
                    "click1": {
                        "x": int(entries["click1_x"].get()),
                        "y": int(entries["click1_y"].get())
                    },
                    "click2": {
                        "x": int(entries["click2_x"].get()),
                        "y": int(entries["click2_y"].get())
                    }
                }
                save_dir = os.path.dirname(POINT_FILE)
                if not os.path.exists(save_dir): os.makedirs(save_dir)

                with open(POINT_FILE, 'w', encoding='utf-8') as f:
                    json.dump(new_data, f, indent=4)
                
                messagebox.showinfo("저장", "좌표가 저장되었습니다.\n다음 수집부터 적용됩니다.")
                popup.destroy()
            except ValueError:
                messagebox.showerror("오류", "좌표값은 숫자만 입력해주세요.")
            except Exception as e:
                messagebox.showerror("오류", str(e))

        btn_frame = tk.Frame(popup, bg=COLOR_HEADER)
        btn_frame.pack(fill="x", pady=20, padx=20)
        tk.Button(btn_frame, text="취소", command=popup.destroy, bg="#555", fg="white", width=10).pack(side="right", padx=5)
        tk.Button(btn_frame, text="저장", command=save_points, bg=COLOR_SUCCESS, fg="white", width=10).pack(side="right", padx=5)

    def open_keyword_popup(self):
        self.open_file_editor("키워드 설정", KEYWORD_FILE)

    def open_config_popup(self):
        self.open_file_editor("시스템 설정", CONFIG_FILE)

    def toggle_sound(self):
        self.is_sound_enabled = not self.is_sound_enabled
        if self.is_sound_enabled:
            self.btn_sound.config(text="알람 ON", bg=COLOR_SUCCESS)
        else:
            self.btn_sound.config(text="알람 OFF", bg=COLOR_GRAY)
            self.stop_alarm()
    
    def toggle_teams(self):
        self.is_teams_enabled = not self.is_teams_enabled
        if self.is_teams_enabled:
            self.btn_teams.config(text="Teams ON", bg=COLOR_SUCCESS)
        else:
            self.btn_teams.config(text="Teams OFF", bg=COLOR_GRAY)

    def toggle_monitoring(self):
        if self.is_monitoring:
            self.stop_monitoring()
        else:
            self.start_monitoring()

    def start_monitoring(self):
        self.is_monitoring = True
        self.stop_event.clear()
        
        self.btn_monitor.config(text="감시 중지", bg=COLOR_DANGER, fg="white")
        self.lbl_last_time.config(fg="white")
        self.update_system_status("시스템 초기화", COLOR_WARNING)
        
        if hasattr(self, 'lbl_guide') and self.lbl_guide and self.lbl_guide.winfo_exists():
            self.lbl_guide.destroy()
        
        self.monitor_thread = threading.Thread(target=self.monitor_task)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()

    def stop_monitoring(self):
        self.is_monitoring = False
        self.stop_event.set()
        self.next_run_time = None 
        
        self.btn_monitor.config(text="감시 시작", bg=COLOR_SUCCESS, fg="white")
        self.update_system_status("수집 중지 상태", "#888")
        self.lbl_next_time.config(text="-", fg="#888")
        self.lbl_last_time.config(fg="#888")
        self.stop_alarm()
        
        try:
            print("🛑 브라우저 강제 종료 요청...")
            self.engine.close_browser()
        except Exception as e:
            print(f"브라우저 종료 중 오류: {e}")

    def monitor_task(self):
        while self.is_monitoring and not self.stop_event.is_set():
            fixed_list = self.config.get('fixed_targets', [])
            
            self.after(0, lambda: self.update_system_status("데이터 수집 상태", COLOR_ACCENT))
            
            try:
                if self.stop_event.is_set(): break 

                results = self.engine.scan_and_check(
                    self.config['regions'], 
                    KEYWORD_FILE,
                    fixed_list,
                    stop_event=self.stop_event 
                )
                
                if self.stop_event.is_set(): break 

                self.last_collection_time = datetime.now()
                time_str = self.last_collection_time.strftime("%H:%M:%S")
                
                self.after(0, lambda t=time_str: self.lbl_last_time.config(text=t, fg=COLOR_BLUE_HIGHLIGHT))
                self.after(0, self.update_ui, results)
                
                self.engine.close_browser()

            except Exception as e:
                print(f"Monitor Error: {e}")
                try: self.engine.close_browser()
                except: pass
            
            if self.stop_event.is_set(): break

            self.after(0, lambda: self.update_system_status("수집 대기 상태", COLOR_SUCCESS))

            wait_sec = self.config.get("check_interval_seconds", 60)
            self.next_run_time = datetime.now() + timedelta(seconds=wait_sec)
            
            self.after(0, lambda s=wait_sec: self.lbl_next_time.config(text=f"{s}초", fg="white"))

            for _ in range(wait_sec):
                if self.stop_event.is_set(): break
                time.sleep(1)

    def update_ui(self, results):
        if not self.is_monitoring: return 

        for widget in self.scroll_frame.winfo_children():
            widget.destroy()

        if not results:
            tk.Label(self.scroll_frame, text="데이터 없음", fg=COLOR_TEXT_SUB, bg=COLOR_BG, font=("Malgun Gothic", 16)).pack(pady=50)
            self.stop_alarm()
            return

        alarm_needed = False
        alarm_groups = set()
        popup_items = []  
        teams_queue = [] 
        
        groups = {}
        for r in results:
            g = r.get('group', 'Unknown')
            if g not in groups: groups[g] = []
            groups[g].append(r)

        ordered_groups = [r['name'] for r in self.config['regions']]
        final_group_keys = [k for k in ordered_groups if k in groups]
        for k in groups:
            if k not in final_group_keys: final_group_keys.append(k)

        MAX_COL = 7

        for grp_name in final_group_keys:
            items = groups[grp_name]
            if len(items) > MAX_COL:
                items = items[:MAX_COL]

            group_frame = tk.Frame(self.scroll_frame, bg=COLOR_BG)
            group_frame.pack(fill="x", pady=(15, 10), padx=5) 
            
            tk.Frame(group_frame, bg="#FFFFFF", width=4, height=22).pack(side="left") 
            tk.Label(group_frame, text=f" {grp_name}", fg=COLOR_TEXT_MAIN, bg=COLOR_BG, font=("Malgun Gothic", 16, "bold")).pack(side="left", padx=8)

            grid_frame = tk.Frame(self.scroll_frame, bg=COLOR_BG)
            grid_frame.pack(fill="x", padx=5)

            for i, item in enumerate(items):
                row = 0
                col = i 

                status_color = COLOR_SUCCESS 
                display_msg = "NORMAL"
                text_color = "white"
                
                if item['status'] == "CRITICAL":
                    status_color = COLOR_DANGER
                    display_msg = "CRITICAL"
                    
                    if self.is_teams_enabled and item.get('alarm_trigger', False):
                        current_time = time.time()
                        
                        unique_key = f"{item['name']}_{item['group']}"
                        last_sent = self.alert_history.get(unique_key, 0)
                        
                        cool_down = self.config.get('alert_cool_down_seconds', 60)
                        
                        if current_time - last_sent > cool_down:
                            teams_queue.append(item)
                            self.alert_history[unique_key] = current_time

                elif item['status'] == "WARNING":
                    status_color = COLOR_WARNING
                    display_msg = "WARNING"
                elif item['status'] == "NODATA":
                    status_color = COLOR_NODATA
                    display_msg = "NO DATA"
                    text_color = "#AAAAAA" 
                else:
                    display_msg = "NORMAL"

                if item['alarm_trigger']:
                    alarm_needed = True
                    alarm_groups.add(grp_name)
                    popup_items.append(item)

                card = tk.Frame(grid_frame, bg=status_color, highlightbackground="#555", highlightthickness=1)
                card.grid(row=row, column=col, padx=8, pady=8, sticky="nsew") 
                
                content = tk.Frame(card, bg=status_color)
                content.pack(fill="both", expand=True, padx=5, pady=40) 

                name_lbl = tk.Label(content, text=item['name'], fg=text_color, bg=status_color, 
                                    font=("Segoe UI", 16, "bold"), 
                                    cursor="hand2", anchor="center")
                name_lbl.pack(fill="x", pady=(0, 15)) 
                name_lbl.bind("<Button-1>", lambda e, url=item['url']: webbrowser.open(url))

                tk.Label(content, text=display_msg, fg=text_color, bg=status_color, 
                         font=("Segoe UI", 14, "bold"), anchor="center").pack(fill="both", expand=True)
                
                detail_text = item['msg'] if display_msg != "NORMAL" else ""
                if detail_text:
                    tk.Label(content, text=detail_text[:25] + "...", fg="white", bg=status_color, 
                             font=("Segoe UI", 10), anchor="center").pack(pady=(10,0))

            for i in range(MAX_COL):
                grid_frame.columnconfigure(i, weight=1, uniform="group1")

        if alarm_needed and self.is_sound_enabled:
            sound_file = None
            if any("Roaming" in g for g in alarm_groups) or any("로밍" in g for g in alarm_groups):
                sound_file = AUDIO_ROM
            elif any("JP" in g for g in alarm_groups) or any("일본" in g for g in alarm_groups):
                sound_file = AUDIO_JP
            else: 
                sound_file = AUDIO_US
            self.trigger_alarm(sound_file)
            
            if popup_items:
                self.show_critical_popup(popup_items, teams_queue)
        else:
            self.stop_alarm()

    def show_critical_popup(self, target_items, teams_items=None):
        if self.current_popup and self.current_popup.winfo_exists():
            self.current_popup.destroy()
        
        popup = tk.Toplevel(self)
        self.current_popup = popup
        popup.title("CRITICAL ALARM")
        
        popup.geometry("700x500") 
        
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - 700) // 2
        y = (screen_height - 500) // 2
        popup.geometry(f"+{x}+{y}")
        
        popup.configure(bg=COLOR_DANGER)
        popup.attributes("-topmost", True) 
        popup.overrideredirect(True) 

        tk.Label(popup, text="⚠️ CRITICAL ALERT ⚠️", fg="yellow", bg=COLOR_DANGER, 
                 font=("Verdana", 26, "bold")).pack(pady=(30, 20))
        
        categorized = {"미국": [], "일본": [], "로밍": []}
        
        for item in target_items:
            grp = item.get('group', 'Unknown')
            name = item.get('name', 'Unknown')
            
            if "US" in grp or "미국" in grp:
                categorized["미국"].append(name)
            elif "JP" in grp or "일본" in grp:
                categorized["일본"].append(name)
            elif "Roaming" in grp or "로밍" in grp:
                categorized["로밍"].append(name)
        
        msg_lines = []
        if categorized["미국"]:
            msg_lines.append(f"미국 : {', '.join(categorized['미국'])}")
        if categorized["일본"]:
            msg_lines.append(f"일본 : {', '.join(categorized['일본'])}")
        if categorized["로밍"]:
            msg_lines.append(f"로밍 : {', '.join(categorized['로밍'])}")
            
        final_msg_text = "\n\n".join(msg_lines)

        tk.Label(popup, text=final_msg_text, fg="white", bg=COLOR_DANGER, 
                 font=("Malgun Gothic", 20, "bold"), wraplength=650, justify="left").pack(pady=10, expand=True)
        
        tk.Label(popup, text="Critical 알람 발생 리포트 수 확인 요망", fg="white", bg=COLOR_DANGER, 
                 font=("Malgun Gothic", 16, "bold")).pack(pady=(0, 30))

        def close_popup():
            if self.current_popup:
                self.current_popup.destroy()
                self.current_popup = None
            
            if teams_items:
                threading.Thread(target=self.process_teams_queue, args=(teams_items,), daemon=True).start()

        popup.after(10000, close_popup)
        popup.bind("<Button-1>", lambda e: close_popup())

    def process_teams_queue(self, items):
        """Teams 대기열에 있는 알람들을 전송"""
        for item in items:
            print(f"🚀 {item['name']} 팝업 종료 후 Teams 알람 전송 시작")
            try:
                self.teams_bot.send_alert(item)
            except Exception as e:
                print(f"Teams 전송 실패 ({item['name']}): {e}")

    def trigger_alarm(self, sound_file):
        if not self.is_alarm_active:
            self.is_alarm_active = True
        
        if sound_file and os.path.exists(sound_file):
            try:
                if not pygame.mixer.music.get_busy():
                     pygame.mixer.music.load(sound_file)
                     pygame.mixer.music.play()
            except: pass

    def stop_alarm(self):
        self.is_alarm_active = False
        if pygame.mixer.music.get_busy(): pygame.mixer.music.stop()
        self.lbl_cur_time.config(fg="white") 

    def quit_app(self, e=None):
        self.stop_monitoring() 
        self.destroy() 
        sys.exit(0)

if __name__ == "__main__":
    app = DASApp()
    app.mainloop()