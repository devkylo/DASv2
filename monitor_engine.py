from DrissionPage import ChromiumPage, ChromiumOptions
import time
from urllib.parse import urlparse
import os
import json
import random
import pyautogui
import pygetwindow as gw

class MonitorEngine:
    def __init__(self):
        self.page = None

    def load_keywords(self, filepath):
        if not os.path.exists(filepath):
            filepath = os.path.join("save", "keywords.json")
        try:
            if os.path.exists(filepath):
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data.get("keywords", [])
            else:
                return []
        except Exception:
            return []

    def _init_browser(self):
        print("브라우저 환경 초기화 중...")
        
        self.close_browser()
        time.sleep(2.5)

        try:
            co = ChromiumOptions()
            
            profile_path = os.path.join(os.getcwd(), "chrome_temp_profile")
            co.set_user_data_path(profile_path)
            
            co.incognito()

            co.set_argument('--disable-extensions')
            co.set_argument('--disable-popup-blocking')
            co.set_argument('--no-first-run')
            co.set_argument('--no-default-browser-check')
            co.set_argument('about:blank') # 빈 페이지로 시작

            co.auto_port()
            self.page = ChromiumPage(co)
            
            try:
                if self.page.tabs_count > 1:
                    self.page.close_other_tabs()
                self.page.get('about:blank')
            except Exception as e:
                print(f"초기 탭 정리 중 오류 (무시 가능): {e}")

            try:
                self.page.set.window.location(1600, 1000)
                self.page.set.window.size(800, 600)
            except: pass
            
            print("새 브라우저 실행 완료")

        except Exception as e:
            print(f"브라우저 실행 실패 (재시도 필요): {e}")
            try:
                co = ChromiumOptions()
                co.auto_port()
                self.page = ChromiumPage(co)
            except:
                print("오류: 브라우저를 열 수 없습니다.")

    def close_browser(self):
        """브라우저 완전 종료 및 메모리 해제"""
        if self.page:
            try:
                print("브라우저 종료 및 리소스 정리...")
                self.page.quit()
            except Exception as e:
                print(f"종료 중 오류 (무시): {e}")
            finally:
                self.page = None

    def load_points(self):
        """save/point.json에서 좌표를 읽어옴. 실패 시 기본값 반환"""
        default_pts = {
            "click1": {"x": 535, "y": 375},
            "click2": {"x": 535, "y": 425}
        }
        
        pt_file = os.path.join("save", "point.json")
        if not os.path.exists(pt_file):
            return default_pts
            
        try:
            with open(pt_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if "click1" in data and "click2" in data:
                    return data
                return default_pts
        except:
            return default_pts

    def solve_cloudflare_gui(self):
            """Cloudflare 물리 우회"""
            try:
                if not self.page: return False

                current_title = self.page.title
                target_keywords = ["잠시만 기다리십시오", "Just a moment"]
                
                if not any(k in current_title for k in target_keywords):
                    return False

                print(f"🚨 Cloudflare 감지됨 ({current_title}). 물리 우회 시도...")

                target_win = None
                all_titles = gw.getAllTitles()
                
                for t in all_titles:
                    if ("잠시만 기다리십시오" in t or "Just a moment" in t) and "Chrome" in t:
                        wins = gw.getWindowsWithTitle(t)
                        if wins:
                            target_win = wins[0]
                            break
                
                if target_win:
                    if target_win.isMinimized: target_win.restore()
                    target_win.activate()
                    time.sleep(1.5)
                    target_win.maximize()
                    time.sleep(2.5)
                    
                    # [좌표 파일 로드]
                    points = self.load_points()
                    c1 = points["click1"]
                    c2 = points["click2"]

                    print(f"   -> 체크박스 클릭 시도... ({c1['x']}, {c1['y']})")
                    pyautogui.moveTo(c1['x'], c1['y'], duration=0.5) 
                    pyautogui.click() 
                    
                    time.sleep(1.5) 

                    print(f"   -> 체크박스 다른거 클릭 시도... ({c2['x']}, {c2['y']})")
                    pyautogui.moveTo(c2['x'], c2['y'], duration=0.5) 
                    pyautogui.click() 

                    time.sleep(5.5) 

                    target_win.restore()
                    time.sleep(1.5)
                    target_win.moveTo(1600, 1000)
                    target_win.resizeTo(800, 600)
                    
                    try:
                        if self.page.tabs_count > 1:
                            self.page.close_other_tabs() 
                    except:
                        pass

                    return True
                else:
                    return False

            except Exception as e:
                print(f"   -> [오류] GUI 우회 실패: {e}")
                return False

    def scan_and_check(self, regions, keyword_filename, fixed_targets=[], stop_event=None):
        final_results = []
        target_keywords = self.load_keywords(keyword_filename)
        
        self._init_browser()
        
        try:
            if self.page and self.page.tabs_count > 1:
                self.page.close_other_tabs()
        except: pass

        for region in regions:
            if stop_event and stop_event.is_set():
                print("🛑 중지 신호 감지! 스캔 루프를 탈출합니다.")
                break

            region_name = region['name']
            main_url = region['url']
            limit = region.get('limit', 10)
            
            print(f"{region_name} 스캔 중...")
            
            try:
                if stop_event and stop_event.is_set(): break

                if not self.page: self._init_browser()

                self.page.get(main_url, retry=2, interval=2)
                self.solve_cloudflare_gui()
                
                time.sleep(3.5) 
                
                links = self.page.eles('css:a[href*="/status/"], a[href*="/shougai/"], a[href*="/masalah/"]')
                
                count = 0
                visited_urls = set()
                found_services = []

                for link in links:
                    if count >= limit: break
                    if stop_event and stop_event.is_set(): break

                    href = link.attr('href')
                    if not href: continue

                    if not href.startswith("http"):
                        parsed_uri = urlparse(main_url)
                        domain = '{uri.scheme}://{uri.netloc}'.format(uri=parsed_uri)
                        full_url = domain + href
                    else:
                        full_url = href

                    if full_url in visited_urls: continue
                    
                    service_name = full_url.rstrip('/').split('/')[-1]
                    found_services.append({"name": service_name, "url": full_url})
                    visited_urls.add(full_url)
                    count += 1
                
                if count == 0:
                    print(f"   ⚠️ {region_name}: 수집된 링크가 없습니다. (Cloudflare 차단 의심)")

                if stop_event and stop_event.is_set(): break

                for svc in found_services:
                    if stop_event and stop_event.is_set(): break
                    
                    res = self.check_single_service(svc['name'], svc['url'], target_keywords, region_name)
                    res['group'] = region_name
                    final_results.append(res)
                    time.sleep(random.uniform(2.0, 3.5))

            except Exception as e:
                err_msg = str(e)
                if "与页面的连接已断开" in err_msg or "disconnected" in err_msg:
                    print(f"🛑 브라우저 연결이 종료되었습니다. ({region_name})")
                    break 
                
                print(f"[Error] {region_name} 처리 중 오류: {e}")
                
                if stop_event and stop_event.is_set(): break
                self._init_browser()

        if fixed_targets and not (stop_event and stop_event.is_set()):
            for item in fixed_targets:
                if stop_event and stop_event.is_set(): break
                
                try:
                    name = item['name']
                    url = item['url']
                    group = item.get('group', 'Fixed')
                    res = self.check_single_service(name, url, target_keywords, group)
                    res['group'] = group 
                    final_results.append(res)
                    time.sleep(random.uniform(2.0, 3.5))
                except Exception as e:
                    print(f"[Error] 고정 타겟 {item.get('name')} 오류: {e}")

        return final_results

    def check_single_service(self, name, url, keywords, group_name=""):
        data = {
            "name": name,
            "url": url,
            "status": "NODATA", 
            "alarm_trigger": False,
            "msg": "데이터 수집 실패",
            "error": None
        }
        
        try:
            if not self.page: raise Exception("Browser not alive")

            self.page.get(url, retry=2, interval=2)
            self.solve_cloudflare_gui()
            time.sleep(4)

            if not self.page.title: raise Exception("Page Load Failed")
            
            # =========================================================================
            # JS 내부 변수 및 CSS 변수 조회
            # =========================================================================
            detection_result = self.page.run_js("""
                return (function() {
                    // -----------------------------------------------------------
                    // [1] 기존 Downdetector UI / Legacy UI: 내부 변수 'window.DD.currentServiceProperties' 조회
                    // -----------------------------------------------------------
                    try {
                        if (window.DD && window.DD.currentServiceProperties) {
                            var s = window.DD.currentServiceProperties.status;
                            // JP에서 상태가 감지되면 즉시 리턴하여 아래 US 코드가 실행되지 않도록 함 (우선순위 보장)
                            if (s === 'danger') return 'CRITICAL';
                            if (s === 'warning') return 'WARNING';
                            if (s === 'success') return 'NORMAL';
                        }
                    } catch(e) {}

                    // -----------------------------------------------------------
                    // [2] 새로운 Downdetector UI / New UI: Recharts 그래프의 'stroke' 속성 CSS 변수 확인
                    // -----------------------------------------------------------
                    try {
                        // 사용자가 찾은 정확한 태그 및 클래스
                        const newGraph = document.querySelector('path.recharts-curve.recharts-area-curve');
                        if (newGraph) {
                            const strokeAttr = newGraph.getAttribute('stroke');
                            if (strokeAttr) {
                                // var(--color-dd-red) -> CRITICAL
                                if (strokeAttr.includes('color-dd-red')) return 'CRITICAL';
                                
                                // var(--color-dd-yellow) -> WARNING
                                if (strokeAttr.includes('color-dd-yellow')) return 'WARNING';
                                
                                // var(--color-chart-success) -> NORMAL (사용자가 찾아낸 값)
                                if (strokeAttr.includes('color-chart-success')) return 'NORMAL';
                            }
                        }
                    } catch(e) {}

                    // [3] 아무것도 발견되지 않음 (텍스트 분석으로 넘어감)
                    return 'CHECK_TEXT';
                })();
            """)

            # JS 분석 결과 처리
            if detection_result == 'CRITICAL':
                data["status"] = "CRITICAL"
                data["msg"] = "CRITICAL (System Detected)"
            elif detection_result == 'WARNING':
                data["status"] = "WARNING"
                data["msg"] = "WARNING (System Detected)"
            elif detection_result == 'NORMAL':
                data["status"] = "NORMAL"
                data["msg"] = "NORMAL"
            
            # 텍스트 분석
            else:
                header_ele = self.page.ele('css:div.indicator_status_message') or \
                             self.page.ele('css:h2') or \
                             self.page.ele('css:h1') or \
                             self.page.ele('text:User reports indicate')
                
                status_text = header_ele.text.strip().replace("\n", " ") if header_ele else ""
                lower_text = status_text.lower()
                
                # 1. success: "no current problems"
                if "no current problems" in lower_text:
                    data["status"] = "NORMAL"
                    data["msg"] = "NORMAL"
                
                # 2. warning: "possible problems"
                elif "possible problems" in lower_text:
                    data["status"] = "WARNING"
                    data["msg"] = f"WARNING: {status_text[:30]}"
                
                # 3. danger: "problems at"
                elif "problems at" in lower_text:
                    data["status"] = "CRITICAL"
                    data["msg"] = f"CRITICAL: {status_text[:30]}"
                
                # 4. 기타 언어 대응
                elif "障害が発生" in status_text:
                    data["status"] = "CRITICAL"
                    data["msg"] = f"CRITICAL: {status_text[:30]}"
                elif "起こり得る問題" in status_text:
                    data["status"] = "WARNING"
                    data["msg"] = f"WARNING: {status_text[:30]}"
                else:
                    data["status"] = "NORMAL"
                    data["msg"] = "NORMAL"

            # 알람
            is_target_keyword = False
            if keywords:
                for kw in keywords:
                    if kw.lower() in name.lower() or kw.lower() in url.lower():
                        is_target_keyword = True
                        break
            
            if data["status"] == "CRITICAL":
                if "Roaming" in group_name or "로밍" in group_name or is_target_keyword:
                    data["alarm_trigger"] = True
                    print(f"    🚨 {name}: CRITICAL (Alarm ON)")
                else:
                    data["alarm_trigger"] = False
                    print(f"    -> {name}: CRITICAL (No Alarm)")
            else:
                data["alarm_trigger"] = False

        except Exception as e:
            data["error"] = str(e)
            data["status"] = "NODATA"
            data["msg"] = "수집 오류"
            
        return data
