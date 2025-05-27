import os
import sys
import json
import threading
import subprocess
import tempfile
import shutil
import winreg
import psutil
import tkinter as tk
import tkinter.messagebox as messagebox
import requests
from PIL import Image, ImageTk
from io import BytesIO
from win32com.client import Dispatch

# 設定ファイルパス
CONFIG_PATH = os.path.join(os.path.dirname(__file__), 'config.json')

def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            messagebox.showerror("設定エラー", f"設定ファイルの読み込みに失敗しました:\n{e}")
        except Exception as e:
            messagebox.showerror("設定エラー", f"設定ファイルの読み込み中にエラーが発生しました:\n{e}")
    return {
        "skip_shortcut_prompt": False,
        "last_url": "",
        "repeat_time": 10,
        "time_unit": "分",
        "repeat_count": 1,
        "infinite_loop": True
    }


def save_config(config):
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showerror("設定エラー", f"設定ファイルの保存に失敗しました:\n{e}")


def get_chrome_path():
    for key in (
        r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\chrome.exe",
        r"SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\chrome.exe",
    ):
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key) as reg:
                return winreg.QueryValue(reg, None)
        except FileNotFoundError:
            continue
    for path in (
        r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        r"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
    ):
        if os.path.exists(path):
            return path
    return None


def kill_process_tree(pid):
    """プロセスとその子プロセスを強制終了する（改良版）"""
    try:
        parent = psutil.Process(pid)
        children = parent.children(recursive=True)
        
        # 子プロセスを終了
        for child in children:
            try:
                child.terminate()
            except:
                pass
                
        # 親プロセスを終了
        try:
            parent.terminate()
        except:
            pass
            
        # 終了を待機（短時間）
        gone, alive = psutil.wait_procs(children + [parent], timeout=3)
        
        # 残っているプロセスを強制終了
        for p in alive:
            try:
                p.kill()
            except:
                pass
    except Exception as e:
        print(f"プロセス終了エラー: {e}")

def create_desktop_shortcut():
    try:
        desktop = os.path.join(os.environ.get('USERPROFILE', ''), 'Desktop')
        shortcut_path = os.path.join(desktop, 'YouTubeRepeater.lnk')
        if os.path.exists(shortcut_path):
            return True
        target = sys.executable
        script = os.path.abspath(__file__)
        icon = os.path.join(os.path.dirname(script), 'app.ico')
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortcut(shortcut_path)
        shortcut.TargetPath = target
        shortcut.Arguments = f'"{script}"'
        shortcut.WorkingDirectory = os.path.dirname(script)
        shortcut.IconLocation = icon if os.path.exists(icon) else target
        shortcut.Save()
        return True
    except Exception as e:
        messagebox.showerror("ショートカットエラー", f"デスクトップショートカットの作成に失敗しました:\n{e}")
        return False

class App:
    def __init__(self, root):
        self.config = load_config()
        self.root = root
        self.root.title("YouTube繰り返しビューア")
        # アイコン設定
        icon_path = os.path.join(os.path.dirname(__file__), 'app.ico')
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception:
                img = tk.PhotoImage(file=icon_path)
                self.root.iconphoto(True, img)
        self.stop_event = threading.Event()

        # UI構築
        tk.Label(root, text="YouTube URL:").grid(row=0, column=0, sticky="e")
        self.entry_url = tk.Entry(root, width=50)
        self.entry_url.grid(row=0, column=1, columnspan=3, pady=4)
        # 前回のURLを設定
        if self.config.get("last_url"):
            self.entry_url.insert(0, self.config.get("last_url"))

        tk.Label(root, text="繰り返し時間:").grid(row=1, column=0, sticky="e")
        self.spin_time = tk.Spinbox(root, from_=1, to=999999, width=6)
        self.spin_time.grid(row=1, column=1, pady=4)
        # 前回の時間を設定
        self.spin_time.delete(0, tk.END)
        self.spin_time.insert(0, str(self.config.get("repeat_time", 10)))
        
        self.unit_var = tk.StringVar(value=self.config.get("time_unit", "分"))
        tk.OptionMenu(root, self.unit_var, "秒", "分", "時間", "日").grid(row=1, column=2)

        tk.Label(root, text="回数:").grid(row=2, column=0, sticky="e")
        self.spin_count = tk.Spinbox(root, from_=1, to=999999, width=6)
        self.spin_count.grid(row=2, column=1, pady=4)
        # 前回の回数を設定
        self.spin_count.delete(0, tk.END)
        self.spin_count.insert(0, str(self.config.get("repeat_count", 1)))
        
        self.infinite_var = tk.BooleanVar(value=self.config.get("infinite_loop", True))
        tk.Checkbutton(root, text="無限繰り返し", variable=self.infinite_var,
                      command=self.toggle_infinite).grid(row=2, column=2)
        
        # 無限繰り返しの状態に合わせて回数入力欄の状態を設定
        if self.infinite_var.get():
            self.spin_count.config(state="disabled")
        
        tk.Label(root, text="再生中タイトル:").grid(row=3, column=0, sticky="e")
        self.label_title = tk.Label(root, text="-")
        self.label_title.grid(row=3, column=1, columnspan=3, sticky="w")
        self.thumbnail_label = tk.Label(root)
        self.thumbnail_label.grid(row=4, column=0, columnspan=4, pady=4)

        tk.Label(root, text="残り時間:").grid(row=5, column=0, sticky="e")
        self.label_timer = tk.Label(root, text="--:--:--")
        self.label_timer.grid(row=5, column=1, columnspan=3, sticky="w")

        # ボタンをフレームに配置して横に並べる
        btn_frame = tk.Frame(root)
        btn_frame.grid(row=6, column=0, columnspan=4, pady=10)
        
        self.btn_start = tk.Button(btn_frame, text="実行", width=10, command=self.on_start)
        self.btn_start.pack(side=tk.LEFT, padx=5)
        
        self.btn_stop = tk.Button(btn_frame, text="停止", width=10, command=self.on_stop, state="disabled")
        self.btn_stop.pack(side=tk.LEFT, padx=5)

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # 初期URLがあれば、起動時に動画情報を取得
        initial_url = self.config.get("last_url", "").strip()
        if initial_url and (initial_url.startswith(("http://", "https://")) and 
                           ("youtube.com" in initial_url or "youtu.be" in initial_url)):
            # 別スレッドで情報取得（UIをブロックしないため）
            threading.Thread(target=self.load_initial_video_info, args=(initial_url,), daemon=True).start()
    
    def load_initial_video_info(self, url):
        """アプリ起動時に初期動画情報を取得"""
        try:
            # ネットワーク接続確認
            try:
                requests.head("https://www.youtube.com", timeout=5)
            except requests.RequestException:
                # 接続エラーは静かに無視（起動時なので）
                return
                
            info = self.fetch_video_info(url, show_error=False)
            if info and info.get("title") != "-":
                self.root.after(0, lambda i=info: self.display_video_info(i))
        except Exception as e:
            # 起動時のエラーは静かに無視
            print(f"初期動画情報取得エラー: {e}")

    def toggle_infinite(self):
        state = "disabled" if self.infinite_var.get() else "normal"
        self.spin_count.config(state=state)

    def format_time(self, secs):
        h, rem = divmod(int(secs), 3600)
        m, s = divmod(rem, 60)
        return f"{h:02d}:{m:02d}:{s:02d}"

    def on_start(self):
        url = self.entry_url.get().strip()
        if not url:
            messagebox.showerror("入力エラー", "URLを入力してください。")
            return
        if not url.startswith(("http://", "https://")):
            messagebox.showerror("入力エラー", "有効なURLを入力してください。")
            return
        if not ("youtube.com" in url or "youtu.be" in url):
            messagebox.showerror("入力エラー", "YouTubeのURLを入力してください。")
            return
        
        try:
            t = float(self.spin_time.get())
        except ValueError:
            messagebox.showerror("入力エラー", "繰り返し時間は数値で入力してください。")
            return
        interval = t * {"秒":1, "分":60, "時間":3600, "日":86400}[self.unit_var.get()]

        if not self.infinite_var.get():
            try:
                count = int(self.spin_count.get())
                if count < 1:
                    raise ValueError()
            except ValueError:
                messagebox.showerror("入力エラー", "回数は1以上の整数で入力してください。")
                return
        else:
            count = None

        self.stop_event.clear()  # 停止フラグをリセット
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")  # 停止ボタンを有効化
        self.label_timer.config(text=self.format_time(interval))
        threading.Thread(target=self.run_loop, args=(url, interval, count), daemon=True).start()

    def fetch_video_info(self, url, show_error=True):
        info = {"title": "-", "thumbnail": None}
        try:
            resp = requests.get(f"https://www.youtube.com/oembed?url={url}&format=json", timeout=5)
            data = resp.json()
            info["title"] = data.get("title", "-")
            thumb = data.get("thumbnail_url")
            if thumb:
                img = Image.open(BytesIO(requests.get(thumb, timeout=5).content))
                img = img.resize((256,144), Image.Resampling.LANCZOS)
                info["thumbnail"] = ImageTk.PhotoImage(img)
        except Exception as e:
            print(f"Error fetching video info: {e}")
            if show_error:
                messagebox.showerror("エラー", f"動画情報の取得に失敗しました:\n{e}")
        return info

    def run_loop(self, url, interval, count):
        iteration = 0
        chrome_proc = None
        temp_dir = None
        
        try:
            # 一時ディレクトリを作成（Chromeのユーザーデータ用）
            temp_dir = tempfile.mkdtemp(prefix="youtube_repeater_")
            
            while not self.stop_event.is_set() and (count is None or iteration < count):
                # ネットワーク接続確認
                try:
                    requests.head("https://www.youtube.com", timeout=5)
                except requests.RequestException:
                    self.root.after(0, lambda: messagebox.showerror("ネットワークエラー", 
                        "YouTubeサーバーに接続できません。ネットワーク接続を確認してください。"))
                    break
                    
                info = self.fetch_video_info(url)
                self.root.after(0, lambda i=info: self.display_video_info(i))

                play_url = url
                if 'youtube.com/watch' in url:
                    sep = '&' if '?' in url else '?'
                    play_url = f"{url}{sep}autoplay=1"
            
                # 独立したChromeインスタンスとして起動
                cmd = [
                    CHROME_PATH,
                    "--new-window",
                    "--autoplay-policy=no-user-gesture-required",
                    f"--user-data-dir={temp_dir}",
                    "--no-first-run",
                    "--no-default-browser-check",
                    "--disable-sync",
                    "--disable-extensions",
                    play_url
                ]
                
                try:
                    # 前回のプロセスが残っていれば終了
                    if chrome_proc and chrome_proc.poll() is None:
                        kill_process_tree(chrome_proc.pid)
                    
                    chrome_proc = subprocess.Popen(cmd)
                except FileNotFoundError:
                    self.root.after(0, lambda: messagebox.showerror("ブラウザエラー", 
                        "Chromeが見つかりません。パスが正しいか確認してください。"))
                    break
                except PermissionError:
                    self.root.after(0, lambda: messagebox.showerror("権限エラー", 
                        "Chromeを起動する権限がありません。"))
                    break
                except Exception as e:
                    self.root.after(0, lambda: messagebox.showerror("実行エラー", f"Chrome起動に失敗しました:\n{e}"))
                    break

                remaining = interval
                while remaining > 0 and not self.stop_event.is_set():
                    self.root.after(0, lambda t=remaining: self.label_timer.config(text=self.format_time(t)))
                    if self.stop_event.wait(1):
                        break
                    remaining -= 1

                # Chromeプロセスを確実に終了
                if chrome_proc and chrome_proc.poll() is None:
                    try:
                        kill_process_tree(chrome_proc.pid)
                        # 終了を確認
                        chrome_proc.wait(timeout=5)
                    except Exception as e:
                        print(f"プロセス終了エラー: {e}")
            
                iteration += 1
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("予期せぬエラー", f"実行中に予期せぬエラーが発生しました:\n{e}"))
        finally:
            # 最終的にChromeプロセスを確実に終了
            if chrome_proc and chrome_proc.poll() is None:
                try:
                    kill_process_tree(chrome_proc.pid)
                except:
                    pass
        
            # 一時ディレクトリを削除
            if temp_dir and os.path.exists(temp_dir):
                try:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                except:
                    pass
                
            self.root.after(0, lambda: self.btn_start.config(state="normal"))
            self.root.after(0, lambda: self.btn_stop.config(state="disabled"))
            self.root.after(0, lambda: self.label_timer.config(text="--:--:--"))

    def display_video_info(self, info):
        self.label_title.config(text=info.get("title", "-"))
        thumb = info.get("thumbnail")
        if thumb:
            self.thumbnail_label.config(image=thumb)
            self.thumbnail_label.image = thumb
        else:
            self.thumbnail_label.config(image="")
            self.thumbnail_label.image = None

    def on_close(self):
        # 現在の設定を保存
        self.config["last_url"] = self.entry_url.get().strip()
        try:
            self.config["repeat_time"] = float(self.spin_time.get())
        except ValueError:
            self.config["repeat_time"] = 5  # デフォルト値
    
        self.config["time_unit"] = self.unit_var.get()
        
        try:
            self.config["repeat_count"] = int(self.spin_count.get())
        except ValueError:
            self.config["repeat_count"] = 1  # デフォルト値
    
        self.config["infinite_loop"] = self.infinite_var.get()
        
        # 設定を保存
        save_config(self.config)
        
        # ショートカット作成確認
        desktop = os.path.join(os.environ.get('USERPROFILE', ''), 'Desktop')
        shortcut_path = os.path.join(desktop, 'YouTubeRepeater.lnk')
        if not os.path.exists(shortcut_path) and not self.config.get("skip_shortcut_prompt", False):
            res = messagebox.askyesnocancel(
                "ショートカット作成",
                "デスクトップにショートカットがありません。作成しますか？\nキャンセルで今後通知しません。"
            )
            if res is True:
                create_desktop_shortcut()
            elif res is None:
                self.config["skip_shortcut_prompt"] = True
                save_config(self.config)
        self.stop_event.set()
        self.root.destroy()

    def on_stop(self):
        """停止ボタンがクリックされたときの処理"""
        self.stop_event.set()
        self.btn_stop.config(state="disabled")
        # 停止中の表示
        self.label_timer.config(text="停止中...")

if __name__ == "__main__":
    CHROME_PATH = get_chrome_path()
    if not CHROME_PATH:
        tk.Tk().withdraw()
        messagebox.showerror("エラー", "Google Chromeが見つかりません。インストールをご確認ください。")
        sys.exit(1)

    root = tk.Tk()
    app = App(root)
    root.mainloop()