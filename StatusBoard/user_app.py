import customtkinter as ctk
import json
import os
import csv
from datetime import datetime
import socket
from tkinter import messagebox, filedialog

# --- ローカル保存用の設定ファイル名 ---
LOCAL_SETTING_FILE = "user_settings.json"

def get_local_config_path():
    """保存された共有設定(config.json)のパスを読み込む"""
    if os.path.exists(LOCAL_SETTING_FILE):
        try:
            with open(LOCAL_SETTING_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("config_path")
        except: return None
    return None

def save_local_config_path(path):
    """共有設定(config.json)のパスをローカルに保存する"""
    with open(LOCAL_SETTING_FILE, "w", encoding="utf-8") as f:
        json.dump({"config_path": path}, f, indent=4, ensure_ascii=False)

class UserApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # 1. 設定ファイルのパス取得
        self.config_path = get_local_config_path()
        
        if not self.config_path or not os.path.exists(self.config_path):
            messagebox.showinfo("初期設定", "管理者の設定ファイル (config.json) を選択してください。")
            self.config_path = filedialog.askopenfilename(
                title="config.jsonを選択",
                filetypes=[("JSON files", "*.json")]
            )
            if not self.config_path:
                self.destroy()
                return
            save_local_config_path(self.config_path)

        # 2. 設定の読み込み
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                self.config = json.load(f)
        except Exception as e:
            messagebox.showerror("エラー", f"設定の読み込みに失敗しました:\n{e}")
            self.destroy()
            return

        # 3. 基本情報のセットアップ
        self.pc_name = socket.gethostname()
        save_path = self.config.get("save_path", "logs")
        self.log_file = os.path.join(save_path, f"{self.pc_name}.csv")
        
        # 状態管理
        self.current_task = None
        self.start_time = None

        # 4. UI初期化
        self.setup_ui()
        self.init_log_file()
        
        # 5. タイマーループ開始
        self.update_timer()

    def setup_ui(self):
        self.title("進捗入力くん")
        self.geometry("400x450")
        ctk.set_appearance_mode("dark")

        # ヘッダー情報
        self.pc_label = ctk.CTkLabel(self, text=f"PC: {self.pc_name}", font=("Arial", 12))
        self.pc_label.pack(pady=5)
        
        self.status_label = ctk.CTkLabel(self, text="未稼働", font=("Arial", 18, "bold"), text_color="red")
        self.status_label.pack(pady=5)

        # 【重要】作成と配置を分けることでAttributeErrorを回避
        self.timer_label = ctk.CTkLabel(self, text="0:00:00", font=("Arial", 28))
        self.timer_label.pack(pady=5)

        # タスク選択
        ctk.CTkLabel(self, text="実行するタスクを選択:").pack(pady=5)
        task_list = self.config.get("task_list", ["タスクが設定されていません"])
        self.task_combo = ctk.CTkComboBox(self, values=task_list, width=280)
        self.task_combo.pack(pady=10)

        # メインアクション
        self.start_btn = ctk.CTkButton(
            self, text="▶ タスク開始 / 切替", 
            command=self.switch_task, 
            fg_color="#2E8B57", hover_color="#1E6B47", width=200, height=40
        )
        self.start_btn.pack(pady=15)

        ctk.CTkLabel(self, text="──────────────────────────", text_color="gray").pack(pady=5)

        # 休憩・終了
        exit_frame = ctk.CTkFrame(self, fg_color="transparent")
        exit_frame.pack(fill="x", padx=40, pady=10)

        self.break_btn = ctk.CTkButton(exit_frame, text="☕ 休憩入り", command=self.start_break, fg_color="#D2691E", width=140)
        self.break_btn.pack(side="left", padx=5, expand=True)

        self.finish_btn = ctk.CTkButton(exit_frame, text="🏁 業務終了", command=self.finish_work, fg_color="#B22222", width=140)
        self.finish_btn.pack(side="left", padx=5, expand=True)

        # 設定変更ボタン
        self.reset_btn = ctk.CTkButton(
            self, text="設定ファイルを再選択", 
            command=self.reset_config_path, 
            fg_color="gray", width=150, height=20, font=("Arial", 10)
        )
        self.reset_btn.pack(pady=20)

    def init_log_file(self):
        try:
            os.makedirs(os.path.dirname(self.log_file), exist_ok=True)
            if not os.path.exists(self.log_file):
                with open(self.log_file, "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["タスク名", "開始時間", "終了時間"])
        except Exception as e:
            messagebox.showerror("接続エラー", f"共有フォルダにアクセスできません:\n{e}")

    def save_log(self, task, start, end):
        try:
            end_str = end.strftime("%Y-%m-%d %H:%M:%S") if end else ""
            with open(self.log_file, "a", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f)
                writer.writerow([task, start.strftime("%Y-%m-%d %H:%M:%S"), end_str])
        except Exception as e:
            print(f"Log Save Error: {e}")

    def stop_current_task(self):
        """現在稼働中のタスクがあればCSVに書き出して終了する"""
        if self.current_task and self.start_time:
            # ここで初めてCSVに1行書き込まれる
            self.save_log(self.current_task, self.start_time, datetime.now())
            # 状態をリセット
            self.current_task = None
            self.start_time = None

    def switch_task(self):
        new_task = self.task_combo.get()
        if not new_task or new_task == "タスクが設定されていません":
            return
            
        # 1. 今までやってたタスクがあれば、その「終了」をログに記録する
        self.stop_current_task()
        
        # 2. 新しいタスクの状態をセット（ここではまだCSVには書かない！）
        self.current_task = new_task
        self.start_time = datetime.now()
        
        # 3. UIの表示だけ即座に更新する
        self.status_label.configure(text=f"稼働中: {self.current_task}", text_color="#00FFFF")
        self.refresh_timer_display()

    def refresh_timer_display(self):
        """1秒待たずに、現在の経過時間を即座にラベルに反映する"""
        if self.current_task and self.start_time:
            elapsed = datetime.now() - self.start_time
            time_str = str(elapsed).split(".")[0]
            self.timer_label.configure(text=time_str)

    def update_timer(self):
        """1秒ごとに実行されるループ"""
        # 表示を更新する
        self.refresh_timer_display()
        
        # 次の1秒後を予約
        self.after(1000, self.update_timer)

    def start_break(self):
        if self.current_task == "休憩": return
        self.stop_current_task()
        self.current_task = "休憩"
        self.start_time = datetime.now()
        self.save_log(self.current_task, self.start_time, None)
        self.status_label.configure(text="休憩中", text_color="#FFA500")
        self.task_combo.set("休憩")

    def finish_work(self):
        """業務終了ボタン"""
        if not self.current_task: return
        if messagebox.askyesno("業務終了", "本日の業務を終了しますか？"):
            self.stop_current_task() # ここで最後のタスクが記録される
            self.status_label.configure(text="業務終了", text_color="red")
            self.timer_label.configure(text="0:00:00")

    def update_timer(self):
        """1秒ごとに実行されるループ"""
        if self.current_task and self.start_time:
            elapsed = datetime.now() - self.start_time
            # timedeltaの文字列表現 (H:MM:SS) を取得
            time_str = str(elapsed).split(".")[0]
            
            # もし 0:00:00.123 のようにドットが残る場合の対策として
            # 常に「時:分:秒」の形になるようにセット
            self.timer_label.configure(text=time_str)
        
        self.after(1000, self.update_timer)

    def reset_config_path(self):
        if messagebox.askyesno("確認", "設定ファイルの参照先を変更しますか？"):
            if os.path.exists(LOCAL_SETTING_FILE):
                os.remove(LOCAL_SETTING_FILE)
            self.destroy()
            messagebox.showinfo("完了", "設定をリセットしました。再起動してください。")

if __name__ == "__main__":
    app = UserApp()
    # 起動に失敗（パス未選択など）した場合はmainloopに入らない
    if hasattr(app, "winfo_exists") and app.winfo_exists():
        app.mainloop()
