import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import csv
import os
import glob
from datetime import datetime

# --- ローカル保存用の設定ファイル名 ---
LOCAL_SETTING_FILE = "admin_settings.json"
DEFAULT_CONFIG_NAME = "config.json"

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

class AdminApp:
    def __init__(self, root):
        self.root = root
        self.root.title("進捗管理者パネル (Ultimate Distributed Edition)")
        self.root.geometry("850x900")

        # 1. 共有設定ファイルのパス取得とロード
        if not self.init_config_path():
            self.root.destroy()
            return

        self.auto_refresh = tk.BooleanVar(value=True)
        self.setup_ui()
        self.refresh_monitor()

    def init_config_path(self):
        """初回起動時のパス選択と設定の読み込み"""
        self.config_path = get_local_config_path()
        
        if not self.config_path or not os.path.exists(self.config_path):
            messagebox.showinfo("初期設定", "共有フォルダにある管理用設定ファイル (config.json) を選択してください。")
            self.config_path = filedialog.askopenfilename(
                title="共有config.jsonを選択",
                filetypes=[("JSON files", "*.json")]
            )
            if not self.config_path:
                messagebox.showerror("エラー", "設定ファイルがないと管理機能は使用できません。")
                return False
            save_local_config_path(self.config_path)

        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                self.config = json.load(f)
                # 必須キーの補完
                if "admins" not in self.config: self.config["admins"] = []
                if "refresh_ms" not in self.config: self.config["refresh_ms"] = 3000
            return True
        except Exception as e:
            messagebox.showerror("読込エラー", f"設定の読み込みに失敗しました:\n{e}")
            return False

    def setup_ui(self):
        # メインスクロール
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # --- 現在の接続先情報 ---
        info_frame = ttk.Frame(self.scrollable_frame)
        info_frame.pack(fill="x", padx=20, pady=5)
        ttk.Label(info_frame, text=f"接続先: {self.config_path}", font=("Arial", 8), foreground="gray").pack(side="left")
        ttk.Button(info_frame, text="接続先変更", command=self.reset_connection, style="Small.TButton").pack(side="right")

        # --- 1. モニターセクション ---
        mon_head = ttk.Frame(self.scrollable_frame)
        mon_head.pack(fill="x", padx=20, pady=10)
        ttk.Label(mon_head, text="【稼働状況モニター】", font=("Arial", 12, "bold")).pack(side="left")
        
        refresh_ctrl = ttk.Frame(mon_head)
        refresh_ctrl.pack(side="right")
        ttk.Label(refresh_ctrl, text="間隔(ms):").pack(side="left")
        self.refresh_entry = ttk.Entry(refresh_ctrl, width=6)
        self.refresh_entry.insert(0, str(self.config["refresh_ms"]))
        self.refresh_entry.pack(side="left", padx=5)
        ttk.Checkbutton(mon_head, text="リアルタイム更新ON", variable=self.auto_refresh).pack(side="right", padx=10)

        self.tree = ttk.Treeview(self.scrollable_frame, columns=("User", "Task", "Time", "IsAdmin"), show="headings", height=8)
        self.tree.heading("User", text="利用者名"); self.tree.heading("Task", text="タスク")
        self.tree.heading("Time", text="経過"); self.tree.heading("IsAdmin", text="権限")
        self.tree.column("IsAdmin", width=80, anchor="center")
        self.tree.pack(fill="x", padx=20)

        # --- 2. ログ保存先（共有フォルダ） ---
        ttk.Separator(self.scrollable_frame, orient="horizontal").pack(fill="x", pady=20)
        path_frame = ttk.Frame(self.scrollable_frame)
        path_frame.pack(fill="x", padx=20)
        ttk.Label(path_frame, text="ログ出力先:").pack(side="left")
        self.path_entry = ttk.Entry(path_frame)
        self.path_entry.insert(0, self.config["save_path"])
        self.path_entry.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(path_frame, text="参照", command=self.select_log_path).pack(side="left")

        # --- 3. ユーザー割当 ---
        ttk.Label(self.scrollable_frame, text="【ユーザー名 & 管理者 割当】", font=("Arial", 10, "bold")).pack(pady=(15, 0))
        assign_frame = ttk.Frame(self.scrollable_frame)
        assign_frame.pack(fill="x", padx=20, pady=5)
        
        ttk.Label(assign_frame, text="未割当PC:").grid(row=0, column=0, sticky="w")
        self.pc_combo = ttk.Combobox(assign_frame, values=self.get_detected_pcs())
        self.pc_combo.grid(row=0, column=1, padx=5, sticky="ew")
        ttk.Label(assign_frame, text="表示名:").grid(row=0, column=2, sticky="w")
        self.name_entry = ttk.Entry(assign_frame)
        self.name_entry.grid(row=0, column=3, padx=5, sticky="ew")
        
        self.is_admin_var = tk.BooleanVar()
        ttk.Checkbutton(assign_frame, text="管理者権限", variable=self.is_admin_var).grid(row=0, column=4, padx=5)
        ttk.Button(assign_frame, text="リストに追加", command=self.add_mapping).grid(row=0, column=5, padx=5)
        assign_frame.columnconfigure((1, 3), weight=1)

        self.mapping_text = tk.Text(self.scrollable_frame, height=5)
        self.refresh_mapping_display()
        self.mapping_text.pack(fill="x", padx=20, pady=5)
        ttk.Button(self.scrollable_frame, text="編集内容を反映", command=self.sync_mapping_from_text).pack(pady=2)

        # --- 4. タスクリスト ---
        ttk.Separator(self.scrollable_frame, orient="horizontal").pack(fill="x", pady=20)
        ttk.Label(self.scrollable_frame, text="【タスクリスト管理 (CSV連携)】", font=("Arial", 10, "bold")).pack()
        t_btn_frame = ttk.Frame(self.scrollable_frame)
        t_btn_frame.pack(pady=5)
        ttk.Button(t_btn_frame, text="テンプレート(CSV)出力", command=self.export_tasks).pack(side="left", padx=5)
        ttk.Button(t_btn_frame, text="CSVから一括読込", command=self.import_tasks).pack(side="left", padx=5)

        self.task_text = tk.Text(self.scrollable_frame, height=6)
        self.task_text.insert("1.0", "\n".join(self.config["task_list"]))
        self.task_text.pack(fill="x", padx=20, pady=5)

        # --- 5. 最終保存 ---
        action_frame = ttk.Frame(self.scrollable_frame)
        action_frame.pack(pady=30)
        ttk.Button(action_frame, text="共有設定を保存(確定)", command=self.save_to_shared_config, style="Accent.TButton").pack(side="left", padx=10)
        ttk.Button(action_frame, text="全ログデータを集計出力", command=self.aggregate_data).pack(side="left", padx=10)

    # --- ロジック ---
    def reset_connection(self):
        if messagebox.askyesno("確認", "共有設定ファイルの参照先を変更しますか？\n(アプリを再起動します)"):
            if os.path.exists(LOCAL_SETTING_FILE): os.remove(LOCAL_SETTING_FILE)
            self.root.destroy()

    def get_detected_pcs(self):
        path = self.config["save_path"]
        if not os.path.exists(path): return []
        return [os.path.basename(f).replace(".csv", "") for f in glob.glob(os.path.join(path, "*.csv"))]

    def add_mapping(self):
        pc, name = self.pc_combo.get().strip(), self.name_entry.get().strip()
        if not pc or not name: return
        self.config["user_mapping"][pc] = name
        if self.is_admin_var.get():
            if pc not in self.config["admins"]: self.config["admins"].append(pc)
        else:
            if pc in self.config["admins"]: self.config["admins"].remove(pc)
        self.refresh_mapping_display()
        self.name_entry.delete(0, "end")

    def sync_mapping_from_text(self):
        new_map, new_adm = {}, []
        for line in self.mapping_text.get("1.0", "end-1c").splitlines():
            if ":" in line:
                main = line.split("[")[0] if "[" in line else line
                pc, name = [x.strip() for x in main.split(":", 1)]
                new_map[pc] = name
                if "[管理者]" in line: new_adm.append(pc)
        self.config["user_mapping"], self.config["admins"] = new_map, new_adm
        messagebox.showinfo("反映", "メモリ上に一時反映しました。最後に保存ボタンを押してください。")

    def refresh_mapping_display(self):
        self.mapping_text.delete("1.0", "end")
        lines = [f"{pc} : {name} {'[管理者]' if pc in self.config['admins'] else ''}" 
                 for pc, name in self.config["user_mapping"].items()]
        self.mapping_text.insert("1.0", "\n".join(lines))

    def refresh_monitor(self):
        if self.auto_refresh.get():
            for item in self.tree.get_children(): self.tree.delete(item)
            path = self.config["save_path"]
            if os.path.exists(path):
                for f in glob.glob(os.path.join(path, "*.csv")):
                    pc = os.path.basename(f).replace(".csv", "")
                    name = self.config["user_mapping"].get(pc, pc)
                    adm = "★" if pc in self.config["admins"] else ""
                    try:
                        with open(f, "r", encoding="utf-8-sig") as csv_f:
                            r = list(csv.reader(csv_f))
                            if len(r) > 1 and not r[-1][2]:
                                start = datetime.strptime(r[-1][1], "%Y-%m-%d %H:%M:%S")
                                elap = str(datetime.now() - start).split(".")[0]
                                self.tree.insert("", "end", values=(name, r[-1][0], elap, adm))
                    except: pass
        try:
            ms = int(self.refresh_entry.get())
        except: ms = 3000
        self.root.after(ms, self.refresh_monitor)

    def select_log_path(self):
        p = filedialog.askdirectory()
        if p:
            self.path_entry.delete(0, "end"); self.path_entry.insert(0, p)
            self.pc_combo["values"] = self.get_detected_pcs()

    def save_to_shared_config(self):
        """共有フォルダのconfig.jsonを上書きする"""
        self.config["save_path"] = self.path_entry.get()
        self.config["task_list"] = [t for t in self.task_text.get("1.0", "end-1c").splitlines() if t.strip()]
        try:
            self.config["refresh_ms"] = int(self.refresh_entry.get())
        except: self.config["refresh_ms"] = 3000
        
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("成功", f"共有設定を保存しました:\n{self.config_path}")
        except Exception as e:
            messagebox.showerror("保存失敗", f"共有ファイルへの書き込み権限がない可能性があります:\n{e}")

    def export_tasks(self):
        p = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="task_list.csv")
        if p:
            with open(p, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["task_name"])
                for t in self.task_text.get("1.0", "end-1c").splitlines():
                    if t.strip(): writer.writerow([t.strip()])

    def import_tasks(self):
        p = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if p:
            with open(p, "r", encoding="utf-8-sig") as f:
                reader = csv.reader(f); next(reader)
                tasks = [r[0] for r in reader if r]
            self.task_text.delete("1.0", "end"); self.task_text.insert("1.0", "\n".join(tasks))

    def aggregate_data(self):
        save_file = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=f"aggregate_{datetime.now().strftime('%Y%m%d')}.csv")
        if not save_file: return
        all_files = glob.glob(os.path.join(self.config["save_path"], "*.csv"))
        with open(save_file, "w", encoding="utf-8-sig", newline="") as out_f:
            writer = csv.writer(out_f)
            writer.writerow(["表示名", "PC名", "管理者", "タスク", "開始", "終了"])
            for f in all_files:
                pc = os.path.basename(f).replace(".csv", "")
                name = self.config["user_mapping"].get(pc, pc)
                adm = "YES" if pc in self.config["admins"] else "NO"
                try:
                    with open(f, "r", encoding="utf-8-sig") as in_f:
                        r = csv.reader(in_f); next(r)
                        for row in r: writer.writerow([name, pc, adm] + row)
                except: pass
        messagebox.showinfo("成功", "集計完了")

if __name__ == "__main__":
    root = tk.Tk()
    app = AdminApp(root)
    root.mainloop()