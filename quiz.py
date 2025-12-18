import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import dataclasses
from enum import Enum
from typing import List, Dict, Optional, Type

# --- 1. バックエンド・ロジック (クイズルールと計算) ---

class PlayerStatus(Enum):
    PLAYING = "回答中"
    WIN = "WIN (抜)"
    LOSE = "LOSE (失)"

@dataclasses.dataclass
class Player:
    id: int
    name: str
    organization: str
    
    score: int = 0
    wrong_count: int = 0
    status: PlayerStatus = PlayerStatus.PLAYING
    
    def reset(self):
        self.score = 0
        self.wrong_count = 0
        self.status = PlayerStatus.PLAYING

class QuizRule:
    """ルールの基底クラス"""
    def __init__(self, name, win_score=None, lose_wrong=None):
        self.name = name
        self.win_score = win_score
        self.lose_wrong = lose_wrong

    def on_correct(self, player: Player):
        if player.status != PlayerStatus.PLAYING: return
        self._calc_correct(player)
        self._check_status(player)

    def on_wrong(self, player: Player):
        if player.status != PlayerStatus.PLAYING: return
        self._calc_wrong(player)
        self._check_status(player)

    def _calc_correct(self, player: Player):
        player.score += 1

    def _calc_wrong(self, player: Player):
        pass # 基本は変動なし

    def _check_status(self, player: Player):
        if self.win_score and player.score >= self.win_score:
            player.status = PlayerStatus.WIN
        if self.lose_wrong and player.wrong_count >= self.lose_wrong:
            player.status = PlayerStatus.LOSE

# 具体的なルール実装
class RuleNbyM(QuizRule): # 5○2×など
    def __init__(self, n=5, m=2):
        super().__init__(f"{n}○{m}×", win_score=n, lose_wrong=m)
    def _calc_wrong(self, player: Player):
        player.wrong_count += 1

class RuleUpDown(QuizRule): # 10 Up-Down
    def __init__(self, target=10):
        super().__init__("10 Up-Down", win_score=target, lose_wrong=None)
    def _calc_wrong(self, player: Player):
        player.score = 0 # 誤答で0点に戻る
        player.wrong_count += 1

class RuleSwedish(QuizRule): # Swedish 10
    def __init__(self, target=10):
        super().__init__("Swedish 10", win_score=target, lose_wrong=None)
    def _calc_wrong(self, player: Player):
        player.wrong_count += 1
        # スウェーデン方式: 誤答数に応じてペナルティが変わる簡易実装
        # 例: 1×→-1, 2×→-2... (ここでは単純に -誤答数 とする)
        player.score -= player.wrong_count

class QuizSystem:
    def __init__(self):
        self.roster: Dict[int, Player] = {}
        self.current_players: List[Player] = []
        self.rule: QuizRule = RuleNbyM(5, 2)
        self.rules_available = {
            "2R: 5○2×": RuleNbyM(5, 2),
            "2R: 7○3×": RuleNbyM(7, 3),
            "3R: Up-Down": RuleUpDown(10),
            "3R: Swedish 10": RuleSwedish(10),
            "F: 10○4×": RuleNbyM(10, 4)
        }

    def load_roster(self, filepath):
        try:
            df = pd.read_csv(filepath)
            # CSVのカラム名に柔軟に対応
            col_id = next((c for c in df.columns if '番号' in c), None)
            col_name = next((c for c in df.columns if '名' in c and '表示' in c), '表示名') # 表示名優先
            if not col_name: col_name = next((c for c in df.columns if '姓' in c), None)
            col_org = next((c for c in df.columns if '所属' in c), '所属')

            if col_id:
                for _, row in df.iterrows():
                    pid = int(row[col_id]) if pd.notna(row[col_id]) else 0
                    name = str(row[col_name]) if col_name and pd.notna(row[col_name]) else f"Player{pid}"
                    org = str(row[col_org]) if col_org and pd.notna(row[col_org]) else ""
                    if pid > 0:
                        self.roster[pid] = Player(pid, name, org)
            return True
        except Exception as e:
            print(f"Error: {e}")
            return False

# --- 2. GUI フロントエンド (Tkinter) ---

class QuizApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("abc23 model - 操作パネル (Python Edition)")
        self.geometry("1000x600")
        
        self.system = QuizSystem()
        self.player_frames = [] # プレイヤー行のウィジェット管理用

        # 自動読み込みトライ
        default_csv = 'abc23model-1.xlsx - 名簿.csv'
        if not self.system.load_roster(default_csv):
            # ダミーデータ生成
            for i in range(1, 13):
                self.system.roster[i] = Player(i, f"参加者{i}", f"所属{i}")

        self._create_widgets()

    def _create_widgets(self):
        # --- 上部：設定エリア ---
        control_frame = tk.Frame(self, bd=2, relief=tk.GROOVE, padx=10, pady=10)
        control_frame.pack(side=tk.TOP, fill=tk.X)

        # ルール選択
        tk.Label(control_frame, text="ルール:").pack(side=tk.LEFT)
        self.rule_var = tk.StringVar(value="2R: 5○2×")
        rule_cb = ttk.Combobox(control_frame, textvariable=self.rule_var, values=list(self.system.rules_available.keys()), state="readonly", width=15)
        rule_cb.pack(side=tk.LEFT, padx=5)

        # 参加者選択
        tk.Label(control_frame, text="参加ID (カンマ区切り):").pack(side=tk.LEFT, padx=10)
        self.entry_ids = tk.Entry(control_frame, width=30)
        self.entry_ids.insert(0, "1, 2, 3, 4") # 初期値
        self.entry_ids.pack(side=tk.LEFT, padx=5)

        btn_start = tk.Button(control_frame, text="ラウンド開始 / リセット", command=self.start_round, bg="#dddddd")
        btn_start.pack(side=tk.LEFT, padx=10)

        # 問題操作
        tk.Label(control_frame, text=" | ").pack(side=tk.LEFT)
        self.lbl_q_num = tk.Label(control_frame, text="Q.1", font=("Arial", 14, "bold"))
        self.lbl_q_num.pack(side=tk.LEFT, padx=5)
        tk.Button(control_frame, text="スルー / 次へ", command=self.next_question).pack(side=tk.LEFT, padx=5)

        # --- メイン：プレイヤースコアパネル ---
        # スクロール可能なエリアを作成
        self.canvas = tk.Canvas(self)
        self.scrollbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scroll_frame = tk.Frame(self.canvas)

        self.scroll_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 初期表示
        self.start_round()

    def start_round(self):
        # ルール適用
        selected_rule_name = self.rule_var.get()
        self.system.rule = self.system.rules_available[selected_rule_name]
        
        # プレイヤーセットアップ
        raw_ids = self.entry_ids.get().replace("、", ",").split(",")
        self.system.current_players = []
        for rid in raw_ids:
            try:
                pid = int(rid.strip())
                if pid in self.system.roster:
                    p = self.system.roster[pid]
                    p.reset()
                    self.system.current_players.append(p)
            except:
                pass
        
        # 画面再描画
        self.refresh_panel()
        self.lbl_q_num.config(text="Q.1")

    def next_question(self):
        # 現在の問題番号を取得して+1
        current_text = self.lbl_q_num.cget("text").replace("Q.", "")
        try:
            next_num = int(current_text) + 1
            self.lbl_q_num.config(text=f"Q.{next_num}")
        except:
            pass

    def refresh_panel(self):
        # 既存ウィジェット削除
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        
        # ヘッダー
        headers = ["ID", "名前", "所属", "スコア", "誤答数", "状態", "判定操作"]
        for i, h in enumerate(headers):
            tk.Label(self.scroll_frame, text=h, font=("MS Gothic", 10, "bold"), relief=tk.RAISED, bg="#eeeeee").grid(row=0, column=i, sticky="nsew", padx=1, pady=1)

        # プレイヤー行生成
        for idx, player in enumerate(self.system.current_players):
            row = idx + 1
            bg_color = "white"
            fg_color = "black"
            
            # 状態による色分け
            if player.status == PlayerStatus.WIN:
                bg_color = "#ffcccc" # 赤背景（勝ち抜け）
                fg_color = "#cc0000"
            elif player.status == PlayerStatus.LOSE:
                bg_color = "#cccccc" # グレー（失格）
                fg_color = "#666666"

            common_opt = {'bg': bg_color, 'fg': fg_color, 'font': ("Arial", 12), 'padx': 10, 'pady': 5}

            tk.Label(self.scroll_frame, text=str(player.id), **common_opt).grid(row=row, column=0, sticky="nsew", padx=1, pady=1)
            tk.Label(self.scroll_frame, text=player.name, **common_opt).grid(row=row, column=1, sticky="w", padx=1, pady=1)
            tk.Label(self.scroll_frame, text=player.organization, **common_opt).grid(row=row, column=2, sticky="w", padx=1, pady=1)
            
            # スコア表示 (大きく)
            tk.Label(self.scroll_frame, text=str(player.score), bg=bg_color, fg="blue", font=("Arial", 16, "bold")).grid(row=row, column=3, sticky="nsew", padx=1, pady=1)
            
            # 誤答数表示
            wrong_text = "×" * player.wrong_count
            tk.Label(self.scroll_frame, text=wrong_text, bg=bg_color, fg="red", font=("Arial", 14, "bold")).grid(row=row, column=4, sticky="nsew", padx=1, pady=1)

            # ステータス
            tk.Label(self.scroll_frame, text=player.status.value, **common_opt).grid(row=row, column=5, sticky="nsew", padx=1, pady=1)

            # 操作ボタンフレーム
            btn_frame = tk.Frame(self.scroll_frame, bg=bg_color)
            btn_frame.grid(row=row, column=6, padx=1, pady=1)
            
            # ○ボタン
            btn_o = tk.Button(btn_frame, text="○", font=("Arial", 12, "bold"), fg="white", bg="#ff3333", width=4,
                              command=lambda p=player: self.action_correct(p))
            # ×ボタン
            btn_x = tk.Button(btn_frame, text="×", font=("Arial", 12, "bold"), fg="white", bg="#3333ff", width=4,
                              command=lambda p=player: self.action_wrong(p))
            
            if player.status != PlayerStatus.PLAYING:
                btn_o.config(state="disabled", bg="#ddaaaa")
                btn_x.config(state="disabled", bg="#aaaadd")

            btn_o.pack(side=tk.LEFT, padx=2)
            btn_x.pack(side=tk.LEFT, padx=2)

    def action_correct(self, player):
        self.system.rule.on_correct(player)
        self.refresh_panel()
        # 正解が出たら問題を進める（任意）
        self.next_question()

    def action_wrong(self, player):
        self.system.rule.on_wrong(player)
        self.refresh_panel()

# --- 実行 ---
if __name__ == "__main__":
    app = QuizApp()
    app.mainloop()