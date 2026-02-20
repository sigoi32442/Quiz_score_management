import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from PIL import Image, ImageDraw, ImageFont, ImageTk, ImageOps
import os
import csv
import re
import copy
import sys
import traceback

# ==========================================
# エラーハンドリング設定
# ==========================================
def show_error_and_exit(exc_type, exc_value, tb):
    traceback.print_exception(exc_type, exc_value, tb)
    print("\n" + "="*60)
    print("エラーが発生しました。")
    print("Enterキーを押すと終了します...")
    print("="*60)
    input()
    sys.exit(1)

sys.excepthook = show_error_and_exit

# ==========================================
# 設定エリア
# ==========================================
MAIN_FONT_FILE = "SourceHanSansJP-Bold.otf" 
RANK_FONT_FILE = "SourceHanSansJP-Bold.otf" 
HEADER_LOGO_FONT_FILE = "Outfit-ExtraLight.ttf"

IMG_WIDTH = 1920
IMG_HEIGHT = 1000 
BG_COLOR = (0, 0, 0)

FRAME_COLOR = "#BFBFBF"
INNER_BG_TOP = "#B0B0B0"
INNER_BG_BOTTOM = "#7CAEA0" 
WIN_BG_COLOR = "#BD7C38" # オレンジ色
INNER_BORDER = "black"
NAME_STROKE_COLOR = "#2F4F4F"
SCORE_COLOR_NORMAL = (255, 255, 100)
SCORE_COLOR_RENTO = (255, 50, 50) 

WIN_POINTS = 5
LOSE_WRONGS = 2
NUM_GROUPS = 4

COURSES = ["未選択", "Swedish10", "Freeze10", "10by10", "10up-down"]
COURSE_DISPLAY_RECT_COLOR = "#333333"

TIMER_DEFAULT_MIN = 5 
TIMER_DEFAULT_SEC = 0

DISPLAY_WINDOW_SEPARATE_DEFAULT = True
OBS_OVERLAY_DEFAULT = False
OBS_CHROMA_KEY_COLOR = (0, 255, 0)
TIMER_VISIBLE_DEFAULT = True
DISPLAY_WINDOW_TITLE = "RUQabc Display"
DISPLAY_WINDOW_GEOMETRY = "1920x1080"
CONTROL_WINDOW_GEOMETRY = "1500x650"
CONTROL_FONT_FAMILY = "Arial"
CONTROL_FONT_SIZE = 10
CONTROL_FONT_SMALL_SIZE = 9
CONTROL_FONT_BOLD_SIZE = 11
VIEW_FRAME_HEIGHT = 420
CONTROL_STAGE_PADDING = 20
CONTROL_AUTO_SCALE = True
CONTROL_SCALE_FILL = 0.98
CONTROL_SCALE_MIN = 0.85
CONTROL_SCALE_MAX = 1.8

SEMI_RULES = {
    1: {"correct": 1, "wrong": -1},
    2: {"correct": 1, "wrong": -2},
    3: {"correct": 2, "wrong": -2},
}

def get_advantage_points(rank_num):
    if 1 <= rank_num <= 4: return 3
    elif 5 <= rank_num <= 12: return 2
    elif 13 <= rank_num <= 24: return 1
    return 0

def get_empty_player(rank_num):
    suffixes = {1: 'st', 2: 'nd', 3: 'rd'}
    suffix = suffixes.get(rank_num % 10 if not 11 <= rank_num <= 13 else 0, 'th')
    
    r_color = "#008000" 
    if 1 <= rank_num <= 4: r_color = "#c00000"    
    elif 5 <= rank_num <= 12: r_color = "#0000c0" 
    elif 13 <= rank_num <= 24: r_color = "#d08000" 
    elif rank_num >= 49: r_color = "#800080"        

    return {
        "rank": f"{rank_num}{suffix}",
        "rank_num": rank_num,
        "rank_color": r_color,
        "name": "---", "univ": "---", 
        "photo_path": None, 
        "score": get_advantage_points(rank_num), "wrong": 0, "rento": False, "win_order": 0,
        
        "10by10_o": 0, "10by10_x": 10, "win_order_10by10": 0,
        "Swedish10_o": 0, "Swedish10_x": 0, "win_order_Swedish10": 0,
        "Freeze10_o": 0, "Freeze10_x": 0, "Freeze10_freeze": 0, "win_order_Freeze10": 0,
        "10up-down_score": 0, "10up-down_wrong": 0, "win_order_10up-down": 0,
        "semi_score": 0, "win_order_semi": 0,
        "semi_status": "active", "semi_exit_set": 0, # active, win, lose
        "final_sets_won": 0, "final_curr_o": 0, "final_curr_x": 0, "final_set_lost": False, "win_order_final": 0,
        "extra_score": 0, "extra_wrong": 0, "win_order_extra": 0
    }

def get_ordinal_str(n):
    if 11 <= n <= 13: return f"{n}th"
    s = n % 10
    if s == 1: return f"{n}st"
    if s == 2: return f"{n}nd"
    if s == 3: return f"{n}rd"
    return f"{n}th"

# ==========================================
# 描画エンジン
# ==========================================
class ScoreboardDrawer:
    def __init__(self):
        self.main_path = self.find_font_path(MAIN_FONT_FILE)
        self.rank_path = self.find_font_path(RANK_FONT_FILE)
        self.header_path = self.find_font_path(HEADER_LOGO_FONT_FILE)
        self.qa_path = self.find_font_path("BIZ-UDPGothic-Regular.ttc") or \
                       self.find_font_path("BIZ-UDPGothic.ttf") or \
                       self.find_font_path("BIZUDPGothic-Regular.ttf")
        self.load_fonts()
        self.photo_cache = {}

    def find_font_path(self, f):
        candidates = [f, os.path.join(os.getcwd(), f), f"C:/Windows/Fonts/{f}", f"/Library/Fonts/{f}",
                      f"/System/Library/Fonts/Supplemental/{f}",
                      f"C:/Users/{os.getlogin()}/AppData/Local/Microsoft/Windows/Fonts/{f}" if os.name == 'nt' else f]
        for p in candidates:
            if os.path.exists(p) and p is not None: return p
        return None

    def load_fonts(self):
        try:
            self.font_logo = ImageFont.truetype(self.header_path or "arial.ttf", 54)
            self.font_header_sub = ImageFont.truetype(self.header_path or "arial.ttf", 54) 
            self.font_msg = ImageFont.truetype(self.qa_path or self.main_path or "arial.ttf", 45)
            self.font_course_display = ImageFont.truetype(self.qa_path or self.main_path or "arial.ttf", 80)
            self.font_mark = ImageFont.truetype(self.main_path or "arial.ttf", 60)
            self.font_main_score = ImageFont.truetype(self.rank_path or "arial.ttf", 80)
            self.font_sub_score = ImageFont.truetype(self.rank_path or "arial.ttf", 50)
            self.font_timer = ImageFont.truetype(self.header_path or "arial.ttf", 150)
            self.font_semi_timer = ImageFont.truetype(self.header_path or "arial.ttf", 180) 
            
            # SFのスコアフォントをOutfit (header_path) に変更
            self.font_semi_score = ImageFont.truetype(self.header_path or "arial.ttf", 100)
            
            self.font_semi_rank = ImageFont.truetype(self.rank_path or "arial.ttf", 40)
            self.font_semi_univ = ImageFont.truetype(self.main_path or "arial.ttf", 24)
            self.font_semi_name = ImageFont.truetype(self.main_path or "arial.ttf", 48)
        except:
            self.font_logo = self.font_header_sub = self.font_msg = ImageFont.load_default()
            self.font_course_display = self.font_mark = self.font_main_score = ImageFont.load_default()
            self.font_sub_score = self.font_timer = ImageFont.load_default()
            self.font_semi_timer = self.font_semi_score = self.font_semi_rank = self.font_semi_univ = self.font_semi_name = ImageFont.load_default()

    def get_resized_photo(self, path, w, h):
        key = (path, w, h)
        if key in self.photo_cache: return self.photo_cache[key]
        try:
            img = Image.open(path).convert("RGBA")
            img = ImageOps.fit(img, (int(w), int(h)), Image.Resampling.LANCZOS)
            self.photo_cache[key] = img
            return img
        except:
            return None

    def draw_text_fit(self, draw, text, cx, bottom_y, max_w, font_path, max_size, color, stroke_w, stroke_c, align="center", x_offset=0):
        size = max_size
        font = None
        while size > 10:
            try:
                font = ImageFont.truetype(font_path or "arial.ttf", size)
            except:
                font = ImageFont.load_default(); break
            bbox = draw.textbbox((0, 0), text, font=font)
            width = bbox[2] - bbox[0]
            if width <= max_w: break
            size -= 2
        
        bbox = draw.textbbox((0, 0), text, font=font)
        w = bbox[2] - bbox[0]
        h = bbox[3] - bbox[1]
        
        draw_x = cx
        if align == "center":
            draw_x = cx - w // 2
        elif align == "right":
            draw_x = cx - w
        elif align == "left":
            draw_x = cx 
        
        draw_x += x_offset
        draw.text((draw_x, bottom_y - h - 5), text, font=font, fill=color, stroke_width=stroke_w, stroke_fill=stroke_c)

    def draw_wrapped_text(self, draw, text, x, y, font, fill, max_width):
        lines = []
        words = list(text); current_line = ""
        for char in words:
            test_line = current_line + char
            if draw.textbbox((0, 0), test_line, font=font)[2] <= max_width: current_line = test_line
            else: lines.append(current_line); current_line = char
        lines.append(current_line)
        cur_y = y
        for line in lines:
            draw.text((x, cur_y), line, font=font, fill=fill); cur_y += 10 + (font.size if hasattr(font, 'size') else 20)
        return cur_y

    def draw_player_plate(self, im, draw, p, xb, sy, cw, ch, scale, is_3rd=False, mode="2R"):
        lost, win = False, False
        win_order_num = 0
        
        if mode == "2R":
            lost = p.get("wrong", 0) >= LOSE_WRONGS
            win_order_num = p.get("win_order", 0)
        elif mode == "10by10":
            lost = p.get("10by10_x", 10) <= 4
            win_order_num = p.get("win_order_10by10", 0)
        elif mode == "Swedish10":
            lost = p.get("Swedish10_x", 0) >= 10
            win_order_num = p.get("win_order_Swedish10", 0)
        elif mode == "Freeze10":
            win_order_num = p.get("win_order_Freeze10", 0)
        elif mode == "10up-down":
            lost = p.get("10up-down_wrong", 0) >= 2
            win_order_num = p.get("win_order_10up-down", 0)
        elif mode == "SEMI":
            win_order_num = p.get("win_order_semi", 0)
            if p.get("semi_status") == "win": win = True
            if p.get("semi_status") == "lose": lost = True
        elif mode == "FINAL":
            lost = p.get("final_set_lost", False) 
            win_order_num = p.get("win_order_final", 0)
        elif mode == "EXTRA":
            lost = p.get("extra_wrong", 0) >= 1
            win_order_num = p.get("win_order_extra", 0)
            
        win = (win_order_num > 0) or (mode == "SEMI" and p.get("semi_status") == "win")

        fc, rc, it = ("#555", "#555", "#666") if lost else (FRAME_COLOR, p["rank_color"], INNER_BG_TOP)
        is_frozen = (mode == "Freeze10" and p.get("Freeze10_freeze", 0) > 0)
        if is_frozen and not win: it = "#444"
        if mode == "FINAL" and p.get("final_set_lost", False) and not win: it = "#444"

        # 外枠
        draw.rectangle((xb, sy, xb+cw, sy+ch), fill=fc)
        h_h = 8 * scale
        # 上下のカラーバー
        draw.rectangle((xb+(cw-40*scale)/2, sy, xb+(cw+40*scale)/2, sy+h_h), fill=rc)
        draw.rectangle((xb+(cw-40)/2, sy+ch-h_h, xb+(cw+40*scale)/2, sy+ch), fill=rc)
        
        pad = 8 * scale
        ix, iy, iw, ih = xb+pad, sy+pad, cw-pad*2, ch-pad*2
        
        draw.rectangle((ix, iy, ix+iw, iy+ih), fill=it)

        if mode == "SEMI":
            text_area_h = 90
            photo_h = ih - text_area_h
            if p.get("photo_path"):
                photo = self.get_resized_photo(p["photo_path"], iw, photo_h)
                if photo:
                    if lost: photo = ImageOps.grayscale(photo).convert("RGBA")
                    im.paste(photo, (int(ix), int(iy)))
            draw.rectangle((ix, iy + photo_h, ix + iw, iy + ih), fill=it)

        elif mode == "FINAL":
            text_area_h = 60
            photo_margin = int(15 * scale)
            draw.rectangle((ix, iy + ih - text_area_h, ix + iw, iy + ih), fill=it)
            
            if not win and not lost:
                ratio = min(p.get("final_curr_o", 0), 7) / 7.0
                if ratio > 0:
                    fill_color = INNER_BG_BOTTOM
                    if p.get("final_curr_o", 0) >= 7:
                        fill_color = WIN_BG_COLOR
                    draw.rectangle((ix, iy+ih*(1-ratio), ix+iw, iy+ih), fill=fill_color)

            if p.get("photo_path"):
                photo_w = iw - (photo_margin * 2)
                photo_h = (ih - text_area_h) - (photo_margin * 2)
                photo = self.get_resized_photo(p["photo_path"], photo_w, photo_h)
                if photo:
                    if lost: photo = ImageOps.grayscale(photo).convert("RGBA")
                    im.paste(photo, (int(ix + photo_margin), int(iy + photo_margin)))
            
            sets_won = p.get("final_sets_won", 0)
            if sets_won > 0:
                try:
                    f_star = ImageFont.truetype(self.main_path or "arial.ttf", int(50 * scale))
                except:
                    f_star = ImageFont.load_default()
                base_star_x = ix + photo_margin - 5
                base_star_y = (iy + ih - text_area_h - photo_margin) - 10 * scale
                for s_i in range(sets_won):
                    star_y = base_star_y - (s_i * (40 * scale)) - (f_star.size)
                    draw.text((base_star_x, star_y), "★", font=f_star, fill="#FFD700", stroke_width=int(2*scale), stroke_fill="black")
            
        else:
            if p.get("photo_path"):
                photo = self.get_resized_photo(p["photo_path"], iw, ih)
                if photo:
                    if lost and mode != "FINAL": photo = ImageOps.grayscale(photo).convert("RGBA")
                    im.paste(photo, (int(ix), int(iy)))

            if not is_3rd and mode not in ["SEMI"]:
                if win: draw.rectangle((ix, iy, ix+iw, iy+ih), fill=WIN_BG_COLOR)
                elif lost: draw.rectangle((ix, iy, ix+iw, iy+ih), fill="#444")
                else:
                    ratio = 0
                    if mode == "2R": ratio = min(p["score"], WIN_POINTS) / WIN_POINTS
                    elif mode == "10by10": ratio = min(p["10by10_o"] * p["10by10_x"], 100) / 100
                    elif mode == "Swedish10": ratio = min(p["Swedish10_o"], 10) / 10
                    elif mode == "Freeze10": ratio = min(p["Freeze10_o"], 10) / 10
                    elif mode == "10up-down": ratio = min(p["10up-down_score"], 10) / 10
                    elif mode == "EXTRA": ratio = min(p["extra_score"], 5) / 5
                    
                    if ratio > 0 and not is_frozen:
                         draw.rectangle((ix, iy+ih*(1-ratio), ix+iw, iy+ih), fill=INNER_BG_BOTTOM)

        draw.rectangle((ix, iy, ix+iw, iy+ih), outline=INNER_BORDER, width=max(1, int(5*scale)))
        
        # --- テキスト描画 ---
        if mode == "SEMI":
            cx = xb + cw / 2
            rank_txt = p["rank"]
            rb = draw.textbbox((0,0), rank_txt, font=self.font_semi_rank)
            draw.text((xb + (cw - (rb[2]-rb[0]))//2, sy - 40), rank_txt, font=self.font_semi_rank, fill="white", stroke_width=3, stroke_fill=rc)
            
            pure_name = p["name"].replace(" ", "").replace("　", "")
            name_len = len(pure_name)
            display_name = pure_name
            if name_len == 3:
                 parts = re.split(r'[ 　]+', p["name"])
                 if len(parts) >= 2: display_name = f"{parts[0]}　{parts[1]}"

            target_size = 52
            if name_len <= 4: target_size = 40
            elif name_len == 5: target_size = 34
            
            self.draw_text_fit(draw, p["univ"], cx, sy + ch - 70, cw - 10, self.main_path, 24, "white", 2,  NAME_STROKE_COLOR, align="center")
            self.draw_text_fit(draw, display_name, cx, sy + ch - 20, cw - 5, self.main_path, target_size, "white", 4, NAME_STROKE_COLOR, align="center")

        elif mode == "FINAL":
            cx = xb + cw / 2
            rank_txt = p["rank"]
            rb = draw.textbbox((0,0), rank_txt, font=self.font_semi_rank)
            draw.text((xb + (cw - (rb[2]-rb[0]))//2, sy - 30), rank_txt, font=self.font_semi_rank, fill="white", stroke_width=3, stroke_fill=rc)

            pure_name = p["name"].replace(" ", "").replace("　", "")
            name_len = len(pure_name)
            display_name = pure_name
            if name_len == 3:
                 parts = re.split(r'[ 　]+', p["name"])
                 if len(parts) >= 2: display_name = f"{parts[0]}　{parts[1]}"
            
            self.draw_text_fit(draw, p["univ"], ix + 15, sy + ch - 27, (cw/2)-20, self.main_path, 28, "white", 3, NAME_STROKE_COLOR, align="left")
            self.draw_text_fit(draw, display_name, ix + iw - 25, sy + ch - 26, (cw/2), self.main_path, 40, "white", 4, NAME_STROKE_COLOR, align="right")

        else:
            f_rank = ImageFont.truetype(self.rank_path or "arial.ttf", int(40 * scale))
            r_w = draw.textbbox((0,0), p["rank"], font=f_rank)[2]
            draw.text((xb+cw/2-r_w/2, sy-25*scale), p["rank"], font=f_rank, fill="white", stroke_width=int(5*scale), stroke_fill=rc)
            
            pure = p["name"].replace(" ","").replace("　","")
            base_sz = 75 if len(pure) < 5 else 58
            f_nm = ImageFont.truetype(self.main_path or "arial.ttf", int(base_sz * scale))
            grid_h = f_nm.size + 10 * scale
            for idx, c in enumerate(pure):
                dy = iy + 5*scale + (idx * grid_h)
                m = re.search(r'[ 　]', p["name"]); sp_idx = m.start() if m else -1
                if len(pure)==3 and sp_idx!=-1 and idx>=sp_idx: dy += grid_h
                draw.text((xb+cw*0.45 - draw.textbbox((0,0),c,font=f_nm)[2]/2, dy), c, font=f_nm, fill="gray" if lost and mode!="FINAL" else "white", stroke_width=int(4*scale), stroke_fill=NAME_STROKE_COLOR)
            
            f_univ = ImageFont.truetype(self.main_path or "arial.ttf", int(24 * scale))
            uy = iy + 10 * scale
            for c in p["univ"]:
                draw.text((xb+cw*0.82-draw.textbbox((0,0),c,font=f_univ)[2]/2, uy), c, font=f_univ, fill="gray" if lost and mode!="FINAL" else "white", stroke_width=int(2*scale), stroke_fill="black")
                uy += (f_univ.size + 2)

        if is_3rd and mode == "2R": return

        def draw_center_scaled(text, font_path, base_size, color, offset_y=10):
            try: f = ImageFont.truetype(font_path or "arial.ttf", int(base_size * scale))
            except: f = ImageFont.load_default()
            bbox = draw.textbbox((0, 0), text, font=f)
            draw.text((xb+cw/2-(bbox[2]-bbox[0])/2, sy+ch+offset_y), text, font=f, fill=color)

        if win:
            rank_val = 0
            if mode == "2R": rank_val = p.get("win_order", 0)
            elif mode == "10by10": rank_val = p.get("win_order_10by10", 0)
            elif mode == "Swedish10": rank_val = p.get("win_order_Swedish10", 0)
            elif mode == "Freeze10": rank_val = p.get("win_order_Freeze10", 0)
            elif mode == "10up-down": rank_val = p.get("win_order_10up-down", 0)
            elif mode == "EXTRA": rank_val = p.get("win_order_extra", 0)
            
            if rank_val > 0:
                txt = get_ordinal_str(rank_val)
                # 変更: rank_path -> header_path
                draw_center_scaled(txt, self.header_path, 80, "red")
            else:
                txt = "WIN"
                # 変更: rank_path -> header_path
                draw_center_scaled(txt, self.header_path, 80, "red")
        elif lost and mode != "FINAL":
            txt = "LOSE"
            # 変更: rank_path -> header_path
            draw_center_scaled(txt, self.header_path, 80, "gray")
        else:
            if mode == "2R":
                draw_center_scaled(str(p["score"]), self.header_path, 80, SCORE_COLOR_RENTO if p["rento"] else SCORE_COLOR_NORMAL)
                if p["wrong"] > 0:
                    f_mark_scaled = ImageFont.truetype(self.main_path or "arial.ttf", int(60 * scale))
                    for xi in range(p["wrong"]):
                        draw.text((xb+cw/2-25*scale+xi*50*scale-draw.textbbox((0,0),"×",font=f_mark_scaled)[2]/2, sy+ch+100*scale), "×", font=f_mark_scaled, fill="white")
            elif mode == "10by10":
                val = p["10by10_o"] * p["10by10_x"]
                draw_center_scaled(str(val), self.header_path, 80, SCORE_COLOR_NORMAL)
                draw_center_scaled(f"{p['10by10_o']}  {p['10by10_x']}", self.header_path, 50, "white", 100)
            elif mode == "Swedish10":
                draw_center_scaled(str(p["Swedish10_o"]), self.header_path, 80, SCORE_COLOR_NORMAL)
                # 修正: Swedish10の誤答数を白に変更
                draw_center_scaled(f"{p['Swedish10_x']}x", self.header_path, 50, "white", 100)
            elif mode == "Freeze10":
                if is_frozen: 
                    # 修正: Freeze表示の変更
                    draw_center_scaled(f"Freeze {p['Freeze10_freeze']}", self.header_path, 50, "cyan", 30)
                else:
                    draw_center_scaled(str(p["Freeze10_o"]), self.header_path, 80, SCORE_COLOR_NORMAL)
                    draw_center_scaled(f"{p['Freeze10_x']}x", self.header_path, 50, "white", 100)
            elif mode == "10up-down":
                draw_center_scaled(str(p["10up-down_score"]), self.header_path, 80, SCORE_COLOR_NORMAL)
                if p["10up-down_wrong"] > 0:
                    f_mark_scaled = ImageFont.truetype(self.main_path or "arial.ttf", int(60 * scale))
                    draw.text((xb+cw/2-draw.textbbox((0,0),"×",font=f_mark_scaled)[2]/2, sy+ch+100*scale), "×", font=f_mark_scaled, fill="red")
            elif mode == "FINAL":
                if p["final_set_lost"]:
                    # 変更: rank_path -> header_path
                    draw_center_scaled("LOSE", self.header_path, 80, "gray")
                else:
                    draw_center_scaled(str(p["final_curr_o"]), self.header_path, 80, SCORE_COLOR_NORMAL)
                    if p["final_curr_x"] > 0:
                         f_mark_scaled = ImageFont.truetype(self.main_path or "arial.ttf", int(60 * scale))
                         for xi in range(p["final_curr_x"]):
                            draw.text((xb+cw/2-25*scale+xi*50*scale-draw.textbbox((0,0),"×",font=f_mark_scaled)[2]/2, sy+ch+100*scale), "×", font=f_mark_scaled, fill="white")

            elif mode == "EXTRA":
                draw_center_scaled(str(p["extra_score"]), self.header_path, 80, SCORE_COLOR_NORMAL)

    def _build_header_text(self, mode, group_idx, semi_set_idx):
        if mode == "2R":
            return f"2nd Round Group{group_idx + 1}"
        if mode.startswith("3rd"):
            return "3rd Round course select"
        if mode == "SEMI":
            return f"Semifinal Nine Hundred - Set {semi_set_idx}"
        if mode == "FINAL":
            return "Final - Triple Seven"
        if mode == "EXTRA":
            return "Extra Round 2nd Step"
        if mode == "SF_FOLLOW":
            return "SF Follow-up"
        return f"3rd Round {mode}"

    def _get_display_players(self, players, mode):
        if mode == "2R":
            return players[:12]
        if mode == "SEMI":
            return players[:9]
        if mode == "FINAL":
            return players[:3]
        if mode == "EXTRA":
            return players[:12]
        return players[:5]

    def _get_obs_score_text(self, p, mode, sf_hide_scores=False):
        if mode == "2R":
            if p.get("win_order", 0) > 0:
                return f"{get_ordinal_str(p['win_order'])} WIN"
            if p.get("wrong", 0) >= LOSE_WRONGS:
                return "LOSE"
            wrong_marks = "x" * p.get("wrong", 0)
            return f"{p.get('score', 0)} {wrong_marks}".strip()
        if mode == "10by10":
            if p.get("win_order_10by10", 0) > 0:
                return f"{get_ordinal_str(p['win_order_10by10'])} WIN"
            return f"{p.get('10by10_o', 0)}o {p.get('10by10_x', 0)}x"
        if mode == "Swedish10":
            if p.get("win_order_Swedish10", 0) > 0:
                return f"{get_ordinal_str(p['win_order_Swedish10'])} WIN"
            return f"{p.get('Swedish10_o', 0)}o {p.get('Swedish10_x', 0)}x"
        if mode == "Freeze10":
            if p.get("win_order_Freeze10", 0) > 0:
                return f"{get_ordinal_str(p['win_order_Freeze10'])} WIN"
            if p.get("Freeze10_freeze", 0) > 0:
                return f"Freeze {p.get('Freeze10_freeze', 0)}"
            return f"{p.get('Freeze10_o', 0)}o {p.get('Freeze10_x', 0)}x"
        if mode == "10up-down":
            if p.get("win_order_10up-down", 0) > 0:
                return f"{get_ordinal_str(p['win_order_10up-down'])} WIN"
            return f"{p.get('10up-down_score', 0)}pt {p.get('10up-down_wrong', 0)}x"
        if mode == "SEMI":
            if p.get("semi_status") == "win":
                return "WIN"
            if p.get("semi_status") == "lose":
                return "LOSE"
            if sf_hide_scores:
                return "?"
            return f"{p.get('semi_score', 0)}pt"
        if mode == "FINAL":
            if p.get("win_order_final", 0) > 0:
                return "CHAMPION"
            if p.get("final_set_lost", False):
                return "LOSE"
            return f"{p.get('final_sets_won', 0)}S {p.get('final_curr_o', 0)}o {p.get('final_curr_x', 0)}x"
        if mode == "EXTRA":
            if p.get("win_order_extra", 0) > 0:
                return "WIN"
            if p.get("extra_wrong", 0) > 0:
                return "LOSE"
            return f"{p.get('extra_score', 0)}pt"
        return "-"

    def generate_image_obs_overlay(self, players, group_idx, question_text, answer_text, timer_str="00:00", timer_alert=False, mode="2R", semi_set_idx=1, sf_hide_scores=False, question_index=None, show_timer=True):
        im = Image.new('RGB', (IMG_WIDTH, IMG_HEIGHT), OBS_CHROMA_KEY_COLOR)
        draw = ImageDraw.Draw(im)

        header_text = self._build_header_text(mode, group_idx, semi_set_idx)
        q_number_text = f"Q{question_index}" if question_index is not None else "Q"

        top_h1 = 52
        top_h2 = 110
        top_h3 = 58
        bottom_top = 760

        draw.rectangle((0, 0, IMG_WIDTH, top_h1), fill="#060606")
        draw.rectangle((0, top_h1, IMG_WIDTH, top_h1 + top_h2), fill="#111111")
        draw.rectangle((0, top_h1 + top_h2, IMG_WIDTH, top_h1 + top_h2 + top_h3), fill="#0a0a0a")
        draw.line([(0, top_h1 + top_h2), (IMG_WIDTH, top_h1 + top_h2)], fill="#2a2a2a", width=2)
        draw.rectangle((0, bottom_top, IMG_WIDTH, IMG_HEIGHT), fill="#2f3338")

        right_margin = 240 if show_timer else 10
        self.draw_text_fit(draw, header_text, 10, 44, IMG_WIDTH - 220, self.qa_path or self.main_path, 46, "white", 0, "black", align="left")
        self.draw_text_fit(draw, q_number_text, IMG_WIDTH - right_margin, 44, 220, self.qa_path or self.main_path, 46, "#d7d7d7", 0, "black", align="right")

        if show_timer:
            timer_color = "#ff3333" if timer_alert else "#3cc8ff"
            self.draw_text_fit(draw, timer_str, IMG_WIDTH - 20, 44, 220, self.header_path or self.main_path, 44, timer_color, 0, "black", align="right")
        self.draw_text_fit(draw, f"Q. {question_text}", 14, top_h1 + top_h2 - 18, IMG_WIDTH - 28, self.qa_path or self.main_path, 52, "#f5f5f5", 0, "black", align="left")
        self.draw_text_fit(draw, f"A. {answer_text}", IMG_WIDTH - 14, top_h1 + top_h2 + top_h3 - 10, IMG_WIDTH - 28, self.qa_path or self.main_path, 52, "#ffe45c", 0, "black", align="right")

        display_players = self._get_display_players(players, mode)
        if not display_players:
            return im

        pad_x = 8
        gap = 3
        n = len(display_players)
        col_w = (IMG_WIDTH - (pad_x * 2) - (gap * (n - 1))) / n

        for i, p in enumerate(display_players):
            x1 = pad_x + i * (col_w + gap)
            x2 = x1 + col_w
            rank_color = p.get("rank_color", "#3a8f3a")
            draw.rectangle((x1, bottom_top + 42, x2, IMG_HEIGHT - 4), fill="#4a4a4a", outline="#9a9a9a", width=1)
            draw.rectangle((x1, bottom_top, x2, bottom_top + 42), fill=rank_color)

            self.draw_text_fit(draw, p.get("rank", "---"), int((x1 + x2) / 2), bottom_top + 38, int(col_w - 8), self.rank_path or self.main_path, 34, "white", 1, "black", align="center")
            self.draw_text_fit(draw, p.get("name", "---"), int((x1 + x2) / 2), bottom_top + 128, int(col_w - 8), self.main_path, 44, "white", 2, "black", align="center")
            self.draw_text_fit(draw, p.get("univ", "---"), int((x1 + x2) / 2), bottom_top + 172, int(col_w - 8), self.main_path, 28, "#efefef", 1, "black", align="center")
            score_txt = self._get_obs_score_text(p, mode, sf_hide_scores=sf_hide_scores)
            self.draw_text_fit(draw, score_txt, int((x1 + x2) / 2), IMG_HEIGHT - 8, int(col_w - 8), self.header_path or self.rank_path, 42, "#fff05f", 2, "black", align="center")

        return im

    def generate_image(self, players, group_idx, question_text, answer_text, timer_str="00:00", timer_alert=False, mode="2R", semi_set_idx=1, sf_hide_scores=False, obs_overlay=False, question_index=None, show_timer=True):
        if obs_overlay and mode != "SF_FOLLOW":
            return self.generate_image_obs_overlay(
                players,
                group_idx,
                question_text,
                answer_text,
                timer_str=timer_str,
                timer_alert=timer_alert,
                mode=mode,
                semi_set_idx=semi_set_idx,
                sf_hide_scores=sf_hide_scores,
                question_index=question_index,
                show_timer=show_timer,
            )

        im = Image.new('RGB', (IMG_WIDTH, IMG_HEIGHT), BG_COLOR)
        draw = ImageDraw.Draw(im)
        
        logo_text = "RUQabc"
        draw.text((50, 40), logo_text, font=self.font_logo, fill="white")
        
        header = self._build_header_text(mode, group_idx, semi_set_idx)
        draw.text((IMG_WIDTH - draw.textbbox((0, 0), header, font=self.font_header_sub)[2] - 50, 40), header, font=self.font_header_sub, fill="white")
        
        # 修正: コロン間隔調整
        def draw_fixed_pitch_timer(cx, cy, text, font, color, pitch):
            offsets = [-1.75, -0.8, 0, 0.8, 1.75]
            for i, char in enumerate(text):
                if i < len(offsets):
                    char_x = cx + offsets[i] * pitch
                    draw.text((char_x, cy), char, font=font, fill=color, anchor="mm")

        if show_timer and mode in ["10by10", "Swedish10", "Freeze10", "10up-down"]:
            center_x, center_y = 1640, 550
            draw_fixed_pitch_timer(center_x, center_y, timer_str, self.font_timer, "red" if timer_alert else "#00FFFF", 100)
        elif show_timer and mode == "SEMI":
            center_x, center_y = IMG_WIDTH // 2 , 250
            draw_fixed_pitch_timer(center_x, center_y, timer_str, self.font_semi_timer, "red" if timer_alert else "#00FFFF", 120)

        if mode != "SEMI" and mode != "SF_FOLLOW":
            last_y = self.draw_wrapped_text(draw, f"Q. {question_text}", 50, 140, self.font_msg, "white", IMG_WIDTH - 100)
            self.draw_wrapped_text(draw, f"A. {answer_text}", 50, last_y + 10, self.font_msg, "yellow", IMG_WIDTH - 100)
        
        # --- Draw Logic for Scoreboards ---
        if mode == "SF_FOLLOW":
            return im

        mx, sy, cw, ch, gap = 0, 350, 0, 0, 0
        current_scale = 1.0

        if mode == "2R" or mode == "EXTRA":
            mx, sy, cw, ch, gap = 50, 350, 132, 380, 22
            current_scale = 1.0
        elif mode == "SEMI":
            cw, ch, gap = 200, 320, 10 
            total_w = (cw * 9) + (gap * 8)
            mx = (IMG_WIDTH - total_w) // 2
            sy = 450 
            current_scale = 1.0 
        elif mode == "FINAL":
            cw, ch, gap = 480, 320, 80
            total_w = (cw * 3) + (gap * 2)
            mx = (IMG_WIDTH - total_w) // 2
            sy = 350
            current_scale = 1.1
        else: # 3rd
            mx, sy, cw, ch, gap = 150, 350, 145, 400, 120
            current_scale = 1.06 # 1.1 -> 1.06
        
        display_players = self._get_display_players(players, mode)

        filtered_players_for_drawing = []
        if mode == "SEMI":
            for p in display_players:
                if p["semi_status"] != "active" and p["semi_exit_set"] < semi_set_idx:
                    filtered_players_for_drawing.append(None) 
                else:
                    filtered_players_for_drawing.append(p)
        elif mode == "EXTRA":
            for p in display_players:
                if p["name"] == "---":
                    filtered_players_for_drawing.append(None)
                else:
                    filtered_players_for_drawing.append(p)
        else:
            filtered_players_for_drawing = display_players

        for i, p in enumerate(filtered_players_for_drawing):
            if p is None: continue 
            px, py = mx+i*(cw+gap), sy
            self.draw_player_plate(im, draw, p, px, py, cw, ch, current_scale, is_3rd=False, mode=mode)
            
            if mode == "SEMI" and p.get("semi_status") == "active":
                if sf_hide_scores:
                    score_text = "?"
                else:
                    score_val = p.get("semi_score", 0)
                    score_text = str(score_val)
                s_bbox = draw.textbbox((0,0), score_text, font=self.font_semi_score)
                sw, sh = s_bbox[2]-s_bbox[0], s_bbox[3]-s_bbox[1]
                draw.text((px + cw//2 - sw//2, py + ch + 30), score_text, font=self.font_semi_score, fill="yellow")

        return im

    def generate_image_3rd_round(self, players, selected_course_name, player_selections):
        im = Image.new('RGB', (IMG_WIDTH, IMG_HEIGHT), BG_COLOR)
        draw = ImageDraw.Draw(im)
        draw.text((50, 40), "RUQabc", font=self.font_logo, fill="white")
        txt = "3rd Round course select"
        draw.text((IMG_WIDTH - draw.textbbox((0, 0), txt, font=self.font_header_sub)[2] - 50, 40), txt, font=self.font_header_sub, fill="white")
        rect_x, rect_y, rect_w, rect_h = IMG_WIDTH // 2 - 350, 180, 700, 120
        draw.rectangle((rect_x, rect_y, rect_x + rect_w, rect_y + rect_h), fill=COURSE_DISPLAY_RECT_COLOR, outline="white", width=2)
        if selected_course_name and selected_course_name != "未選択":
            bbox = draw.textbbox((0, 0), selected_course_name, font=self.font_course_display)
            draw.text((rect_x + (rect_w - (bbox[2]-bbox[0]))//2, rect_y + (rect_h - (bbox[3]-bbox[1]))//2), selected_course_name, font=self.font_course_display, fill="white")
        
        # 0.6 -> 0.65 に戻した
        mx, sy, cw, ch, gap, scale = 15, 480, 88, 380*0.65, 7, 0.65
        
        for i in range(20):
            if i in player_selections and player_selections[i] > 0: continue
            p = players[i]
            if p["name"] == "---": continue
            self.draw_player_plate(im, draw, p, mx+i*(cw+gap), sy, cw, ch, scale, is_3rd=True, mode="2R")
        return im

    def generate_image_sf_follow(self, questions, start_idx, end_idx, offset):
        im = Image.new('RGB', (IMG_WIDTH, IMG_HEIGHT), BG_COLOR)
        draw = ImageDraw.Draw(im)
        draw.text((50, 40), "RUQabc", font=self.font_logo, fill="white")
        
        header = "SF Follow-up"
        draw.text((IMG_WIDTH - draw.textbbox((0, 0), header, font=self.font_header_sub)[2] - 50, 40), header, font=self.font_header_sub, fill="white")
        
        base_y_list = [150, 420, 690] 
        
        for i in range(3):
            q_idx = start_idx + offset + i
            if q_idx > end_idx or q_idx >= len(questions):
                break
            
            y_pos = base_y_list[i]
            q_data = questions[q_idx]
            
            # Question Text
            next_y = self.draw_wrapped_text(draw, q_data["q"], 50, y_pos, self.font_msg, "white", IMG_WIDTH - 100)
            
            # Answer Text
            self.draw_wrapped_text(draw, f"A. {q_data['a']}", 50, next_y + 10, self.font_msg, "yellow", IMG_WIDTH - 100)
            
            if i < 2:
                draw.line([(50, base_y_list[i+1] - 30), (IMG_WIDTH-50, base_y_list[i+1] - 30)], fill="#444", width=2)

        return im

# ==========================================
# メインアプリケーション
# ==========================================
class QuizApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RUQabc HQ Manager"); self.geometry(CONTROL_WINDOW_GEOMETRY)
        self.all_groups_data = [[get_empty_player((i*4)+(g+1)) for i in range(12)] for g in range(NUM_GROUPS)]
        self.questions = []; self.current_q_idx = 0; self.current_group_idx = 0
        self.history_stacks = {} 
        self.history_stacks_3rd = {} 
        self.mode = "SCORE"
        self.players_3rd_20 = [get_empty_player(99) for _ in range(20)]
        self.current_selected_course_name_3rd = "未選択"
        self.player_selections_3rd = {}
        self.timer_seconds = TIMER_DEFAULT_MIN * 60 + TIMER_DEFAULT_SEC
        self.timer_running = False
        
        self.semi_set_idx = 1
        self.players_semi_9 = [get_empty_player(99) for _ in range(9)]
        self.players_final_3 = [get_empty_player(99) for _ in range(3)]
        self.players_extra_12 = [get_empty_player(99) for _ in range(12)]
        
        # SF Follow State
        self.sf_follow_start = 0
        self.sf_follow_end = 0
        self.sf_follow_cursor = 0
        
        # SF Hide Score State
        self.sf_hide_scores = False

        self.display_window = None
        self.display_separate_var = tk.BooleanVar(value=DISPLAY_WINDOW_SEPARATE_DEFAULT)
        self.obs_overlay_var = tk.BooleanVar(value=OBS_OVERLAY_DEFAULT)
        self.timer_visible_var = tk.BooleanVar(value=TIMER_VISIBLE_DEFAULT)
        self.controls_container = None
        self.control_stage = None
        self.default_font = f"{CONTROL_FONT_FAMILY} {CONTROL_FONT_SIZE}"
        self.small_font = (CONTROL_FONT_FAMILY, CONTROL_FONT_SMALL_SIZE)
        self.bold_font = (CONTROL_FONT_FAMILY, CONTROL_FONT_BOLD_SIZE, "bold")
        self.control_base_size = None
        self.control_scale = 1.0
        self._control_scale_job = None
        self.base_tk_scaling = float(self.tk.call("tk", "scaling"))

        self.drawer = ScoreboardDrawer(); self._update_job = None
        self.setup_ui()
        self.refresh_ui()

    def _build_view_frame(self, parent, fill="x", expand=False, before=None):
        if self.view_frame is not None:
            self.view_frame.destroy()
        frame = tk.Frame(parent, bg="#222", height=VIEW_FRAME_HEIGHT)
        if before is not None:
            frame.pack(side="top", fill=fill, expand=expand, before=before)
        else:
            frame.pack(side="top", fill=fill, expand=expand)
        frame.pack_propagate(False)
        label = tk.Label(frame, bg="#222")
        label.place(relx=0.5, rely=0.5, anchor="center")
        frame.bind("<Configure>", lambda e: self.schedule_image_update())
        self.view_frame = frame
        self.preview_label = label
        self.schedule_image_update()

    def _open_display_window(self):
        if self.display_window is None or not self.display_window.winfo_exists():
            self.display_window = tk.Toplevel(self)
            self.display_window.title(DISPLAY_WINDOW_TITLE)
            self.display_window.geometry(DISPLAY_WINDOW_GEOMETRY)
            self.display_window.protocol("WM_DELETE_WINDOW", self._on_display_window_close)
        return self.display_window

    def _on_display_window_close(self):
        self.display_separate_var.set(False)
        self._embed_display_window()

    def toggle_display_window(self):
        if self.display_separate_var.get():
            self._show_display_window()
        else:
            self._embed_display_window()

    def toggle_obs_overlay(self):
        self.schedule_image_update()

    def toggle_timer_visibility(self):
        self.schedule_image_update()

    def _show_display_window(self):
        if self.view_frame is not None:
            self.view_frame.destroy()
        win = self._open_display_window()
        self._build_view_frame(win, fill="both", expand=True)

    def _embed_display_window(self):
        if self.view_frame is not None:
            self.view_frame.destroy()
        if self.display_window is not None and self.display_window.winfo_exists():
            self.display_window.destroy()
        self.display_window = None
        before = self.control_stage if self.control_stage and self.control_stage.winfo_manager() else None
        self._build_view_frame(self, fill="x", expand=False, before=before)

    def _capture_control_base_size(self):
        if self.controls_container is None:
            return
        self.update_idletasks()
        w = self.controls_container.winfo_reqwidth()
        h = self.controls_container.winfo_reqheight()
        if w <= 0 or h <= 0:
            return
        self.control_base_size = (w, h)

    def _schedule_control_scale(self, event=None):
        if not CONTROL_AUTO_SCALE:
            return
        if self._control_scale_job:
            self.after_cancel(self._control_scale_job)
        self._control_scale_job = self.after(50, self._apply_control_scale)

    def _apply_control_scale(self):
        self._control_scale_job = None
        if not CONTROL_AUTO_SCALE or self.control_stage is None or not self.control_stage.winfo_exists():
            return
        if self.control_base_size is None:
            self._capture_control_base_size()
        if self.control_base_size is None:
            return
        base_w, base_h = self.control_base_size
        avail_w = self.control_stage.winfo_width()
        avail_h = self.control_stage.winfo_height()
        if avail_w <= 0 or avail_h <= 0:
            return
        scale = min(avail_w / base_w, avail_h / base_h) * CONTROL_SCALE_FILL
        scale = max(CONTROL_SCALE_MIN, min(CONTROL_SCALE_MAX, scale))
        if abs(scale - self.control_scale) < 0.02:
            return
        self.control_scale = scale
        self.tk.call("tk", "scaling", self.base_tk_scaling * scale)
        self.update_idletasks()

    def setup_ui(self):
        self.view_frame = None
        self.preview_label = None
        self.control_stage = tk.Frame(self, bg="#eee")
        self.control_stage.pack(side="top", fill="both", expand=True)
        self.control_stage.grid_rowconfigure(0, weight=1)
        self.control_stage.grid_rowconfigure(1, weight=0)
        self.control_stage.grid_rowconfigure(2, weight=1)
        self.control_stage.grid_columnconfigure(0, weight=1)
        self.control_stage.grid_columnconfigure(1, weight=0)
        self.control_stage.grid_columnconfigure(2, weight=1)

        self.controls_container = tk.Frame(self.control_stage, bg="#eee")
        self.controls_container.option_add("*Font", self.default_font)
        self.controls_container.grid(row=1, column=1, padx=CONTROL_STAGE_PADDING, pady=CONTROL_STAGE_PADDING)
        if self.display_separate_var.get():
            self._show_display_window()
        else:
            self._embed_display_window()

        nav = tk.Frame(self.controls_container, bg="#333")
        self.tool = tk.Frame(self.controls_container, bg="#eee")
        self.ctrl_area = tk.Frame(self.controls_container, bg="#eee")
        
        self.score_ctrls = tk.Frame(self.ctrl_area, bg="#eee")
        self.course_ctrls = tk.Frame(self.ctrl_area, bg="#eee")
        self.special_ctrls = tk.Frame(self.ctrl_area, bg="#eee") 
        self.sf_follow_ctrls = tk.Frame(self.ctrl_area, bg="#eee") 

        nav.pack(side="top", fill="x")
        self.tab_btns = []
        for i in range(NUM_GROUPS):
            b = tk.Button(nav, text=f"第{i+1}組", width=6, command=lambda idx=i: self.switch_tab(idx))
            b.pack(side="left", padx=1, pady=5); self.tab_btns.append(b)
        
        self.btn_3rd = tk.Button(nav, text="3rd選択", width=8, bg="#55acee", fg="white", command=lambda: self.switch_tab("3RD"))
        self.btn_3rd.pack(side="left", padx=5, pady=5); self.tab_btns.append(self.btn_3rd)
        self.course_btns = []
        for c_name in ["10by10", "Swedish10", "Freeze10", "10up-down"]:
            b = tk.Button(nav, text=c_name, width=8, bg="#ff8c00", fg="white", command=lambda n=c_name: self.switch_tab(n))
            b.pack(side="left", padx=1, pady=5); self.course_btns.append(b)
        
        self.btn_semi = tk.Button(nav, text="Semi", width=6, bg="#9932cc", fg="white", command=lambda: self.switch_tab("SEMI"))
        self.btn_semi.pack(side="left", padx=5, pady=5); self.tab_btns.append(self.btn_semi)
        
        self.btn_sf_follow = tk.Button(nav, text="SFフォロー", width=9, bg="#4682b4", fg="white", command=lambda: self.switch_tab("SF_FOLLOW"))
        self.btn_sf_follow.pack(side="left", padx=1, pady=5); self.tab_btns.append(self.btn_sf_follow)

        self.btn_final = tk.Button(nav, text="Final", width=6, bg="#dc143c", fg="white", command=lambda: self.switch_tab("FINAL"))
        self.btn_final.pack(side="left", padx=1, pady=5); self.tab_btns.append(self.btn_final)
        self.btn_extra = tk.Button(nav, text="Ex 2nd", width=6, bg="#2e8b57", fg="white", command=lambda: self.switch_tab("EXTRA"))
        self.btn_extra.pack(side="left", padx=1, pady=5); self.tab_btns.append(self.btn_extra)

        self.tool.pack(side="top", fill="x")
        tk.Button(self.tool, text="48名読込", command=self.load_all_csv).pack(side="left", padx=5)
        tk.Button(self.tool, text="問題読込", command=self.load_questions_csv).pack(side="left", padx=5)
        
        tk.Checkbutton(self.tool, text="OBS合成UI", variable=self.obs_overlay_var,
                       command=self.toggle_obs_overlay, bg="#eee").pack(side="right", padx=10)
        tk.Checkbutton(self.tool, text="表示分離", variable=self.display_separate_var,
                       command=self.toggle_display_window, bg="#eee").pack(side="right", padx=10)
        timer_f = tk.Frame(self.tool, bg="#ddd", padx=5); timer_f.pack(side="right", padx=10)
        tk.Checkbutton(timer_f, text="表示", variable=self.timer_visible_var,
                       command=self.toggle_timer_visibility, bg="#ddd").pack(side="right", padx=(8, 0))
        tk.Label(timer_f, text="Timer:", bg="#ddd").pack(side="left")
        self.entry_min = tk.Entry(timer_f, width=3); self.entry_min.insert(0, str(TIMER_DEFAULT_MIN)); self.entry_min.pack(side="left")
        tk.Label(timer_f, text=":", bg="#ddd").pack(side="left")
        self.entry_sec = tk.Entry(timer_f, width=3); self.entry_sec.insert(0, str(TIMER_DEFAULT_SEC).zfill(2)); self.entry_sec.pack(side="left")
        tk.Button(timer_f, text="Set", command=self.set_timer_val, width=4).pack(side="left", padx=2)
        tk.Button(timer_f, text="Start/Stop", command=self.toggle_timer, width=8).pack(side="left", padx=2)
        tk.Button(timer_f, text="Reset", command=self.reset_timer, width=5).pack(side="left", padx=2)

        self.ctrl_area.pack(side="top", fill="both", expand=True)
        
        prog = tk.Frame(self.score_ctrls, bg="#eee"); prog.pack(side="top", fill="x")
        tk.Button(prog, text="スルー", width=12, command=self.next_question_manual).pack(side="left", padx=5, pady=2)
        tk.Button(prog, text="Undo", width=12, command=self.undo).pack(side="left", pady=2)
        
        self.btn_grid_f = tk.Frame(self.score_ctrls, bg="#eee"); self.btn_grid_f.pack(side="top", fill="x")
        self.player_widgets = []
        for i in range(12):
            f = tk.Frame(self.btn_grid_f, relief="groove", borderwidth=1, bg="white"); f.grid(row=i//6, column=i%6, padx=1, pady=1, sticky="nsew")
            n = tk.Label(f, text="", font=self.bold_font, bg="white", width=18); n.pack(side="top")
            s = tk.Label(f, text="", font=self.bold_font, bg="white"); s.pack(side="top")
            b_row = tk.Frame(f, bg="white"); b_row.pack(side="top", fill="x")
            
            tk.Button(b_row, text="O", bg="#dfd", width=3, command=lambda x=i: self.act(x, "o")).pack(side="left", expand=True)
            tk.Button(b_row, text="X", bg="#fdd", width=3, command=lambda x=i: self.act(x, "x")).pack(side="left", expand=True)
            tk.Button(b_row, text="R", bg="#eee", width=2, command=lambda x=i: self.act(x, "r")).pack(side="left", expand=True)
            tk.Button(b_row, text="W", bg="#ffd700", width=2, command=lambda x=i: self.act_win_lose(x, "win")).pack(side="left", expand=True)
            tk.Button(b_row, text="L", bg="#888", width=2, command=lambda x=i: self.act_win_lose(x, "lose")).pack(side="left", expand=True)

            self.player_widgets.append({"name": n, "score": s, "frame": f})

        display_ctrl = tk.LabelFrame(self.course_ctrls, text="中央モニター表示操作", bg="#eee", padx=10, pady=5)
        display_ctrl.pack(side="top", fill="x", padx=10, pady=5)
        for c in COURSES:
            tk.Button(display_ctrl, text=c, width=10, command=lambda name=c: self.set_display_course_3rd(name)).pack(side="left", padx=2)
        tk.Button(self.course_ctrls, text="進出者を更新", command=self.update_3rd_list).pack(pady=2)
        c_grid = tk.Frame(self.course_ctrls, bg="#eee"); c_grid.pack()
        self.cb_list = []
        for i in range(20):
            f = tk.Frame(c_grid, relief="groove", borderwidth=1, bg="white"); f.grid(row=i//10, column=i%10, padx=1, pady=1)
            tk.Label(f, text=f"{i+1}", font=self.small_font, bg="white").pack(side="left")
            cb = ttk.Combobox(f, values=COURSES, state="readonly", width=8); cb.set("未選択"); cb.bind("<<ComboboxSelected>>", lambda e, idx=i: self.select_course_3rd(idx)); cb.pack(side="left"); self.cb_list.append(cb)

        self.special_ctrl_lbl = tk.Label(self.special_ctrls, text="", font=self.bold_font, bg="#eee")
        self.special_ctrl_lbl.pack(side="left", padx=10)
        
        tk.Button(self.special_ctrls, text="参加者手動設定", bg="#ffcccc", command=self.open_manual_entry_window).pack(side="left", padx=5)

        self.semi_btns_f = tk.Frame(self.special_ctrls, bg="#eee")
        self.semi_set_var = tk.IntVar(value=1)
        tk.Radiobutton(self.semi_btns_f, text="Set 1", variable=self.semi_set_var, value=1, command=self.change_semi_set).pack(side="left")
        tk.Radiobutton(self.semi_btns_f, text="Set 2", variable=self.semi_set_var, value=2, command=self.change_semi_set).pack(side="left")
        tk.Radiobutton(self.semi_btns_f, text="Set 3", variable=self.semi_set_var, value=3, command=self.change_semi_set).pack(side="left")
        
        self.btn_sf_hide = tk.Button(self.semi_btns_f, text="Score Hide", command=self.toggle_sf_hide, relief="raised")
        self.btn_sf_hide.pack(side="left", padx=10)

        self.final_btn_f = tk.Frame(self.special_ctrls, bg="#eee")
        tk.Button(self.final_btn_f, text="次のセットへ (Reset)", command=self.reset_final_set).pack(side="left")

        # --- SF Follow Controls ---
        tk.Label(self.sf_follow_ctrls, text="SFフォローモード (3問表示)", font=self.bold_font, bg="#eee").pack(pady=5)
        f_range = tk.Frame(self.sf_follow_ctrls, bg="#eee"); f_range.pack(pady=5)
        tk.Label(f_range, text="開始No:", bg="#eee").pack(side="left")
        self.e_sf_start = tk.Entry(f_range, width=5); self.e_sf_start.pack(side="left", padx=2)
        tk.Label(f_range, text="終了No:", bg="#eee").pack(side="left")
        self.e_sf_end = tk.Entry(f_range, width=5); self.e_sf_end.pack(side="left", padx=2)
        tk.Button(f_range, text="範囲設定", command=self.set_sf_follow_range).pack(side="left", padx=10)
        
        f_nav = tk.Frame(self.sf_follow_ctrls, bg="#eee"); f_nav.pack(pady=5)
        tk.Button(f_nav, text="<< Prev 3", width=12, command=self.prev_sf_follow).pack(side="left", padx=5)
        tk.Button(f_nav, text="Next 3 >>", width=12, command=self.next_sf_follow).pack(side="left", padx=5)

        self.score_ctrls.pack(fill="both", expand=True)
        self.update_idletasks()
        self._capture_control_base_size()
        self.control_stage.bind("<Configure>", self._schedule_control_scale)
        self.schedule_image_update()

    def refresh_ui(self):
        players = self.get_current_mode_players()
        for i in range(12):
            if i < len(players):
                p = players[i]
                self.player_widgets[i]["frame"].grid()
                
                status_suffix = ""
                if self.mode == "SEMI":
                    if p["semi_status"] == "win": status_suffix = " [WIN]"
                    elif p["semi_status"] == "lose": status_suffix = " [LOSE]"
                
                self.player_widgets[i]["name"].config(text=f"{p['rank']} {p['name']}{status_suffix}")
                
                txt = ""
                if self.mode == "10by10": txt = f"{p['10by10_o'] * p['10by10_x']} ({p['10by10_o']}○{p['10by10_x']}×)"
                elif self.mode == "Swedish10": txt = f"{p['Swedish10_o']}○ {p['Swedish10_x']}×"
                elif self.mode == "Freeze10": txt = f"{p['Freeze10_o']}○ {p['Freeze10_x']}×"
                elif self.mode == "10up-down": txt = f"{p['10up-down_score']}pts {p['10up-down_wrong']}×"
                elif self.mode == "SEMI": txt = f"{p['semi_score']} pts"
                elif self.mode == "FINAL": txt = f"{p['final_sets_won']}Sets ({p['final_curr_o']}○{p['final_curr_x']}×)"
                elif self.mode == "EXTRA": txt = f"{p['extra_score']}pts {p['extra_wrong']}×"
                else: txt = f"{p['score']}pts / {p['wrong']}×"
                
                self.player_widgets[i]["score"].config(text=txt)
            else:
                self.player_widgets[i]["frame"].grid_remove()
        self.schedule_image_update()

    def set_timer_val(self):
        try:
            m = int(self.entry_min.get())
            s = int(self.entry_sec.get())
            self.timer_seconds = m * 60 + s
            self.refresh_ui()
        except: pass

    def toggle_timer(self):
        self.timer_running = not self.timer_running
        if self.timer_running: self.update_timer_loop()

    def reset_timer(self):
        self.timer_running = False
        self.set_timer_val()
        self.refresh_ui()

    def update_timer_loop(self):
        if self.timer_running:
            if self.timer_seconds > 0:
                self.timer_seconds -= 1
                self.refresh_ui()
                self.after(1000, self.update_timer_loop)
            else:
                self.timer_running = False 
                self.refresh_ui()

    def get_timer_str(self):
        m = self.timer_seconds // 60
        s = self.timer_seconds % 60
        return f"{m:02}:{s:02}"

    def update_3rd_list(self):
        winners = {i:[] for i in range(1, 13)}
        for g in range(4):
            for p in self.all_groups_data[g]:
                if p["win_order"] > 0: winners[p["win_order"]].append(p)
        selected = []
        for o in range(1, 13): selected.extend(sorted(winners[o], key=lambda x: x["rank_num"]))
        while len(selected) < 20: selected.append(get_empty_player(99))
        self.players_3rd_20 = selected[:20]; self.refresh_ui()

    def select_course_3rd(self, p_idx):
        c_name = self.cb_list[p_idx].get()
        if c_name == "未選択": self.player_selections_3rd.pop(p_idx, None)
        else: self.player_selections_3rd[p_idx] = COURSES.index(c_name)
        self.refresh_ui()

    def set_display_course_3rd(self, name):
        self.current_selected_course_name_3rd = name; self.refresh_ui()

    def change_semi_set(self):
        self.semi_set_idx = self.semi_set_var.get()
        self.refresh_ui()

    def reset_final_set(self):
        for p in self.players_final_3:
            p["final_curr_o"] = 0
            p["final_curr_x"] = 0
            p["final_set_lost"] = False
        self.refresh_ui()

    def set_sf_follow_range(self):
        try:
            s = int(self.e_sf_start.get())
            e = int(self.e_sf_end.get())
            if s > e: return
            self.sf_follow_start = s - 1 # 0-based
            self.sf_follow_end = e - 1
            self.sf_follow_cursor = 0
            self.refresh_ui()
        except: pass

    def next_sf_follow(self):
        # Move cursor by 3
        if self.sf_follow_start + self.sf_follow_cursor + 3 <= self.sf_follow_end + 3: # Allow to show last partial page
             self.sf_follow_cursor += 3
             self.refresh_ui()

    def prev_sf_follow(self):
        if self.sf_follow_cursor >= 3:
            self.sf_follow_cursor -= 3
            self.refresh_ui()

    def toggle_sf_hide(self):
        self.sf_hide_scores = not self.sf_hide_scores
        if self.sf_hide_scores:
            self.btn_sf_hide.config(relief="sunken", bg="#aaa")
        else:
            self.btn_sf_hide.config(relief="raised", bg="SystemButtonFace")
        self.refresh_ui()

    def open_manual_entry_window(self):
        if self.mode not in ["SEMI", "EXTRA"]:
            messagebox.showinfo("info", "このモードでは手動設定は使用しません")
            return
            
        win = tk.Toplevel(self)
        win.title("参加者・スコア手動設定")
        win.geometry("500x600")
        win.transient(self)
        win.grab_set()
        
        current_list = self.players_semi_9 if self.mode == "SEMI" else self.players_extra_12
        entries = []
        
        tk.Label(win, text=f"{self.mode} 手動編集").pack(pady=5)
        
        f_list = tk.Frame(win)
        f_list.pack(fill="both", expand=True, padx=10)
        
        tk.Label(f_list, text="No.", width=5).grid(row=0, column=1)
        tk.Label(f_list, text="画像", width=5).grid(row=0, column=2)
        tk.Label(f_list, text="正解/Pts", width=8).grid(row=0, column=3)
        tk.Label(f_list, text="誤答/✕", width=8).grid(row=0, column=4)

        first_entry = None
        for i, p in enumerate(current_list):
            tk.Label(f_list, text=f"{i+1}").grid(row=i+1, column=0, pady=2)
            
            e_rank = tk.Entry(f_list, width=6)
            if p["rank_num"] != 99: e_rank.insert(0, str(p["rank_num"]))
            e_rank.grid(row=i+1, column=1)
            
            def sel_img(idx=i):
                fp = filedialog.askopenfilename(filetypes=[("Image Files", "*.jpg;*.png;*.jpeg")])
                if fp:
                    entries[idx][4] = fp
                    messagebox.showinfo("OK", f"Slot {idx+1} に画像を設定しました")

            btn_img = tk.Button(f_list, text="...", command=lambda ix=i: sel_img(ix), width=3)
            btn_img.grid(row=i+1, column=2)
            
            e_o = tk.Entry(f_list, width=8)
            val_o = 0
            if self.mode == "SEMI": val_o = p["semi_score"]
            elif self.mode == "EXTRA": val_o = p["extra_score"]
            e_o.insert(0, str(val_o))
            e_o.grid(row=i+1, column=3)
            
            e_x = tk.Entry(f_list, width=8)
            val_x = 0
            if self.mode == "EXTRA": val_x = p["extra_wrong"]
            e_x.insert(0, str(val_x))
            e_x.grid(row=i+1, column=4)
            if self.mode == "SEMI": e_x.config(state="disabled")

            entries.append([e_rank, e_o, e_x, btn_img, p["photo_path"]])
            
            if i == 0: first_entry = e_rank
            
        def apply():
            loaded_check = any(p["name"] != "---" for g in self.all_groups_data for p in g)
            if not loaded_check:
                messagebox.showerror("エラー", "まだCSVファイルが読み込まれていません。\n先に「48名読込」を行ってください。")
                win.destroy()
                return

            new_list = []
            for item in entries:
                e_r, e_o, e_x, _, photo_p = item
                rank_val = e_r.get().strip()
                player_obj = get_empty_player(99)
                
                if rank_val.isdigit():
                    r_num = int(rank_val)
                    for g in range(NUM_GROUPS):
                        for pl in self.all_groups_data[g]:
                            if pl["rank_num"] == r_num:
                                player_obj = copy.deepcopy(pl)
                                break
                
                player_obj["photo_path"] = photo_p

                try:
                    s_val = int(e_o.get().strip()) if e_o.get().strip() else 0
                    w_val = int(e_x.get().strip()) if e_x.get().strip() else 0
                except:
                    s_val, w_val = 0, 0

                if self.mode == "SEMI":
                    player_obj["semi_score"] = s_val
                elif self.mode == "EXTRA":
                    player_obj["extra_score"] = s_val
                    player_obj["extra_wrong"] = w_val
                
                new_list.append(player_obj)
            
            if self.mode == "SEMI":
                self.players_semi_9 = new_list
            else:
                self.players_extra_12 = new_list
            
            self.refresh_ui()
            win.destroy()
            
        tk.Button(win, text="更新", command=apply, bg="#aecf00", width=20).pack(pady=10)
        if first_entry: first_entry.focus_set()

    def next_question_manual(self):
        self.save_history()
        players = self.get_current_mode_players()
        self._process_end_of_question(players)
        self.refresh_ui()

    def switch_tab(self, target):
        self.mode = "SCORE" if isinstance(target, int) else target
        
        self.score_ctrls.pack_forget()
        self.course_ctrls.pack_forget()
        self.special_ctrls.pack_forget()
        self.sf_follow_ctrls.pack_forget()
        self.semi_btns_f.pack_forget()
        self.final_btn_f.pack_forget()

        if target == "3RD":
            self.course_ctrls.pack(fill="both", expand=True)
            self.update_3rd_list()
        elif target == "SF_FOLLOW":
            self.sf_follow_ctrls.pack(fill="both", expand=True)
        else:
            if target == "SEMI":
                self.special_ctrl_lbl.config(text="準決勝 Nine Hundred")
                self.special_ctrls.pack(side="top", fill="x")
                self.semi_btns_f.pack(side="left")
            elif target == "FINAL":
                self.special_ctrl_lbl.config(text="決勝 Triple Seven")
                self.special_ctrls.pack(side="top", fill="x")
                self.final_btn_f.pack(side="left")
                
                final_candidates = [get_empty_player(99) for _ in range(3)]
                for p in self.players_semi_9:
                    if p["semi_status"] == "win":
                        s_idx = p.get("semi_exit_set", 0)
                        if 1 <= s_idx <= 3:
                            final_candidates[s_idx - 1] = p
                self.players_final_3 = final_candidates

            elif target == "EXTRA":
                self.special_ctrl_lbl.config(text="Extra Round 2nd step")
                self.special_ctrls.pack(side="top", fill="x")
            elif target not in ["10by10", "Swedish10", "Freeze10", "10up-down"]: 
                self.current_group_idx = target
            
            self.score_ctrls.pack(fill="both", expand=True)
        
        all_btns = self.tab_btns + self.course_btns
        for b in all_btns:
            is_active = False
            if b in self.tab_btns:
                idx = self.tab_btns.index(b)
                if target == idx: is_active = True
                if target == "3RD" and b == self.btn_3rd: is_active = True
                if target == "SEMI" and b == self.btn_semi: is_active = True
                if target == "FINAL" and b == self.btn_final: is_active = True
                if target == "EXTRA" and b == self.btn_extra: is_active = True
                if target == "SF_FOLLOW" and b == self.btn_sf_follow: is_active = True
            else:
                if b["text"] == target: is_active = True
            
            base_col = "#bbb"
            if b == self.btn_3rd: base_col = "#55acee"
            elif b in self.course_btns: base_col = "#ff8c00"
            elif b == self.btn_semi: base_col = "#9932cc"
            elif b == self.btn_final: base_col = "#dc143c"
            elif b == self.btn_extra: base_col = "#2e8b57"
            elif b == self.btn_sf_follow: base_col = "#4682b4"
            
            b.config(bg="#555" if is_active else base_col)
        
        self.refresh_ui()

    def get_current_mode_players(self):
        if self.mode in ["10by10", "Swedish10", "Freeze10", "10up-down"]:
            selected = [self.players_3rd_20[i] for i, c_idx in self.player_selections_3rd.items() if COURSES[c_idx] == self.mode]
            return sorted(selected, key=lambda x: x["rank_num"])[:5]
        elif self.mode == "SEMI": return self.players_semi_9
        elif self.mode == "FINAL": return self.players_final_3
        elif self.mode == "EXTRA": return self.players_extra_12
        return self.all_groups_data[self.current_group_idx]

    def load_all_csv(self):
        fp = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not fp: return
        success = False
        for enc in ['utf-8-sig', 'shift_jis', 'cp932', 'utf-8']:
            try:
                with open(fp, "r", encoding=enc) as f:
                    reader = csv.reader(f)
                    for row in reader:
                        if not row: continue
                        if len(row) == 1: row = row[0].replace('"', '').split(',')
                        if len(row) < 3: continue
                        match = re.search(r'\d+', str(row[0]))
                        if match:
                            r_num = int(match.group())
                            g_idx = (r_num - 1) % 4
                            p_idx = (r_num - 1) // 4
                            if 0 <= g_idx < NUM_GROUPS and 0 <= p_idx < 12:
                                self.all_groups_data[g_idx][p_idx]["univ"] = str(row[1])
                                self.all_groups_data[g_idx][p_idx]["name"] = str(row[2])
                                success = True
                    if success: break
            except Exception: continue 
        self.refresh_ui()

    def load_questions_csv(self):
        fp = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if fp:
            try:
                with open(fp, "r", encoding='utf-8-sig') as f:
                    self.questions = [{"q": r[0], "a": r[1]} for r in csv.reader(f) if len(r) >= 2]
                self.current_q_idx = 0; self.refresh_ui()
            except: pass

    # --- 修正: W/Lボタンの処理を汎用化 ---
    def act_win_lose(self, idx, status):
        """全モード共通のW/Lボタン処理"""
        self.save_history()
        players = self.get_current_mode_players()
        if idx >= len(players): return
        p = players[idx]
        
        if self.mode == "SEMI":
            if status == "win":
                p["semi_status"] = "win"
                p["semi_exit_set"] = self.semi_set_idx
            elif status == "lose":
                p["semi_status"] = "lose"
                p["semi_exit_set"] = self.semi_set_idx
        
        elif self.mode == "EXTRA":
            if status == "win":
                p["extra_score"] = 5
                # 修正: 勝者が確定していない場合のみ勝ち抜け順を設定
                if p["win_order_extra"] == 0:
                      p["win_order_extra"] = len([pl for pl in players if pl["win_order_extra"] > 0]) + 1
            elif status == "lose":
                p["extra_wrong"] = 1
        
        elif self.mode == "10by10":
            if status == "win":
                p["10by10_o"] = 10; p["10by10_x"] = 10
                if p["win_order_10by10"] == 0:
                      p["win_order_10by10"] = len([pl for pl in players if pl["win_order_10by10"] > 0]) + 1
            elif status == "lose":
                p["10by10_x"] = 0 # 失格条件
        
        elif self.mode == "Swedish10":
            if status == "win":
                p["Swedish10_o"] = 10
                if p["win_order_Swedish10"] == 0:
                      p["win_order_Swedish10"] = len([pl for pl in players if pl["win_order_Swedish10"] > 0]) + 1
            elif status == "lose":
                p["Swedish10_x"] = 10 # 失格条件
        
        elif self.mode == "10up-down":
            if status == "win":
                p["10up-down_score"] = 10
                if p["win_order_10up-down"] == 0:
                      p["win_order_10up-down"] = len([pl for pl in players if pl["win_order_10up-down"] > 0]) + 1
            elif status == "lose":
                p["10up-down_wrong"] = 2 # 失格条件
                
        # 修正: モードが "SCORE" (2R) の場合の判定を追加
        elif self.mode == "SCORE" or isinstance(self.mode, int): # 2R
            if status == "win":
                p["score"] = WIN_POINTS
                # 修正: 勝者が確定していない場合のみ勝ち抜け順を設定
                if p["win_order"] == 0:
                    p["win_order"] = len([pl for pl in self.all_groups_data[self.current_group_idx] if pl["win_order"] > 0]) + 1
            elif status == "lose":
                p["wrong"] = LOSE_WRONGS

        self.refresh_ui()

    def act(self, idx, t):
        self.save_history()
        players = self.get_current_mode_players()
        if idx >= len(players): return
        p = players[idx]
        
        if self.mode == "Freeze10" and p["Freeze10_freeze"] > 0: return 
        if self.mode == "FINAL" and p["final_set_lost"]: return
        if self.mode == "EXTRA" and p["extra_wrong"] >= 1: return 

        advance = False
        
        if self.mode == "10by10":
            if t == "o": 
                p["10by10_o"] += 1
                if p["10by10_o"] * p["10by10_x"] >= 100 and p["win_order_10by10"] == 0:
                    p["win_order_10by10"] = len([pl for pl in players if pl["win_order_10by10"] > 0]) + 1
                advance = True
            elif t == "x": p["10by10_x"] -= 1; advance = True
            elif t == "r": p["10by10_o"] = 0; p["10by10_x"] = 10; p["win_order_10by10"] = 0
        
        elif self.mode == "10up-down":
            if t == "o":
                p["10up-down_score"] += 1
                if p["10up-down_score"] >= 10 and p["win_order_10up-down"] == 0:
                    p["win_order_10up-down"] = len([pl for pl in players if pl["win_order_10up-down"] > 0]) + 1
                advance = True
            elif t == "x":
                p["10up-down_score"] = 0 
                p["10up-down_wrong"] += 1
                advance = True
            elif t == "r":
                p["10up-down_score"] = 0; p["10up-down_wrong"] = 0; p["win_order_10up-down"] = 0

        elif self.mode == "Freeze10":
            if t == "o":
                p["Freeze10_o"] += 1
                if p["Freeze10_o"] >= 10 and p["win_order_Freeze10"] == 0:
                    p["win_order_Freeze10"] = len([pl for pl in players if pl["win_order_Freeze10"] > 0]) + 1
                advance = True
            elif t == "x":
                p["Freeze10_x"] += 1
                p["Freeze10_freeze"] = p["Freeze10_x"] + 1 
                advance = True
            elif t == "r":
                p["Freeze10_o"] = 0; p["Freeze10_x"] = 0; p["Freeze10_freeze"] = 0; p["win_order_Freeze10"] = 0

        elif self.mode == "Swedish10":
            if t == "o":
                p["Swedish10_o"] += 1
                if p["Swedish10_o"] >= 10 and p["win_order_Swedish10"] == 0:
                    p["win_order_Swedish10"] = len([pl for pl in players if pl["win_order_Swedish10"] > 0]) + 1
                advance = True
            elif t == "x":
                pen = 0; sc = p["Swedish10_o"]
                if sc == 0: pen = 1
                elif 1 <= sc <= 2: pen = 2
                elif 3 <= sc <= 5: pen = 3
                elif 6 <= sc <= 9: pen = 4
                p["Swedish10_x"] += pen
                advance = True
            elif t == "r":
                p["Swedish10_o"] = 0; p["Swedish10_x"] = 0; p["win_order_Swedish10"] = 0

        elif self.mode == "SEMI":
            conf = SEMI_RULES[self.semi_set_idx]
            if t == "o": p["semi_score"] += conf["correct"]; advance = True
            elif t == "x": p["semi_score"] += conf["wrong"]; advance = True
            elif t == "r": p["semi_score"] = 0; p["semi_status"] = "active"

        elif self.mode == "FINAL":
            if t == "o":
                if p["final_curr_o"] < 7:
                    p["final_curr_o"] += 1
                    if p["final_curr_o"] == 7:
                        p["final_sets_won"] += 1
                        messagebox.showinfo("Set Winner", f"{p['name']} がセット獲得！")
                        if p["final_sets_won"] >= 3: p["win_order_final"] = 1
                    
                    # 7に到達した場合もそうでない場合も、問題を送るためにTrueにする
                    advance = True
            elif t == "x":
                p["final_curr_x"] += 1
                if p["final_curr_x"] >= 3: p["final_set_lost"] = True
                advance = True
            elif t == "r":
                p["final_sets_won"] = 0; self.reset_final_set()

            active = [pl for pl in players if not pl["final_set_lost"]]
            if len(active) == 1:
                winner = active[0]
                winner["final_sets_won"] += 1
                
                winner["final_curr_o"] = 7
                messagebox.showinfo("Set Winner", f"他者失格により {winner['name']} がセット獲得！")
                
                if winner["final_sets_won"] >= 3: winner["win_order_final"] = 1

        elif self.mode == "EXTRA":
            if t == "o":
                p["extra_score"] += 1
                if p["extra_score"] >= 5:
                    # 修正: 勝ち抜け順の重複加算を防ぐ
                    if p["win_order_extra"] == 0:
                        p["win_order_extra"] = 1
                advance = True
            elif t == "x":
                p["extra_wrong"] += 1
                advance = True
            elif t == "r":
                p["extra_score"] = 0; p["extra_wrong"] = 0; p["win_order_extra"] = 0
            
            active = [pl for pl in players if pl["name"] != "---" and pl["extra_wrong"] == 0]
            already_win = any(pl["win_order_extra"] > 0 for pl in players)
            
            if not already_win and len(active) == 1:
                winner = active[0]
                winner["win_order_extra"] = 1
                messagebox.showinfo("Winner", f"他者全滅により {winner['name']} が復活！")

        else: # 2nd Round
            if t == "o":
                p["score"] += (2 if p["rento"] else 1)
                if p["rento"]: p["rento"] = False
                else:
                    if p["score"] < WIN_POINTS: p["rento"] = True
                
                if p["score"] >= WIN_POINTS:
                    # 修正: 勝ち抜け順の重複加算を防ぐ
                    if p["win_order"] == 0:
                         p["win_order"] = len([pl for pl in self.all_groups_data[self.current_group_idx] if pl["win_order"] > 0]) + 1
                
                for other in self.all_groups_data[self.current_group_idx]:
                    if other != p: other["rento"] = False
                self._advance_q()
            elif t == "x":
                p["wrong"] += 1; p["rento"] = False; self._advance_q()
            elif t == "r":
                p["score"], p["wrong"], p["rento"], p["win_order"] = get_advantage_points(p["rank_num"]), 0, False, 0
        
        if advance:
            self._process_end_of_question(players)
            
        self.refresh_ui()

    def _process_end_of_question(self, players):
        if self.mode in ["10by10", "Swedish10", "Freeze10", "10up-down", "SEMI", "FINAL", "EXTRA"]:
            self._advance_q()
            if self.mode == "Freeze10":
                for p in players:
                    if p["Freeze10_freeze"] > 0:
                        p["Freeze10_freeze"] -= 1
        else:
            pass

    def _advance_q(self):
        if self.questions: self.current_q_idx = (self.current_q_idx + 1) % len(self.questions)

    def undo(self): 
        # 修正: キーの判定ロジックを修正 ("SCORE" または int の場合は current_group_idx をキーにする)
        if self.mode == "SCORE" or isinstance(self.mode, int):
            key = self.current_group_idx
        else:
            key = self.mode
            
        stack = self.history_stacks.get(key, [])
        
        if stack:
            last = stack.pop()
            restored_players = last["players"]
            self.current_q_idx = last["q_idx"]
            
            # 各モードの変数に書き戻す
            if self.mode in ["10by10", "Swedish10", "Freeze10", "10up-down"]:
                self.players_3rd_20 = restored_players
            elif self.mode == "SEMI":
                self.players_semi_9 = restored_players
            elif self.mode == "FINAL":
                self.players_final_3 = restored_players
            elif self.mode == "EXTRA":
                self.players_extra_12 = restored_players
            # 修正: "SCORE"の場合もここで処理するように条件を追加
            elif self.mode == "SCORE" or isinstance(self.mode, int):
                self.all_groups_data[self.current_group_idx] = restored_players
                
            self.refresh_ui()

    def save_history(self):
        # 修正: キーの判定ロジックを修正 ("SCORE" または int の場合は current_group_idx をキーにする)
        if self.mode == "SCORE" or isinstance(self.mode, int):
            key = self.current_group_idx
        else:
            key = self.mode
        
        if key not in self.history_stacks:
            self.history_stacks[key] = []
            
        target_list = []
        if self.mode in ["10by10", "Swedish10", "Freeze10", "10up-down"]:
            target_list = self.players_3rd_20
        elif self.mode == "SEMI":
            target_list = self.players_semi_9
        elif self.mode == "FINAL":
            target_list = self.players_final_3
        elif self.mode == "EXTRA":
            target_list = self.players_extra_12
        # 修正: "SCORE"の場合もここで処理するように条件を追加
        elif self.mode == "SCORE" or isinstance(self.mode, int):
            target_list = self.all_groups_data[self.current_group_idx]
            
        self.history_stacks[key].append({
            "players": copy.deepcopy(target_list), 
            "q_idx": self.current_q_idx
        })

    def schedule_image_update(self):
        if self._update_job: self.after_cancel(self._update_job)
        self._update_job = self.after(30, self.update_preview_image)

    def update_preview_image(self):
        if self.view_frame is None or not self.view_frame.winfo_exists():
            return
        if self.mode == "3RD":
            pil_img = self.drawer.generate_image_3rd_round(self.players_3rd_20, self.current_selected_course_name_3rd, self.player_selections_3rd)
        elif self.mode == "SF_FOLLOW":
            pil_img = self.drawer.generate_image_sf_follow(self.questions, self.sf_follow_start, self.sf_follow_end, self.sf_follow_cursor)
        else:
            players = self.get_current_mode_players()
            if self.questions:
                try:
                    q = self.questions[self.current_q_idx]["q"]
                    a = self.questions[self.current_q_idx]["a"]
                except IndexError:
                    q, a = "問題", "解答"
            else:
                q, a = "問題", "解答"
            
            t_str = self.get_timer_str()
            t_alert = (self.timer_seconds == 0)
            pil_img = self.drawer.generate_image(
                players,
                self.current_group_idx,
                q,
                a,
                timer_str=t_str,
                timer_alert=t_alert,
                mode="2R" if self.mode=="SCORE" or isinstance(self.mode, int) else self.mode,
                semi_set_idx=self.semi_set_idx,
                sf_hide_scores=self.sf_hide_scores,
                obs_overlay=self.obs_overlay_var.get(),
                question_index=self.current_q_idx + 1,
                show_timer=self.timer_visible_var.get(),
            )
        
        w = self.view_frame.winfo_width()
        h = self.view_frame.winfo_height()
        if w <= 1: w = 1400
        if h <= 1: h = 300
        r = min(w/IMG_WIDTH, h/IMG_HEIGHT)
        img = pil_img.resize((int(IMG_WIDTH*r), int(IMG_HEIGHT*r)), Image.Resampling.LANCZOS)
        self.tk_img = ImageTk.PhotoImage(img)
        if self.preview_label is not None:
            self.preview_label.config(image=self.tk_img)

if __name__ == "__main__":
    print("アプリケーションを開始します...")
    try:
        QuizApp().mainloop()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("エラーが発生しました。Enterキーで終了します。")
