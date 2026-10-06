# -*- coding: utf-8 -*-
"""
Luna 预翻译工具 v1.0
---------------------
① 译前处理：原文本 txt -> 清洗 -> xlsx（原文 A 列，供翻译）
② 译后处理：译文 xlsx（A 原文 / B 译文）-> 替换规则 -> sqlite（machineTrans）
③ 日志

完全离线运行：不联网、不上传、不写注册表、不读取任何个人目录。

视觉：原生 ttk 控件 + 平角卡片（1px 边框、无圆角、无图标），优先性能与清晰度。
"""
import ctypes
import json
import os
import queue
import re
import shutil
import sqlite3
import sys
import threading
import time
import traceback

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import openpyxl

APP_NAME = 'Luna 预翻译工具'
APP_VERSION = '1.0'
IDEOSPACE = '\u3000'
EMPTY_MT = '{"rengong": ""}'

# ============================ 字体 ============================
# 中文界面：正文 11pt 起，次要文字 10pt。
# 低于 10pt 的中文在 100% / 125% 缩放下笔画会糊成一团，这里统一抬高一级。
FONT_FAMILY = 'Microsoft YaHei UI'
FONT = (FONT_FAMILY, 11)               # 正文 / 输入 / 按钮 / 复选
FONT_B = (FONT_FAMILY, 11, 'bold')     # 表头
FONT_SM = (FONT_FAMILY, 10)            # 次要说明 / 日志
FONT_H3 = (FONT_FAMILY, 11, 'bold')    # 卡片标题
FONT_H2 = (FONT_FAMILY, 13, 'bold')    # 页面标题
FONT_H1 = (FONT_FAMILY, 16, 'bold')    # 应用标题
FONT_LOG = (FONT_FAMILY, 10)           # 日志正文


# ============================ 高 DPI ============================
def enable_dpi():
    """在 Tk 创建前声明 DPI 感知，否则系统会把窗口位图放大，看起来很糊。"""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)   # 仅 Windows 8.1+
        return True
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
            return True
        except Exception:
            return False


def system_scale(root):
    """把 tk scaling 设成真实 DPI 比例，字体按点数渲染出正确像素。"""
    try:
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)   # LOGPIXELSX
        ctypes.windll.user32.ReleaseDC(0, hdc)
        scale = max(1.0, dpi / 96.0)
    except Exception:
        scale = 1.0
    try:
        root.tk.call('tk', 'scaling', scale)
    except Exception:
        pass
    return scale


# ============================ 资源 ============================
def resource_path(rel):
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


BUILTIN_SQLITE = resource_path(os.path.join('resources', 'default.sqlite'))
BUILTIN_XLSX = resource_path(os.path.join('resources', 'default.xlsx'))


def app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


RULES_FILE = os.path.join(app_dir(), '替换规则.json')
UI_FILE = os.path.join(app_dir(), '界面设置.json')

DEFAULT_RULES = [
    {'on': True, 'find': r'\"', 'repl': '"', 're': False},
    {'on': True, 'find': "['‘’〟〝\\[\\]]", 'repl': '"', 're': True},
    {'on': True, 'find': '"', 'repl': r'\"', 're': False},
    {'on': True, 'find': '...', 'repl': '…', 're': False},
    {'on': True, 'find': '・・・', 'repl': '…', 're': False},
    {'on': True, 'find': '俺', 'repl': '我', 're': False},
    {'on': True, 'find': '老子', 'repl': '我', 're': False},
    {'on': True, 'find': '妳', 'repl': '你', 're': False},
    {'on': True, 'find': '。」', 'repl': '」', 're': False},
    {'on': True, 'find': '。)', 'repl': '）', 're': False},
    {'on': True, 'find': '。）', 'repl': '）', 're': False},
    {'on': True, 'find': '^', 'repl': '{"rengong": "', 're': True},
    {'on': True, 'find': '$', 'repl': '"}', 're': True},
]

# ============================ 主题 ============================
# 单一强调色（靛紫），中性色统一走冷灰；状态色只用语义位置。
# 平角设计：卡片 1px 边框、无圆角、无图标，优先清晰度与性能。
THEMES = {
    'light': {
        'bg': '#f2f3f7', 'surface': '#ffffff', 'fg': '#171a21',
        'label': '#4b5260', 'muted': '#6b7280',
        'accent': '#5753d6', 'accent_hover': '#4541c4', 'accent_fg': '#ffffff',
        'accent_soft': '#ecebfb', 'accent_line': '#c9c7f5',
        'border': '#e3e5ee', 'border_strong': '#d2d6e3',
        'input': '#ffffff', 'input_line': '#d8dce8',
        'hover': '#f0f1f6', 'track': '#edeff5', 'shadow': '#dfe2ec',
        'tree_bg': '#ffffff', 'tree_alt': '#f8f9fc', 'tree_sel': '#e8e9fb',
        'tree_head': '#f6f7fb',
        'log_bg': '#fbfbfd', 'log_fg': '#2b2f39',
        'ok': '#1f7a55', 'ok_soft': '#e6f4ec', 'err': '#b3261e', 'err_soft': '#fbeae8',
        'dis_fg': '#a6abb8', 'dis_bg': '#edeff5',
    },
    'dark': {
        'bg': '#1a1a1a', 'surface': '#242424', 'fg': '#e8eaef',
        'label': '#c5c5c5', 'muted': '#9a9a9a',
        'accent': '#8b86f0', 'accent_hover': '#9a95f7', 'accent_fg': '#15151a',
        'accent_soft': '#2e2e38', 'accent_line': '#4a4a78',
        'border': '#343434', 'border_strong': '#404040',
        'input': '#2a2a2a', 'input_line': '#3c3c3c',
        'hover': '#2c2c2c', 'track': '#202020', 'shadow': '#0a0a0a',
        'tree_bg': '#242424', 'tree_alt': '#2a2a2a', 'tree_sel': '#2e2e38',
        'tree_head': '#2c2c2c',
        'log_bg': '#1f1f1f', 'log_fg': '#dfe2ea',
        'ok': '#4cc38a', 'ok_soft': '#26302a', 'err': '#f2777a', 'err_soft': '#33242a',
        'dis_fg': '#6b6b6b', 'dis_bg': '#2a2a2a',
    },
}

TONES = {'idle': ('muted', 'surface'), 'run': ('accent', 'accent_soft'),
         'ok': ('ok', 'ok_soft'), 'err': ('err', 'err_soft')}


# ============================ 控件：平角 / 原生 ============================
class Card(tk.Frame):
    """平角卡片：surface 底 + 1px 边框，内部用 inner Frame 做留白。"""

    def __init__(self, master, colors, pad=16, expand=False):
        super().__init__(master, bg=colors['surface'],
                         highlightbackground=colors['border'], highlightthickness=1, bd=0)
        self.c = colors
        self.inner = tk.Frame(self, bg=colors['surface'])
        self.inner.pack(fill='both', expand=expand, padx=pad, pady=pad)

    def set_colors(self, c):
        self.c = c
        self.configure(bg=c['surface'], highlightbackground=c['border'])
        self.inner.configure(bg=c['surface'])


class Button(ttk.Button):
    """原生按钮：primary（实心强调）/ ghost（描边）。主题由 ttk style 统一处理。"""

    def __init__(self, master, colors, text='', command=None, kind='primary', **kw):
        super().__init__(master, text=text, command=command, **kw)
        self._kind = kind
        self.configure(style='Accent.TButton' if kind == 'primary' else 'TButton')

    def set_state(self, enabled):
        self.state(('!disabled',) if enabled else ('disabled',))

    def set_text(self, text):
        self.configure(text=text)


class Segmented(tk.Frame):
    """分段控件：轨道 + 选中块（原生 Label 平铺，方角）。"""

    def __init__(self, master, colors, items, command=None):
        super().__init__(master, bg=colors['track'],
                         highlightbackground=colors['border'], highlightthickness=1, bd=0)
        self.c = colors
        self.items = items
        self.command = command
        self.index = 0
        self._labels = []
        for i, txt in enumerate(items):
            l = tk.Label(self, text=txt, bg=colors['track'], fg=colors['muted'],
                         font=(FONT_B if i == 0 else FONT), padx=16, pady=10, cursor='hand2')
            l.bind('<Button-1>', lambda e, i=i: self.select(i, fire=True))
            l.pack(side='left', fill='both', expand=True, padx=3, pady=3)
            self._labels.append(l)
        self._paint()

    def _paint(self):
        c = self.c
        for i, l in enumerate(self._labels):
            if i == self.index:
                l.configure(bg=c['accent'], fg=c['accent_fg'], font=FONT_B)
            else:
                l.configure(bg=c['track'], fg=c['muted'], font=FONT)

    def select(self, i, fire=False):
        if i != self.index:
            self.index = i
            self._paint()
        if fire and self.command:
            self.command(i)

    def set_colors(self, c):
        self.c = c
        self.configure(bg=c['track'], highlightbackground=c['border'])
        self._paint()


class ProgressBar(ttk.Progressbar):
    """原生进度条（平角）。"""

    def __init__(self, master, colors, **kw):
        super().__init__(master, orient='horizontal', mode='determinate',
                         maximum=100, style='Horizontal.TProgressbar', **kw)

    def set(self, v):
        self.configure(value=max(0.0, min(100.0, float(v))))


class StatusPill(tk.Frame):
    """状态胶囊：语义色圆点 + 文字（平角方块指示点）。"""

    def __init__(self, master, colors, text='就绪'):
        super().__init__(master, bg=colors['bg'])
        self.c = colors
        self.tone = 'idle'
        self.dot = tk.Frame(self, width=9, height=9, bg=colors['muted'])
        self.dot.pack(side='left', padx=(0, 7))
        self.lbl = tk.Label(self, text=text, bg=colors['bg'], fg=colors['muted'],
                            font=FONT_SM)
        self.lbl.pack(side='left')

    def set(self, text, tone='idle'):
        self.tone = tone
        self.lbl.configure(text=text)
        fg_key, _ = TONES.get(tone, TONES['idle'])
        self.dot.configure(bg=self.c[fg_key])
        self.lbl.configure(fg=self.c[fg_key])

    def set_colors(self, c):
        self.c = c
        fg_key, _ = TONES.get(self.tone, TONES['idle'])
        self.configure(bg=c['bg'])
        self.lbl.configure(bg=c['bg'], fg=c[fg_key])
        self.dot.configure(bg=c[fg_key])


# ============================ 业务逻辑 ============================
def read_text(path):
    with open(path, 'rb') as f:
        raw = f.read()
    if raw.startswith(b'\xef\xbb\xbf'):
        text = raw.decode('utf-8-sig')
    elif raw.startswith(b'\xff\xfe') or raw.startswith(b'\xfe\xff'):
        text = raw.decode('utf-16')
    else:
        text = None
        for enc in ('utf-8', 'cp932', 'gb18030'):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            text = raw.decode('utf-8', 'replace')
    return text.replace('\r\n', '\n').replace('\r', '\n')


def split_lines(text):
    lines = text.split('\n')
    if lines and lines[-1] == '':
        lines.pop()
    return lines


def clean_lines(lines, drop_ideo=True, drop_blank=True, drop_dup=True):
    out, seen = [], set()
    st = {'ideo': 0, 'blank': 0, 'dup': 0}
    for line in lines:
        s = line
        if drop_ideo:
            st['ideo'] += s.count(IDEOSPACE)
            s = s.replace(IDEOSPACE, '')
        if drop_blank and s.strip() == '':
            st['blank'] += 1
            continue
        if drop_dup:
            if s in seen:
                st['dup'] += 1
                continue
            seen.add(s)
        out.append(s)
    return out, st


def build_xlsx(template, out_path, lines):
    wb = openpyxl.load_workbook(template)
    ws = wb.active
    if ws.max_row > 1 or ws['A1'].value not in (None, ''):
        ws.delete_rows(1, ws.max_row)
    for i, text in enumerate(lines, start=1):
        cell = ws.cell(row=i, column=1)
        cell.value = text
        cell.data_type = 's'
    wb.save(out_path)


def fill_sqlite(template, out_path, lines):
    shutil.copyfile(template, out_path)
    con = sqlite3.connect(out_path)
    try:
        cur = con.cursor()
        ids = [r[0] for r in cur.execute('SELECT id FROM artificialtrans ORDER BY id')]
        if len(lines) > len(ids):
            raise RuntimeError(f'文本 {len(lines)} 行，超出模板 {len(ids)} 行')
        if lines:
            cur.executemany('UPDATE artificialtrans SET source = ?, origin = ? WHERE id = ?',
                            [(lines[i], lines[i], ids[i]) for i in range(len(lines))])
            cur.execute('DELETE FROM artificialtrans WHERE id > ?', (ids[len(lines) - 1],))
        else:
            cur.execute('DELETE FROM artificialtrans')
        con.commit()
        con.execute('VACUUM')
        con.commit()
    finally:
        con.close()


def collect(src_dir, recursive, exts):
    items = []
    if recursive:
        for root, _d, files in os.walk(src_dir):
            for fn in files:
                if fn.lower().endswith(exts) and not fn.startswith('~$'):
                    full = os.path.join(root, fn)
                    items.append((full, os.path.relpath(full, src_dir)))
    else:
        for fn in sorted(os.listdir(src_dir)):
            full = os.path.join(src_dir, fn)
            if os.path.isfile(full) and fn.lower().endswith(exts) and not fn.startswith('~$'):
                items.append((full, fn))
    items.sort(key=lambda x: x[1])
    return items


def process_import(src_file, rel, cfg, log):
    raw_lines = split_lines(read_text(src_file))
    lines, st = clean_lines(raw_lines, cfg['drop_ideo'], cfg['drop_blank'], cfg['drop_dup'])
    stem = os.path.splitext(rel)[0]
    p = os.path.join(cfg['out_dir'], stem + '.xlsx')
    os.makedirs(os.path.dirname(p), exist_ok=True)
    build_xlsx(BUILTIN_XLSX, p, lines)
    st['final'] = len(lines)
    log(f'  {rel}: {len(raw_lines)} → {len(lines)} 行'
        f'（全角空格 {st["ideo"]} / 空行 {st["blank"]} / 重复 {st["dup"]}）')
    return st


def apply_rules(text, rules):
    for r in rules:
        if not r.get('on'):
            continue
        text = re.sub(r['find'], r['repl'], text) if r.get('re') else text.replace(r['find'], r['repl'])
    return text


def is_empty_mt(mt):
    if mt == EMPTY_MT:
        return True
    m = re.fullmatch(r'\{"rengong":\s*"(.*)"\}', mt or '', re.S)
    return bool(m) and m.group(1).strip() == ''


def process_translation(xlsx_path, rel, cfg, log):
    stem = os.path.splitext(rel)[0]
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    if ws.max_column < 2:
        raise RuntimeError(f'{rel}: 只有 {ws.max_column} 列，需要 A=原文 B=译文')
    colA = [ws.cell(r, 1).value or '' for r in range(1, ws.max_row + 1)]
    colB = [ws.cell(r, 2).value or '' for r in range(1, ws.max_row + 1)]
    wb.close()

    entries = [apply_rules(v, cfg['rules']) for v in colB]

    out_sqlite = os.path.join(cfg['out_dir'], stem + '.sqlite')
    os.makedirs(os.path.dirname(out_sqlite), exist_ok=True)
    fill_sqlite(BUILTIN_SQLITE, out_sqlite, colA)

    con = sqlite3.connect(out_sqlite)
    try:
        cur = con.cursor()
        rows = cur.execute('SELECT id, origin FROM artificialtrans ORDER BY id').fetchall()
        n = min(len(rows), len(entries))
        if len(rows) != len(entries):
            log(f'  ⚠ {rel}: 底库 {len(rows)} 行 / 译文 {len(entries)} 条，按 {n} 对齐')
        diff_a = sum(1 for i in range(min(n, len(colA))) if colA[i] != rows[i][1])
        filled = 0
        upd = []
        for i in range(n):
            mt = entries[i]
            if is_empty_mt(mt):
                mt = '{"rengong": "' + (rows[i][1] or '') + '"}'
                filled += 1
            upd.append((mt, rows[i][0]))
        cur.executemany('UPDATE artificialtrans SET machineTrans = ? WHERE id = ?', upd)
        con.commit()
        con.execute('VACUUM')
        con.commit()
    finally:
        con.close()

    log(f'  {rel}: {n} 行；回填空译文 {filled} 处；A列与origin不一致 {diff_a} 处')
    return {'rows': n, 'filled': filled, 'diff_a': diff_a}


# ============================ 命令行自检 ============================
def run_cli(kind, src, out_dir, extra=None):
    logs = []

    def log(m):
        logs.append(m)

    if kind == 'import':
        cfg = {'out_dir': out_dir, 'drop_ideo': True, 'drop_blank': True, 'drop_dup': True}
        for full, rel in collect(src, True, ('.txt',)):
            process_import(full, rel, cfg, log)
    else:
        rules = json.load(open(extra, encoding='utf-8')) if (extra and os.path.isfile(extra)) else DEFAULT_RULES
        cfg = {'out_dir': out_dir, 'rules': rules}
        for full, rel in collect(src, True, ('.xlsx',)):
            process_translation(full, rel, cfg, log)

    os.makedirs(out_dir, exist_ok=True)
    with open(os.path.join(out_dir, '_selftest.log'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(logs))
    return 0


# ============================ 规则编辑对话框 ============================
class RuleDialog:
    def __init__(self, parent, colors, title, rule=None):
        self.result = None
        c = colors
        top = tk.Toplevel(parent)
        top.title(title)
        top.configure(background=c['bg'])
        top.transient(parent)
        top.resizable(False, False)
        r = rule or {'on': True, 'find': '', 'repl': '', 're': False}
        v_find = tk.StringVar(value=r.get('find', ''))
        v_repl = tk.StringVar(value=r.get('repl', ''))
        v_re = tk.BooleanVar(value=r.get('re', False))
        v_on = tk.BooleanVar(value=r.get('on', True))

        card = Card(top, c, pad=18)
        card.pack(fill='both', expand=True, padx=16, pady=16)
        inner = card.inner

        head = tk.Frame(inner, background=c['surface'])
        head.pack(fill='x', pady=(0, 16))
        tk.Label(head, text=title, font=FONT_H3, background=c['surface'],
                 foreground=c['fg']).pack(anchor='w')
        tk.Label(head, text='规则按顺序逐条执行，开启正则时请写完整的表达式。',
                 font=FONT_SM, background=c['surface'],
                 foreground=c['muted']).pack(anchor='w', pady=(4, 0))

        def field(label, var):
            box = tk.Frame(inner, background=c['surface'])
            box.pack(fill='x', pady=(0, 14))
            tk.Label(box, text=label, font=FONT_SM, background=c['surface'],
                     foreground=c['label']).pack(anchor='w', pady=(0, 6))
            e = ttk.Entry(box, textvariable=var, font=FONT, width=42)
            e.pack(fill='x')
            return e

        e1 = field('查找', v_find)
        field('替换为', v_repl)

        opts = tk.Frame(inner, background=c['surface'])
        opts.pack(fill='x', pady=(2, 0))
        ttk.Checkbutton(opts, text='正则表达式', variable=v_re,
                        style='TCheckbutton').pack(side='left', padx=(0, 22))
        ttk.Checkbutton(opts, text='启用', variable=v_on,
                        style='TCheckbutton').pack(side='left')

        brow = tk.Frame(inner, background=c['surface'])
        brow.pack(fill='x', pady=(24, 0))

        def ok():
            if not v_find.get():
                messagebox.showwarning(APP_NAME, '查找内容不能为空。', parent=top)
                return
            if v_re.get():
                try:
                    re.compile(v_find.get())
                except re.error as e:
                    messagebox.showwarning(APP_NAME, f'正则无效：{e}', parent=top)
                    return
            self.result = {'on': v_on.get(), 'find': v_find.get(),
                           'repl': v_repl.get(), 're': v_re.get()}
            top.destroy()

        b_ok = Button(brow, c, text='确定', command=ok, kind='primary')
        b_cancel = Button(brow, c, text='取消', command=top.destroy, kind='ghost')
        b_cancel.pack(side='right', padx=(10, 0))
        b_ok.pack(side='right')

        top.bind('<Return>', lambda e: ok())
        top.bind('<Escape>', lambda e: top.destroy())
        e1.focus_set()
        top.update_idletasks()
        try:
            px = parent.winfo_rootx() + (parent.winfo_width() - top.winfo_width()) // 2
            py = parent.winfo_rooty() + max(40, (parent.winfo_height() - top.winfo_height()) // 3)
            top.geometry('+%d+%d' % (max(px, 0), max(py, 0)))
        except Exception:
            pass
        top.grab_set()
        top.wait_window()


# ============================ 主界面 ============================
class App:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.stop_flag = threading.Event()
        self.rules = self.load_rules()
        self.theme = self.load_theme()
        self.c = THEMES[self.theme]
        self.style = ttk.Style()
        self.themed = []           # 带 set_colors(c) 的自定义控件
        self.plain = {'bg': [], 'surface': []}
        root.title(f'{APP_NAME}  v{APP_VERSION}')
        root.geometry('1000x900')
        root.minsize(920, 800)
        root.configure(background=self.c['bg'])

        self.v_src = tk.StringVar()
        self.v_out = tk.StringVar()
        self.v_ideo = tk.BooleanVar(value=True)
        self.v_blank = tk.BooleanVar(value=True)
        self.v_dup = tk.BooleanVar(value=True)
        self.v_recursive = tk.BooleanVar(value=False)

        self.v_tsrc = tk.StringVar()
        self.v_tout = tk.StringVar()
        self.v_trecursive = tk.BooleanVar(value=False)

        self.v_progress = tk.DoubleVar(value=0)
        self.v_status = tk.StringVar(value='就绪')

        self._build()
        self.apply_theme(self.theme)
        self.v_status.trace_add('write', self._sync_status)
        self.v_progress.trace_add('write', self._sync_progress)
        self._sync_status()
        root.protocol('WM_DELETE_WINDOW', self.on_close)
        self._poll()

    # ---------- 配置 ----------
    def load_rules(self):
        try:
            if os.path.isfile(RULES_FILE):
                d = json.load(open(RULES_FILE, encoding='utf-8'))
                if isinstance(d, list) and d:
                    return d
        except Exception:
            pass
        return [dict(r) for r in DEFAULT_RULES]

    def save_rules(self):
        try:
            json.dump(self.rules, open(RULES_FILE, 'w', encoding='utf-8'),
                      ensure_ascii=False, indent=1)
        except Exception:
            pass

    def load_theme(self):
        try:
            if os.path.isfile(UI_FILE):
                return json.load(open(UI_FILE, encoding='utf-8')).get('theme', 'light')
        except Exception:
            pass
        return 'light'

    def save_theme(self):
        try:
            json.dump({'theme': self.theme}, open(UI_FILE, 'w', encoding='utf-8'))
        except Exception:
            pass

    def on_close(self):
        self.save_rules()
        self.save_theme()
        self.root.destroy()

    # ---------- 主题 ----------
    def apply_theme(self, mode):
        self.theme = mode
        c = self.c = THEMES[mode]
        self.root.configure(background=c['bg'])
        s = self.style
        try:
            s.theme_use('clam')
        except Exception:
            pass
        s.configure('.', background=c['bg'], foreground=c['fg'], fieldbackground=c['input'],
                    bordercolor=c['border'], lightcolor=c['surface'], darkcolor=c['border'],
                    troughcolor=c['bg'], font=FONT)
        s.configure('TFrame', background=c['surface'])
        s.configure('Card.TFrame', background=c['surface'])
        s.configure('TLabel', background=c['surface'], foreground=c['fg'], font=FONT)
        s.configure('Muted.TLabel', background=c['surface'], foreground=c['muted'], font=FONT_SM)
        s.configure('Title.TLabel', background=c['surface'], foreground=c['fg'], font=FONT_B)
        s.configure('TButton', background=c['surface'], foreground=c['fg'], font=FONT,
                    bordercolor=c['border'], padding=(14, 9))
        s.map('TButton', background=[('active', c['accent_soft'])],
              foreground=[('active', c['accent'])], bordercolor=[('active', c['accent'])])
        s.configure('Accent.TButton', background=c['accent'], foreground=c['accent_fg'],
                    font=FONT_B, bordercolor=c['accent'], padding=(16, 10))
        s.map('Accent.TButton', background=[('active', c['accent_hover'])],
              foreground=[('active', c['accent_fg'])])
        s.configure('TCheckbutton', background=c['surface'], foreground=c['fg'], font=FONT,
                    indicatorcolor=c['accent_fg'], indicatorbackground=c['input'],
                    bordercolor=c['border'], borderwidth=0, focuscolor=c['surface'])
        s.map('TCheckbutton',
              background=[('active', c['accent_soft'])],
              foreground=[('active', c['accent'])],
              indicatorbackground=[('selected', c['accent']),
                                   ('active', c['hover'])])
        s.configure('TEntry', fieldbackground=c['input'], foreground=c['fg'], font=FONT,
                    bordercolor=c['input_line'], lightcolor=c['input_line'],
                    darkcolor=c['input_line'], padding=(11, 8), insertcolor=c['fg'])
        s.map('TEntry', bordercolor=[('focus', c['accent'])],
              lightcolor=[('focus', c['accent'])], darkcolor=[('focus', c['accent'])])
        s.configure('Treeview', background=c['tree_bg'], fieldbackground=c['tree_bg'],
                    foreground=c['fg'], bordercolor=c['border'], font=FONT_SM, rowheight=34)
        s.configure('Treeview.Heading', background=c['tree_head'], foreground=c['label'],
                    bordercolor=c['border'], font=FONT_B, padding=(8, 7))
        s.map('Treeview.Heading', background=[('active', c['accent_soft'])])
        s.map('Treeview',
              background=[('selected', '!focus', c['accent_soft']),
                          ('selected', c['accent_soft'])],
              foreground=[('selected', '!focus', c['accent']),
                          ('selected', c['accent'])])
        s.configure('Vertical.TScrollbar', background=c['surface'], bordercolor=c['border'],
                    arrowcolor=c['muted'], troughcolor=c['bg'])
        s.map('Vertical.TScrollbar', background=[('active', c['border_strong'])])
        s.configure('Horizontal.TProgressbar', background=c['accent'], troughcolor=c['border'],
                    bordercolor=c['border'], thickness=12)
        for w in self.themed:
            try:
                w.set_colors(c)
            except Exception:
                pass
        for kind, widgets in self.plain.items():
            for w in widgets:
                w.configure(background=c[kind])
                if w.winfo_class() == 'Label':
                    key = getattr(w, '_color', None) or \
                          ('muted' if getattr(w, '_muted', False) else 'fg')
                    w.configure(foreground=c[key])
        if hasattr(self, 'log'):
            self.log.configure(background=c['log_bg'], foreground=c['log_fg'],
                               insertbackground=c['fg'], highlightbackground=c['border'],
                               highlightcolor=c['accent'])
            self.log.tag_configure('hint', foreground=c['muted'])
            try:
                self.log.vbar.configure(background=c['surface'], troughcolor=c['track'],
                                        activebackground=c['border_strong'],
                                        highlightthickness=0, borderwidth=0,
                                        relief='flat', width=12)
            except Exception:
                pass

    def toggle_theme(self):
        self.apply_theme('dark' if self.theme == 'light' else 'light')
        self.save_theme()

    # ---------- 基础控件工厂 ----------
    def _frame(self, parent, kind='surface'):
        f = tk.Frame(parent)
        self.plain[kind].append(f)
        f.configure(background=self.c[kind])
        return f

    def _label(self, parent, kind='surface', muted=False, color=None, **kw):
        l = tk.Label(parent, **kw)
        l._muted = muted
        l._color = color
        self.plain[kind].append(l)
        l.configure(background=self.c[kind],
                    foreground=self.c[color] if color else
                    (self.c['muted'] if muted else self.c['fg']))
        return l

    # ---------- 构建 ----------
    def _build(self):
        # 顶栏
        head = self._frame(self.root, 'bg')
        head.pack(fill='x', padx=22, pady=(18, 12))
        box = self._frame(head, 'bg')
        box.pack(side='left')
        self._label(box, 'bg', text=APP_NAME, font=FONT_H1).pack(anchor='w')
        self._label(box, 'bg', text=f'v{APP_VERSION} · 离线运行，不上传任何文本',
                    font=FONT_SM, muted=True).pack(anchor='w', pady=(4, 0))
        self.btn_theme = Button(head, self.c, text='深色', command=self.toggle_theme,
                                kind='ghost')
        self.btn_theme.pack(side='right', pady=4)

        # 分段导航
        self.seg = Segmented(self.root, self.c,
                             ['译前处理', '译后处理', '运行日志'],
                             command=self.show_page)
        self.themed.append(self.seg)
        self.seg.pack(fill='x', padx=22, pady=(0, 14))

        # 页面容器（每页可纵向滚动：窗口高度不足时在右侧出现滚动条）
        self.pages = self._frame(self.root, 'bg')
        self.pages.pack(fill='both', expand=True, padx=22)
        self.frames = []
        for build in (self._page_pre, self._page_post, self._page_log):
            f = self._frame(self.pages, 'bg')
            f.place(relx=0, rely=0, relwidth=1, relheight=1)
            self._scrollable(f, build)
            self.frames.append(f)
        self.show_page(0)

        # 底部状态
        foot = self._frame(self.root, 'bg')
        foot.pack(fill='x', padx=22, pady=(12, 16))
        top = self._frame(foot, 'bg')
        top.pack(fill='x')
        self.pill = StatusPill(top, self.c, text='就绪')
        self.themed.append(self.pill)
        self.pill.pack(side='left')
        self.lbl_pct = self._label(top, 'bg', text='0%', font=FONT_SM, muted=True)
        self.lbl_pct.pack(side='right')
        self.bar = ProgressBar(foot, self.c)
        self.bar.pack(fill='x', pady=(9, 0))

    def _scrollable(self, parent, build):
        """在 parent 内放 Canvas + 右侧滚动条；内容超出可视高度时滚动条才出现。"""
        cv = tk.Canvas(parent, highlightthickness=0, borderwidth=0,
                       background=self.c['bg'])
        self.plain['bg'].append(cv)
        vsb = ttk.Scrollbar(parent, orient='vertical', command=cv.yview)
        cv.configure(yscrollcommand=vsb.set)
        cv.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y', padx=(8, 0))   # 与左侧内容留出间距
        inner = self._frame(cv, 'bg')
        win = cv.create_window(0, 0, window=inner, anchor='nw')

        def _sync(event=None):
            cv.update_idletasks()
            bb = cv.bbox('all')
            cv.configure(scrollregion=bb if bb else (0, 0, 1, 1))
            cv.itemconfigure(win, width=cv.winfo_width())
            need = bool(bb) and cv.winfo_height() > 0 and bb[3] > cv.winfo_height() + 4
            packed = vsb.winfo_manager() == 'pack'
            if need and not packed:
                vsb.pack(side='right', fill='y', padx=(8, 0))
            elif not need and packed:
                vsb.pack_forget()

        def _needs_scroll():
            """内容是否真的溢出视口——只有溢出时才允许滚轮滚动。"""
            bb = cv.bbox('all')
            return bool(bb) and cv.winfo_height() > 0 and bb[3] > cv.winfo_height() + 4

        def _on_wheel(event):
            if _needs_scroll():
                cv.yview_scroll(int(-1 * (event.delta / 120)), 'units')
            return 'break'

        def _on_up(_):
            if _needs_scroll():
                cv.yview_scroll(-1, 'units')
            return 'break'

        def _on_down(_):
            if _needs_scroll():
                cv.yview_scroll(1, 'units')
            return 'break'

        # Tk 的父控件不会收到子控件的滚轮事件，因此给每个内容控件单独绑定
        # （跳过自身带滚动条的 Treeview / 日志文本框，让它们各自滚动）。
        def _bind_wheel(w):
            if w.winfo_class() in ('Treeview', 'Text'):
                return
            w.bind('<MouseWheel>', _on_wheel)
            w.bind('<Button-4>', _on_up)
            w.bind('<Button-5>', _on_down)
            for ch in w.winfo_children():
                _bind_wheel(ch)

        inner.bind('<Configure>', _sync)
        cv.bind('<Configure>', _sync)
        build(inner)
        _bind_wheel(inner)
        cv.after(80, _sync)       # 布局完成后校准一次

    def _page_head(self, parent, title, desc):
        f = self._frame(parent, 'bg')
        f.pack(fill='x', pady=(0, 10))
        self._label(f, 'bg', text=title, font=FONT_H2).pack(anchor='w')
        self._label(f, 'bg', text=desc, font=FONT_SM, muted=True).pack(anchor='w', pady=(5, 0))

    def _card(self, parent, title, hint='', expand=False):
        card = Card(parent, self.c, pad=16, expand=expand)
        self.themed.append(card)
        card.pack(fill='both' if expand else 'x', expand=expand, pady=(0, 12))
        if title:
            head = self._frame(card.inner, 'surface')
            head.pack(fill='x', pady=(0, 12))
            self._label(head, 'surface', text=title, font=FONT_H3).pack(side='left')
            if hint:
                self._label(head, 'surface', text=hint, font=FONT_SM,
                            muted=True).pack(side='right')
        return card

    def _field(self, parent, label, var, cmd):
        box = self._frame(parent, 'surface')
        box.pack(fill='x', pady=(0, 12))
        self._label(box, 'surface', text=label, font=FONT_SM,
                    color='label').pack(anchor='w', pady=(0, 6))
        row = self._frame(box, 'surface')
        row.pack(fill='x')
        ttk.Entry(row, textvariable=var, font=FONT).pack(side='left', fill='x',
                                                         expand=True, padx=(0, 10))
        b = Button(row, self.c, text='浏览', command=cmd, kind='ghost')
        b.pack(side='left')

    def _checks(self, parent, items, vertical=False):
        box = self._frame(parent, 'surface')
        box.pack(fill='x', pady=2)
        for text, var in items:
            ck = ttk.Checkbutton(box, text=text, variable=var, style='TCheckbutton')
            if vertical:
                ck.pack(anchor='w', pady=3)
            else:
                ck.pack(side='left', padx=(0, 26))
        return box

    def _actions(self, parent, specs):
        """specs: [(文字, 回调, kind), ...]；按钮建在行内，行贴在内容下方。"""
        row = self._frame(parent, 'bg')
        row.pack(fill='x', pady=(6, 2))
        out = []
        for text, cmd, kind in specs:
            b = Button(row, self.c, text=text, command=cmd, kind=kind)
            b.pack(side='left', padx=(0, 10))
            out.append(b)
        return out

    # ---- 页1：译前处理 ----
    def _page_pre(self, p):
        self._page_head(p, '译前处理', '将原文 txt 导出为待翻译的 xlsx，原文统一放在 A 列。')
        c1 = self._card(p, '选择文件')
        self._field(c1.inner, '原文 txt 文件夹', self.v_src, self.pick_src)
        self._field(c1.inner, '输出文件夹', self.v_out, self.pick_out)
        self._checks(c1.inner, [('包含子文件夹', self.v_recursive)])

        c2 = self._card(p, '清洗规则', hint='三项可按需单独关闭')
        self._checks(c2.inner, [('去除全角空格（保留半角空格）', self.v_ideo),
                                ('去除空行和纯空白行', self.v_blank),
                                ('去除重复行（保留首次出现）', self.v_dup)])

        self.btn_run1, self.btn_stop1, self.btn_open1 = self._actions(p, [
            ('开始处理', self.start_import, 'primary'),
            ('停止', self.stop, 'ghost'),
            ('打开输出文件夹', lambda: self.open_dir(self.v_out.get()), 'ghost'),
        ])
        self.btn_stop1.set_state(False)

    # ---- 页2：译后处理 ----
    def _page_post(self, p):
        self._page_head(p, '译后处理',
                        '按替换规则将译文 xlsx 导出为 sqlite 预翻译文件。')
        c1 = self._card(p, '选择文件')
        self._field(c1.inner, '译文 xlsx 文件夹', self.v_tsrc, self.pick_tsrc)
        self._field(c1.inner, '输出文件夹', self.v_tout, self.pick_tout)
        self._checks(c1.inner, [('包含子文件夹', self.v_trecursive)])

        c2 = self._card(p, '替换规则', hint='顺序敏感，双击行可编辑')
        wrap = self._frame(c2.inner, 'surface')
        wrap.pack(fill='x')
        self.tree = ttk.Treeview(wrap, columns=('on', 'find', 'repl', 're'),
                                 show='headings', height=5)
        for col, w, t in (('on', 66, '启用'), ('find', 240, '查找'),
                          ('repl', 240, '替换为'), ('re', 70, '正则')):
            self.tree.heading(col, text=t)
            self.tree.column(col, width=w, anchor='center' if col in ('on', 're') else 'w')
        vs = ttk.Scrollbar(wrap, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscroll=vs.set)
        self.tree.pack(side='left')
        vs.pack(side='left', fill='y', padx=(8, 0))
        self.tree.bind('<Double-1>', lambda e: self.edit_rule())
        self.refresh_rules()

        side = self._frame(wrap, 'surface')
        side.pack(side='left', fill='y', padx=(16, 0))
        side_items = (('添加', self.add_rule), ('删除', self.del_rule),
                      ('编辑', self.edit_rule), ('恢复默认', self.reset_rules),
                      ('上移', lambda: self.move_rule(-1)), ('下移', lambda: self.move_rule(1)),
                      ('导入', self.import_rules), ('导出', self.export_rules))
        for idx, (t, cmd) in enumerate(side_items):
            b = Button(side, self.c, text=t, command=cmd, kind='ghost', width=12)
            b.grid(row=idx // 2, column=idx % 2, padx=4, pady=4, sticky='ew')
        side.grid_columnconfigure(0, weight=1)
        side.grid_columnconfigure(1, weight=1)

        self.btn_run2, self.btn_stop2, self.btn_open2 = self._actions(p, [
            ('开始处理', self.start_trans, 'primary'),
            ('停止', self.stop, 'ghost'),
            ('打开输出文件夹', lambda: self.open_dir(self.v_tout.get()), 'ghost'),
        ])
        self.btn_stop2.set_state(False)

    # ---- 页3：日志 ----
    def _page_log(self, p):
        self._page_head(p, '运行日志', '处理进度与异常信息会实时显示在这里。')
        c = self._card(p, '', expand=True)
        self.log = scrolledtext.ScrolledText(c.inner, height=14, wrap='word', relief='flat',
                                             borderwidth=0, highlightthickness=1,
                                             font=FONT_LOG, spacing1=3, spacing3=3,
                                             padx=12, pady=10)
        self.log.pack(fill='both', expand=True)
        self.log.tag_configure('hint', foreground=self.c['muted'])
        self.log.insert('1.0', '暂无日志。开始处理后，详细信息会实时显示在这里。', 'hint')
        self.log.configure(state='disabled')
        self._actions(p, [('清空日志', self.clear_log, 'ghost')])

    # ---------- 状态同步 ----------
    def _sync_status(self, *_):
        t = self.v_status.get()
        if t.startswith('完成'):
            tone = 'ok'
        elif t.startswith('出错') or t.startswith('未'):
            tone = 'err'
        elif t.startswith('已停止'):
            tone = 'idle'
        elif t == '就绪':
            tone = 'idle'
        else:
            tone = 'run'
        if hasattr(self, 'pill'):
            self.pill.set(t, tone)

    def _sync_progress(self, *_):
        if hasattr(self, 'bar'):
            self.bar.set(self.v_progress.get())
        if hasattr(self, 'lbl_pct'):
            self.lbl_pct.configure(text='%d%%' % int(self.v_progress.get() + 0.5))

    def show_page(self, i):
        self.seg.select(i)
        self.frames[i].tkraise()

    # ---------- 规则 ----------
    def refresh_rules(self):
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(self.rules):
            self.tree.insert('', 'end',
                             values=('启用' if r.get('on') else '停用',
                                     r.get('find', ''),
                                     r.get('repl', ''),
                                     '是' if r.get('re') else '否'))

    def _sel(self):
        s = self.tree.selection()
        return self.tree.index(s[0]) if s else None

    def add_rule(self):
        d = RuleDialog(self.root, self.c, '添加规则')
        if d.result:
            i = self._sel()
            self.rules.insert(i + 1 if i is not None else len(self.rules), d.result)
            self.refresh_rules()

    def edit_rule(self):
        i = self._sel()
        if i is None:
            messagebox.showinfo(APP_NAME, '请先选中一行规则。')
            return
        d = RuleDialog(self.root, self.c, '编辑规则', self.rules[i])
        if d.result:
            self.rules[i] = d.result
            self.refresh_rules()

    def del_rule(self):
        i = self._sel()
        if i is None:
            return
        del self.rules[i]
        self.refresh_rules()

    def move_rule(self, delta):
        i = self._sel()
        if i is None:
            return
        j = i + delta
        if 0 <= j < len(self.rules):
            self.rules[i], self.rules[j] = self.rules[j], self.rules[i]
            self.refresh_rules()
            self.tree.selection_set(self.tree.get_children()[j])

    def reset_rules(self):
        if messagebox.askyesno(APP_NAME, '恢复为内置默认规则？当前规则将被覆盖。'):
            self.rules = [dict(r) for r in DEFAULT_RULES]
            self.refresh_rules()

    def export_rules(self):
        f = filedialog.asksaveasfilename(defaultextension='.json', filetypes=[('JSON', '*.json')])
        if f:
            json.dump(self.rules, open(f, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

    def import_rules(self):
        f = filedialog.askopenfilename(filetypes=[('JSON', '*.json')])
        if f:
            self.rules = json.load(open(f, encoding='utf-8'))
            self.refresh_rules()

    # ---------- 路径 ----------
    def pick_src(self):
        d = filedialog.askdirectory(title='选择原文本文件夹')
        if d:
            d = os.path.normpath(d)
            self.v_src.set(d)
            if not self.v_out.get():
                self.v_out.set(os.path.normpath(os.path.join(d, '待翻译')))

    def pick_out(self):
        d = filedialog.askdirectory(title='选择输出文件夹')
        if d:
            self.v_out.set(os.path.normpath(d))

    def pick_tsrc(self):
        d = filedialog.askdirectory(title='选择译文 xlsx 所在文件夹')
        if d:
            d = os.path.normpath(d)
            self.v_tsrc.set(d)
            if not self.v_tout.get():
                self.v_tout.set(os.path.normpath(os.path.join(d, 'sqlite预翻译')))

    def pick_tout(self):
        d = filedialog.askdirectory(title='选择输出文件夹')
        if d:
            self.v_tout.set(os.path.normpath(d))

    def open_dir(self, d):
        if d and os.path.isdir(d):
            try:
                os.startfile(d)
            except Exception as e:
                messagebox.showwarning(APP_NAME, f'打开失败：{e}')
        else:
            messagebox.showinfo(APP_NAME, '目录还不存在，请先处理一次。')

    # ---------- 日志 / 进度 ----------
    def clear_log(self):
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')

    def append_log(self, msg):
        self.log.configure(state='normal')
        if 'hint' in self.log.tag_names('1.0'):
            self.log.delete('1.0', 'end')
        self.log.insert('end', msg + '\n')
        self.log.see('end')
        self.log.configure(state='disabled')

    def _log_sep(self, title):
        """在新一轮处理前插入分隔标记，保留之前的日志。"""
        self.log.configure(state='normal')
        if 'hint' in self.log.tag_names('1.0'):
            self.log.delete('1.0', 'end')
        stamp = time.strftime('%H:%M:%S')
        self.log.insert('end', '\n' + ('─' * 18) + f' {title} {stamp} ' + ('─' * 18) + '\n')
        self.log.see('end')
        self.log.configure(state='disabled')

    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == 'log':
                    self.append_log(payload)
                elif kind == 'status':
                    self.v_status.set(payload)
                elif kind == 'progress':
                    self.v_progress.set(payload)
                elif kind == 'done':
                    self.btn_run1.set_state(True)
                    self.btn_run2.set_state(True)
                    self.btn_stop1.set_state(False)
                    self.btn_stop2.set_state(False)
                    self.v_status.set(payload)
                    if payload.startswith('完成'):
                        messagebox.showinfo(APP_NAME, payload)
                    else:
                        messagebox.showwarning(APP_NAME, payload)
                        self.show_page(2)
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def stop(self):
        self.stop_flag.set()
        self.q.put(('status', '正在停止…'))

    def _launch(self, run_btn, stop_btn, target, args, title=''):
        self.stop_flag.clear()
        run_btn.set_state(False)
        stop_btn.set_state(True)
        self.v_progress.set(0)          # 重置进度，避免上轮残留值影响本轮百分比
        self.v_status.set('正在处理…')
        self._log_sep(title or '处理')  # 保留历史日志，仅插入分隔
        threading.Thread(target=target, args=args, daemon=True).start()

    # ---------- 执行 ----------
    def start_import(self):
        src, out = self.v_src.get().strip(), self.v_out.get().strip()
        if not src or not os.path.isdir(src):
            messagebox.showwarning(APP_NAME, '请先选择有效的原文本文件夹。')
            return
        if not out:
            messagebox.showwarning(APP_NAME, '请先选择输出文件夹。')
            return
        cfg = {'out_dir': out, 'drop_ideo': self.v_ideo.get(),
               'drop_blank': self.v_blank.get(), 'drop_dup': self.v_dup.get()}
        self._launch(self.btn_run1, self.btn_stop1, self._run_import,
                     (src, cfg, self.v_recursive.get()), '译前处理')

    def _run_import(self, src, cfg, recursive):
        def log(m):
            self.q.put(('log', m))
        try:
            items = collect(src, recursive, ('.txt',))
            if not items:
                self.q.put(('done', '未在源文件夹找到 .txt 文件'))
                return
            log(f'译前处理：共 {len(items)} 个文件')
            total = 0
            for i, (full, rel) in enumerate(items, 1):
                if self.stop_flag.is_set():
                    self.q.put(('done', f'已停止（完成 {i-1}/{len(items)}）'))
                    return
                self.q.put(('status', f'处理 {i}/{len(items)}'))
                self.q.put(('progress', (i - 1) * 100.0 / len(items)))
                total += process_import(full, rel, cfg, log)['final']
                self.q.put(('progress', i * 100.0 / len(items)))
            self.q.put(('done', f'完成：{len(items)} 个 xlsx，共 {total} 行'))
        except Exception:
            log(traceback.format_exc())
            self.q.put(('done', '出错：详见日志'))

    def start_trans(self):
        src, out = self.v_tsrc.get().strip(), self.v_tout.get().strip()
        if not src or not os.path.isdir(src):
            messagebox.showwarning(APP_NAME, '请先选择译文 xlsx 所在文件夹。')
            return
        if not out:
            messagebox.showwarning(APP_NAME, '请先选择输出文件夹。')
            return
        self.save_rules()
        cfg = {'out_dir': out, 'rules': self.rules}
        self._launch(self.btn_run2, self.btn_stop2, self._run_trans,
                     (src, cfg, self.v_trecursive.get()), '译后处理')

    def _run_trans(self, src, cfg, recursive):
        def log(m):
            self.q.put(('log', m))
        try:
            items = collect(src, recursive, ('.xlsx',))
            if not items:
                self.q.put(('done', '未找到 .xlsx 文件'))
                return
            log(f'译后处理：共 {len(items)} 个文件；'
                f'替换规则 {sum(1 for r in cfg["rules"] if r.get("on"))} 条生效')
            rows = filled = 0
            for i, (full, rel) in enumerate(items, 1):
                if self.stop_flag.is_set():
                    self.q.put(('done', f'已停止（完成 {i-1}/{len(items)}）'))
                    return
                self.q.put(('status', f'处理 {i}/{len(items)}'))
                self.q.put(('progress', (i - 1) * 100.0 / len(items)))
                st = process_translation(full, rel, cfg, log)
                rows += st['rows']
                filled += st['filled']
                self.q.put(('progress', i * 100.0 / len(items)))
            self.q.put(('done', f'完成：{len(items)} 个 sqlite，{rows} 行，回填空译文 {filled} 处'))
        except Exception:
            log(traceback.format_exc())
            self.q.put(('done', '出错：详见日志'))


def main():
    if '--selftest' in sys.argv:
        i = sys.argv.index('--selftest')
        return run_cli('import', sys.argv[i + 1], sys.argv[i + 2])
    if '--selftest-trans' in sys.argv:
        i = sys.argv.index('--selftest-trans')
        return run_cli('trans', sys.argv[i + 1], sys.argv[i + 2],
                       sys.argv[i + 3] if len(sys.argv) > i + 3 else None)
    enable_dpi()
    root = tk.Tk()
    system_scale(root)
    try:
        icon = resource_path(os.path.join('resources', 'feather.ico'))
        if os.path.isfile(icon):
            root.iconbitmap(icon)
    except Exception:
        pass
    App(root)
    root.mainloop()
    return 0


if __name__ == '__main__':
    sys.exit(main())
