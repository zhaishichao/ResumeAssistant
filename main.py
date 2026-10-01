# -*- coding: utf-8 -*-
"""简历助手：解析个人简历 Word 文档，以大纲树 + 内容面板展示，支持点击复制与置顶。"""
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document

BG = '#f5f6f8'
CARD = '#ffffff'
ACCENT = '#3b82f6'
FG = '#1f2937'
MUTED = '#9ca3af'
FONT = ('Microsoft YaHei UI', 10)
FONT_BOLD = ('Microsoft YaHei UI', 10, 'bold')
FONT_TITLE = ('Microsoft YaHei UI', 11, 'bold')
FONT_SMALL = ('Microsoft YaHei UI', 9)
BALL_HOTKEY = '<F8>'
BALL_KEY = '#ff00ff'

LABEL_RE = re.compile(r'^([一-龥A-Za-z]{1,8})：(.*)$')
NUM_RE = re.compile(r'^\d+、')
DATE_RE = re.compile(r'20\d{2}[.\-/年]')
FIELD_SEP_RE = re.compile(r'([一-龥A-Za-z]{1,8})：')
DATE_TAIL_RE = re.compile(r'(20\d{2}\.\d{2}(?:\.\d{2})?(?:—20\d{2}\.\d{2}(?:\.\d{2})?)?)\s*$')
DATE_PAREN_RE = re.compile(r'（(20\d{2}[^）]*)）\s*$')


def heading_level(p):
    n = (p.style.name or '').lower()
    if n in ('heading 1', '标题 1'):
        return 1
    if n in ('heading 2', '标题 2'):
        return 2
    return 0


def is_noise(t):
    if not t or not t.strip():
        return True
    t = t.strip()
    if t in ('目录', '个人基本信息'):
        return True
    if 'Ctrl+A' in t or '更新目录' in t:
        return True
    return False


def detect_kind(title):
    if '论文' in title:
        return 'paper'
    if '专利' in title:
        return 'patent'
    if '实习' in title:
        return 'intern'
    if '项目' in title:
        return 'project'
    if '学生工作' in title:
        return 'work'
    if '社会实践' in title:
        return 'practice'
    return ''


def split_name(text, kind):
    """把名称行拆成 公司/岗位/名称/职位/时间 等字段，便于分别复制。"""
    fields = []
    time_val = ''
    head = text
    m = DATE_PAREN_RE.search(head)
    if m:
        time_val = m.group(1)
        head = head[:m.start()].rstrip()
    else:
        m = DATE_TAIL_RE.search(head)
        if m:
            time_val = m.group(1)
            head = head[:m.start()].rstrip()
    if kind in ('intern', 'project') and '—' in head:
        left, right = head.split('—', 1)
        left, right = left.strip(), right.strip()
        if left:
            fields.append(('公司' if kind == 'intern' else '名称', left))
        if right:
            fields.append(('岗位', right))
    elif kind == 'work' and ' ' in head:
        left, right = head.rsplit(' ', 1)
        if left.strip():
            fields.append(('名称', left.strip()))
        if right.strip():
            fields.append(('职位', right.strip()))
    elif head:
        fields.append(('名称', head))
    if time_val:
        fields.append(('时间', time_val))
    return fields


def split_paper(text):
    """论文：拆成 作者/题目/期刊/补充信息。"""
    t = re.sub(r'^\d+、', '', text).strip()
    fields = []
    extra = ''
    m = re.search(r'（([^）]*)）\s*$', t)
    if m:
        extra = m.group(1)
        t = t[:m.start()].rstrip(' 。.,;')
    year = ''
    m = re.search(r',\s*(20\d{2})\s*\.?\s*$', t)
    if m:
        year = m.group(1)
        t = t[:m.start()].rstrip()
    title = ''
    authors = ''
    journal = ''
    m = re.search(r'"([^"]*)"', t)
    if m:
        title = m.group(1).strip().rstrip(',')
        authors = t[:m.start()].rstrip(' ,')
        after = t[m.end():].strip()
        jm = re.match(r'in\s+(.+)$', after)
        journal = jm.group(1).rstrip(' ,.') if jm else after.rstrip(' ,.')
    else:
        authors = t.rstrip(' ,')
    if authors:
        fields.append(('作者', authors))
    if title:
        fields.append(('题目', title))
    if journal:
        fields.append(('期刊', journal + (', ' + year if year else '')))
    if extra:
        fields.append(('补充信息', extra))
    return fields


def split_patent(text):
    """专利：拆成 题目/专利号/补充信息。"""
    t = re.sub(r'^\d+、', '', text).strip()
    fields = []
    extra = ''
    m = re.search(r'（([^）]*)）\s*$', t)
    if m:
        extra = m.group(1)
        t = t[:m.start()].rstrip(' ，。.,;')
    title = ''
    rest = ''
    m = re.search(r'^(.*?)：(.*)$', t)
    if m:
        title = m.group(1).strip()
        rest = m.group(2).strip()
    else:
        title = t
    number = ''
    tail = ''
    nm = re.match(r'(\S+?\[P\])(?:[，,]\s*(.*))?$', rest)
    if nm:
        number = nm.group(1)
        tail = (nm.group(2) or '').strip()
    else:
        parts = re.split(r'[，,]', rest, 1)
        number = parts[0].strip()
        tail = parts[1].strip() if len(parts) > 1 else ''
    extra_all = (tail + ('（' + extra + '）' if extra else '')).strip()
    if title:
        fields.append(('题目', title))
    if number:
        fields.append(('专利号', number))
    if extra_all:
        fields.append(('补充信息', extra_all))
    return fields


def blocks_of(text, context=''):
    """把一段正文拆成若干展示块：(kind, tag, text)。
    kind: name 名称/字段 field/内容 content/空标签 subhead。"""
    m = LABEL_RE.match(text)
    if m:
        label, val = m.group(1), m.group(2).strip()
        if not val:
            return [('subhead', label, '')]
        segs = FIELD_SEP_RE.split(val)
        if len(segs) > 1:
            blocks = [('field', label, segs[0].strip())]
            for i in range(1, len(segs), 2):
                lab, v = segs[i], (segs[i + 1].strip() if i + 1 < len(segs) else '')
                if v:
                    blocks.append(('field', lab, v))
            return blocks
        return [('field', label, val)]
    if NUM_RE.match(text):
        if context == 'paper':
            return [('field', lab, val) for lab, val in split_paper(text)]
        if context == 'patent':
            return [('field', lab, val) for lab, val in split_patent(text)]
        return [('content', '内容', text)]
    if DATE_RE.search(text) or ('—' in text and len(text) <= 40):
        if context in ('intern', 'project', 'work', 'practice'):
            return [('field', lab, val) for lab, val in split_name(text, context)]
        return [('name', '名称', text)]
    return [('content', '内容', text)]


def parse_docx(path):
    """返回节点列表，每个节点 {'title', 'children', 'body'}。"""
    doc = Document(path)
    sections = []
    cur1 = cur2 = None
    pre = None

    def prenode():
        nonlocal pre
        if pre is None:
            pre = {'title': '个人基本信息', 'children': [], 'body': []}
            sections.insert(0, pre)
        return pre

    for p in doc.paragraphs:
        t = p.text.strip()
        lvl = heading_level(p)
        if lvl == 1:
            cur1 = {'title': t, 'children': [], 'body': []}
            sections.append(cur1)
            cur2 = None
        elif lvl == 2:
            node = {'title': t, 'children': [], 'body': []}
            if cur1 is None:
                sections.append(node)
                cur1 = node
            else:
                cur1['children'].append(node)
            cur2 = node
        else:
            if is_noise(t):
                continue
            if cur2 is not None:
                cur2['body'].append(t)
            elif cur1 is not None:
                cur1['body'].append(t)
            else:
                prenode()['body'].append(t)
    return sections


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title('简历助手')
        self.root.geometry('920x640')
        self.root.minsize(280, 240)
        self.root.configure(bg=BG)
        self.topmost = False
        self.ball_mode = False
        self.sections = []
        self.idmap = {}
        self._wrap_labels = []
        self._build_ui()
        self._welcome()
        self.root.after(200, self.open_dialog)

    def _build_ui(self):
        style = ttk.Style()
        style.configure('Treeview', font=FONT, rowheight=30)
        style.configure('Treeview', background='#ffffff', fieldbackground='#ffffff')

        bar = tk.Frame(self.root, bg='#ffffff', height=50)
        bar.pack(fill='x')
        bar.pack_propagate(False)
        ttk.Button(bar, text='打开文档', command=self.open_dialog).pack(side='left', padx=10, pady=8)
        self.top_btn = ttk.Button(bar, text='置顶', command=self.toggle_top)
        self.top_btn.pack(side='left', padx=4)
        self.ball_btn = ttk.Button(bar, text='悬浮球', command=self.toggle_ball)
        self.ball_btn.pack(side='left', padx=4)
        self.path_lbl = tk.Label(bar, text='', bg='#ffffff', fg=MUTED, font=FONT_SMALL, anchor='w')
        self.path_lbl.pack(side='left', padx=14, fill='x', expand=True)

        paned = ttk.PanedWindow(self.root, orient='horizontal')
        paned.pack(fill='both', expand=True)

        left = tk.Frame(paned, bg='#ffffff')
        tk.Label(left, text='大纲', bg='#ffffff', fg=FG, font=FONT_BOLD, anchor='w').pack(
            fill='x', padx=14, pady=(12, 4))
        self.tree = ttk.Treeview(left, show='tree', selectmode='browse')
        self.tree.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        self.tree.column('#0', width=250, minwidth=80)
        self.tree.bind('<<TreeviewSelect>>', self.on_select)

        right = tk.Frame(paned, bg=BG)
        self.breadcrumb = tk.Label(right, text='', bg=BG, fg=MUTED, font=FONT_SMALL, anchor='w')
        self.breadcrumb.pack(fill='x', padx=18, pady=(14, 2))

        self.canvas = tk.Canvas(right, bg=BG, highlightthickness=0)
        vsb = ttk.Scrollbar(right, orient='vertical', command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        self.canvas.pack(side='left', fill='both', expand=True)
        self.inner = tk.Frame(self.canvas, bg=BG)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor='nw')
        self.inner.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.canvas.bind('<Configure>', self._on_resize)
        self.canvas.bind('<MouseWheel>', self._wheel)

        paned.add(left, weight=0)
        paned.add(right, weight=1)

        self.status = tk.Label(self.root, text='', bg='#eef1f5', fg=MUTED, font=FONT_SMALL, anchor='w')
        self.status.pack(fill='x', side='bottom')

        self._create_ball()
        self.root.bind(BALL_HOTKEY, lambda e: self.toggle_ball())

    def _welcome(self):
        tk.Label(self.inner, text='请点击左上角【打开文档】，选择个人简历 Word 文档',
                 bg=BG, fg=MUTED, font=FONT).pack(pady=40)

    def _wheel(self, e):
        self.canvas.yview_scroll(int(-e.delta / 120), 'units')

    def _bind_wheel_recursive(self, w):
        w.bind('<MouseWheel>', self._wheel)
        for c in w.winfo_children():
            self._bind_wheel_recursive(c)

    def _on_resize(self, e):
        self.canvas.itemconfig(self._win, width=e.width)
        w = max(60, e.width - 140)
        for l in self._wrap_labels:
            try:
                l.configure(wraplength=w)
            except Exception:
                pass

    def open_dialog(self):
        path = filedialog.askopenfilename(
            title='选择简历 Word 文档',
            filetypes=[('Word 文档', '*.docx'), ('所有文件', '*.*')])
        if not path:
            return
        try:
            self.load(path)
        except Exception as ex:
            messagebox.showerror('打开失败', str(ex))

    def load(self, path):
        self.sections = parse_docx(path)
        self.path_lbl.config(text=path.replace('/', '\\').split('\\')[-1])
        self.refresh_tree()

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        self.idmap = {}
        for sec in self.sections:
            iid = self.tree.insert('', 'end', text=sec['title'], open=True)
            self.idmap[iid] = sec
            for c in sec['children']:
                cid = self.tree.insert(iid, 'end', text=c['title'])
                self.idmap[cid] = c
        if self.sections:
            first = self.tree.get_children()[0]
            self.tree.selection_set(first)

    def on_select(self, e=None):
        sel = self.tree.selection()
        if not sel:
            return
        node = self.idmap.get(sel[0])
        if node is None:
            return
        crumbs = [node['title']]
        for sec in self.sections:
            if node in sec['children']:
                crumbs.insert(0, sec['title'])
                break
        self.breadcrumb.config(text=' / '.join(crumbs))
        self.render(node)

    def render(self, node):
        for w in self.inner.winfo_children():
            w.destroy()
        self._wrap_labels = []
        ctx = detect_kind(node['title'])
        if node['children']:
            for b in node['body']:
                self._expand(b, ctx)
            for c in node['children']:
                self._subtitle(c['title'])
                cctx = detect_kind(c['title'])
                for b in c['body']:
                    self._expand(b, cctx)
        else:
            for b in node['body']:
                self._expand(b, ctx)
        if not self.inner.winfo_children():
            tk.Label(self.inner, text='（本节暂无内容）', bg=BG, fg=MUTED, font=FONT).pack(pady=40)
        else:
            self._apply_wrap()
        self._bind_wheel_recursive(self.inner)
        self.canvas.yview_moveto(0)

    def _expand(self, text, context=''):
        for kind, tag, txt in blocks_of(text, context):
            if kind == 'subhead':
                self._subhead(tag)
            else:
                self._card(kind, tag, txt)

    def _apply_wrap(self):
        w = max(60, self.canvas.winfo_width() - 140)
        for l in self._wrap_labels:
            l.configure(wraplength=w)

    def _subtitle(self, text):
        row = tk.Frame(self.inner, bg=BG)
        row.pack(fill='x', padx=16, pady=(12, 2))
        tk.Label(row, text='▎', bg=BG, fg=ACCENT, font=FONT_TITLE).pack(side='left')
        tk.Label(row, text=text, bg=BG, fg=ACCENT, font=FONT_TITLE, anchor='w').pack(side='left')

    def _subhead(self, tag):
        row = tk.Frame(self.inner, bg=BG)
        row.pack(fill='x', padx=16, pady=(8, 0))
        tk.Label(row, text='· ' + tag, bg=BG, fg=FG, font=FONT_BOLD, anchor='w').pack(side='left')

    def _card(self, kind, tag, text):
        card = tk.Frame(self.inner, bg=CARD, highlightthickness=1, highlightbackground='#e5e7eb')
        card.pack(fill='x', padx=16, pady=5)
        row = tk.Frame(card, bg=CARD)
        row.pack(fill='x', padx=12, pady=6)
        tk.Label(row, text=tag, bg='#eef2ff', fg=ACCENT, font=FONT_SMALL,
                 padx=6, pady=1).pack(side='left', anchor='n')
        bold = kind == 'name'
        lbl = tk.Label(row, text=text, bg=CARD, fg=FG,
                       font=FONT_BOLD if bold else FONT, anchor='w', justify='left',
                       wraplength=600, cursor='hand2')
        lbl.pack(side='left', fill='x', expand=True, padx=(8, 0))
        lbl.bind('<Button-1>', lambda e, t=text: self.copy(t))
        self._wrap_labels.append(lbl)

    def copy(self, text):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self._show_toast('已复制')
            self._flash('已复制：' + (text[:40] + '…' if len(text) > 40 else text))
        except Exception as ex:
            self._flash('复制失败：' + str(ex))

    def _show_toast(self, msg):
        old = getattr(self, '_toast', None)
        if old is not None:
            try:
                old.destroy()
            except Exception:
                pass
        try:
            x = self.root.winfo_pointerx() - self.root.winfo_rootx() + 12
            y = self.root.winfo_pointery() - self.root.winfo_rooty() + 14
        except Exception:
            x = self.root.winfo_width() // 2
            y = self.root.winfo_height() // 2
        self._toast = tk.Label(self.root, text=msg, bg='#2f3542', fg='#ffffff',
                               font=('Microsoft YaHei UI', 9), padx=12, pady=5)
        self._toast.place(x=x, y=y)
        self._toast.lift()
        self.root.after(1200, self._toast.destroy)

    def _flash(self, msg):
        self.status.config(text=msg, fg=ACCENT)
        self.root.after(2200, lambda: self.status.config(text='', fg=MUTED))

    def toggle_top(self):
        self.topmost = not self.topmost
        self.root.attributes('-topmost', self.topmost)
        self.ball.attributes('-topmost', self.topmost)
        self.top_btn.config(text='置顶 ✓' if self.topmost else '置顶')

    def toggle_ball(self, enable=None):
        if enable is None:
            enable = not self.ball_mode
        self.ball_mode = enable
        if enable:
            self.root.withdraw()
            self.ball.attributes('-topmost', self.topmost)
            self.ball.deiconify()
            self.ball_btn.config(text='恢复窗口')
        else:
            self.ball.withdraw()
            self.root.deiconify()
            self.ball_btn.config(text='悬浮球')

    def _create_ball(self):
        self.ball = tk.Toplevel(self.root)
        self.ball.overrideredirect(True)
        sw = self.root.winfo_screenwidth()
        self.ball.geometry('56x56+%d+%d' % (sw - 96, 96))
        self.ball.attributes('-topmost', self.topmost)
        try:
            self.ball.wm_attributes('-transparentcolor', BALL_KEY)
        except Exception:
            pass
        c = tk.Canvas(self.ball, width=56, height=56, highlightthickness=0, bg=BALL_KEY)
        c.pack()
        c.create_oval(2, 2, 54, 54, fill=ACCENT, outline='')
        c.create_text(28, 28, text='简历', fill='#ffffff', font=('Microsoft YaHei UI', 11, 'bold'))
        c.bind('<Button-1>', self._ball_press)
        c.bind('<B1-Motion>', self._ball_drag)
        c.bind('<Double-Button-1>', lambda e: self.toggle_ball(False))
        c.bind(BALL_HOTKEY, lambda e: self.toggle_ball())
        self._ball_menu = tk.Menu(self.ball, tearoff=0)
        self._ball_menu.add_command(label='恢复窗口', command=lambda: self.toggle_ball(False))
        self._ball_menu.add_command(label='退出', command=self.root.destroy)
        c.bind('<Button-3>', lambda e: self._ball_menu.tk_popup(e.x_root, e.y_root))
        self.ball.withdraw()

    def _ball_press(self, e):
        self._ball_ox = e.x
        self._ball_oy = e.y

    def _ball_drag(self, e):
        x = self.ball.winfo_x() + (e.x - self._ball_ox)
        y = self.ball.winfo_y() + (e.y - self._ball_oy)
        self.ball.geometry('+%d+%d' % (x, y))

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    App().run()
