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

LABEL_RE = re.compile(r'^([一-龥A-Za-z]{1,8})：(.*)$')
NUM_RE = re.compile(r'^\d+、')
DATE_RE = re.compile(r'20\d{2}[.\-/年]')
FIELD_SEP_RE = re.compile(r'([一-龥A-Za-z]{1,8})：')


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


def blocks_of(text):
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
        return [('content', '内容', text)]
    if DATE_RE.search(text) or ('—' in text and len(text) <= 40):
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
        self.root.geometry('1040x700')
        self.root.minsize(760, 520)
        self.root.configure(bg=BG)
        self.topmost = False
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
        self.path_lbl = tk.Label(bar, text='', bg='#ffffff', fg=MUTED, font=FONT_SMALL, anchor='w')
        self.path_lbl.pack(side='left', padx=14, fill='x', expand=True)

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill='both', expand=True)

        left = tk.Frame(body, bg='#ffffff', width=270)
        left.pack(side='left', fill='y')
        left.pack_propagate(False)
        tk.Label(left, text='大纲', bg='#ffffff', fg=FG, font=FONT_BOLD, anchor='w').pack(
            fill='x', padx=14, pady=(12, 4))
        self.tree = ttk.Treeview(left, show='tree', selectmode='browse')
        self.tree.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        self.tree.bind('<<TreeviewSelect>>', self.on_select)

        right = tk.Frame(body, bg=BG)
        right.pack(side='left', fill='both', expand=True)
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

        self.status = tk.Label(self.root, text='', bg='#eef1f5', fg=MUTED, font=FONT_SMALL, anchor='w')
        self.status.pack(fill='x', side='bottom')

    def _welcome(self):
        tk.Label(self.inner, text='请点击左上角【打开文档】，选择个人简历 Word 文档',
                 bg=BG, fg=MUTED, font=FONT).pack(pady=40)

    def _wheel(self, e):
        self.canvas.yview_scroll(int(-e.delta / 120), 'units')

    def _on_resize(self, e):
        self.canvas.itemconfig(self._win, width=e.width)
        w = max(220, e.width - 140)
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
        if node['children']:
            for b in node['body']:
                self._expand(b)
            for c in node['children']:
                self._subtitle(c['title'])
                for b in c['body']:
                    self._expand(b)
        else:
            for b in node['body']:
                self._expand(b)
        if not self.inner.winfo_children():
            tk.Label(self.inner, text='（本节暂无内容）', bg=BG, fg=MUTED, font=FONT).pack(pady=40)
            return
        self._apply_wrap()

    def _expand(self, text):
        for kind, tag, txt in blocks_of(text):
            if kind == 'subhead':
                self._subhead(tag)
            else:
                self._card(kind, tag, txt)

    def _apply_wrap(self):
        w = max(220, self.canvas.winfo_width() - 140)
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
            self._flash('已复制：' + (text[:40] + '…' if len(text) > 40 else text))
        except Exception as ex:
            self._flash('复制失败：' + str(ex))

    def _flash(self, msg):
        self.status.config(text=msg, fg=ACCENT)
        self.root.after(2200, lambda: self.status.config(text='', fg=MUTED))

    def toggle_top(self):
        self.topmost = not self.topmost
        self.root.attributes('-topmost', self.topmost)
        self.top_btn.config(text='置顶 ✓' if self.topmost else '置顶')

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    App().run()
