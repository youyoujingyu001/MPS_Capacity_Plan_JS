#!/usr/bin/env python3
"""
MPS 主计划约束表生成器 —— macOS GUI 入口 (Apple Silicon)

界面流程: 选数据源 -> 选源文件夹 -> 按关键词自动匹配 4~5 个源文件 -> 生成

实现要点 (打包成 .app 后仍然可用):
  冻结后的 sys.executable 是 App 自身二进制, 不能再用 subprocess 去调 Python 脚本,
  因此这里把核心脚本 MPS-Capacity_V2.0.py 作为模块 import 进来, 在同进程内调用
  run_prelaser() / run_fatp_sub_smt(), 并把 stdout/stderr 重定向到界面日志框。
"""
import os
import sys
import io
import glob
import queue
import threading
import importlib.util
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

CORE_NAME = 'MPS-Capacity_V2.0.py'

MODES = {
    'prelaser': {
        'label': 'prelaser  |  Pre-laser / FWDL / Fresh',
        'func': 'run_prelaser',
        'roles': [
            ('Capacity (含 PA 页)',  '*Capacity*.xlsx',       False),
            ('PreLaserIOS',         '*Pre*Laser*.xlsx',       False),
            ('PPVT_IOS',            '*PPVT*.xlsx',            False),
            ('Mapping',             '*mapping*.xlsx',         False),
            ('目标UPHmapping (可选)', '*UPHmapping*.xlsx',      True),
        ],
    },
    'fatp': {
        'label': 'fatp      |  FATP / SUB / SMT',
        'func': 'run_fatp_sub_smt',
        'roles': [
            ('SMT Mapping',       '*SMT*mapping*.xlsx',        False),
            ('Sub FATP Mapping',  '*Sub FATP*mapping*.xlsx',   False),
            ('FATP Capacity',     '*Capacity Simulation*.xlsx', False),
            ('SMT&Sub Capacity',  '*SMT & Sub Capacity*.xlsx', False),
            ('目标UPHmapping (可选)', '*UPHmapping*.xlsx',       True),
        ],
    },
}
# 排除规则: 避免 '*mapping*' 把《目标UPHmapping》误当成 Mapping 文件
EXCLUDE_KEYWORD = {'Mapping': 'UPHmapping'}


def find_core():
    """在多个候选位置里找核心脚本 (适配: 直接跑 / .app 打包后)"""
    exe_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    cands = [
        os.path.join(exe_dir, CORE_NAME),
        os.path.normpath(os.path.join(exe_dir, '..', 'Resources', CORE_NAME)),
        os.path.normpath(os.path.join(exe_dir, '..', '..', 'Resources', CORE_NAME)),
    ]
    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        cands.append(os.path.join(meipass, CORE_NAME))
    if '__file__' in globals():
        cands.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), CORE_NAME))
    cands.append(os.path.join(os.getcwd(), CORE_NAME))
    for c in cands:
        if c and os.path.isfile(c):
            return os.path.abspath(c)
    return None


def load_core(path):
    spec = importlib.util.spec_from_file_location('mps_core', path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def match_files(folder, roles):
    """按关键词匹配 -> [(role_label, picked_path|None, other_candidates, optional)]"""
    out = []
    for role_label, pattern, optional in roles:
        cands = []
        if folder and os.path.isdir(folder):
            for f in glob.glob(os.path.join(folder, pattern)):
                name = os.path.basename(f)
                if name.startswith('~$'):
                    continue
                kw = EXCLUDE_KEYWORD.get(role_label)
                if kw and kw.lower() in name.lower():
                    continue
                cands.append(f)
        cands.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        out.append((role_label, cands[0] if cands else None, cands[1:], optional))
    return out


class QueueWriter(io.TextIOBase):
    """把 print() 输出塞进队列, 由主线程写入日志框"""

    def __init__(self, q):
        self.q = q

    def write(self, s):
        if s:
            self.q.put(s)
        return len(s)

    def flush(self):
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('MPS 主计划约束表生成器')
        self.geometry('760x600')
        self.minsize(700, 540)
        self.folder_var = tk.StringVar()
        self.mode_labels = [v['label'] for v in MODES.values()]
        self.mode_var = tk.StringVar(value=self.mode_labels[0])
        self.role_rows = []
        self.log_q = queue.Queue()
        self.busy = False
        self.core_path = find_core()
        self._build()
        self._refresh()
        self._log(f'核心脚本: {self.core_path or "未找到 " + CORE_NAME}\n')
        if not self.core_path:
            messagebox.showwarning('缺少核心脚本',
                                   f'未找到 {CORE_NAME}\n请把它放在与本 App 相同的目录下')
        self.after(120, self._drain)

    # ---------- 界面 ----------
    def _build(self):
        pad = {'padx': 10, 'pady': 4}

        top = ttk.Frame(self)
        top.pack(fill='x', **pad)
        ttk.Label(top, text='数据源:').pack(side='left')
        cb = ttk.Combobox(top, values=self.mode_labels, textvariable=self.mode_var,
                          width=36, state='readonly')
        cb.pack(side='left', padx=6)
        cb.bind('<<ComboboxSelected>>', lambda e: self._rebuild_roles())

        row = ttk.Frame(self)
        row.pack(fill='x', **pad)
        ttk.Label(row, text='源文件夹:').pack(side='left')
        ttk.Entry(row, textvariable=self.folder_var).pack(side='left', fill='x', expand=True, padx=6)
        ttk.Button(row, text='选择文件夹', command=self._pick_folder).pack(side='left')
        ttk.Button(row, text='刷新匹配', command=self._refresh).pack(side='left', padx=4)

        self.files_frm = ttk.LabelFrame(self, text='文件匹配')
        self.files_frm.pack(fill='x', **pad)
        self._rebuild_roles()

        btns = ttk.Frame(self)
        btns.pack(fill='x', **pad)
        self.run_btn = ttk.Button(btns, text='▶ 生成主计划约束表', command=self._run)
        self.run_btn.pack(side='left')
        ttk.Button(btns, text='清空日志',
                   command=lambda: self.log_text.delete('1.0', 'end')).pack(side='left', padx=6)
        self.status = tk.StringVar(value='就绪')
        ttk.Label(btns, textvariable=self.status).pack(side='right')

        log_frm = ttk.LabelFrame(self, text='运行日志')
        log_frm.pack(fill='both', expand=True, **pad)
        font = ('Menlo', 11) if sys.platform == 'darwin' else ('Consolas', 10)
        self.log_text = scrolledtext.ScrolledText(log_frm, height=15, state='disabled', font=font)
        self.log_text.pack(fill='both', expand=True, padx=4, pady=4)

    def _mode_key(self):
        for k, v in MODES.items():
            if v['label'] == self.mode_var.get():
                return k
        return 'prelaser'

    def _rebuild_roles(self):
        for w in self.files_frm.winfo_children():
            w.destroy()
        self.role_rows.clear()
        for i, (role_label, pattern, optional) in enumerate(MODES[self._mode_key()]['roles']):
            r = ttk.Frame(self.files_frm)
            r.pack(fill='x', padx=6, pady=2)
            ttk.Label(r, text=('* ' if optional else '● ') + role_label,
                      width=24, anchor='w').pack(side='left')
            lbl = ttk.Label(r, text='(未匹配)' + ('  [可选, 跳过]' if optional else ''),
                            foreground='#888')
            lbl.pack(side='left', fill='x', expand=True)
            pv = tk.StringVar()
            ttk.Button(r, text='手选', command=lambda i=i: self._manual_pick(i)).pack(side='right')
            self.role_rows.append({'label': role_label, 'pattern': pattern,
                                   'optional': optional, 'path': pv, 'widget': lbl})
        self._refresh()

    # ---------- 交互 ----------
    def _pick_folder(self):
        f = filedialog.askdirectory(title='选择存放源文件的文件夹')
        if f:
            self.folder_var.set(f)
            self._refresh()

    def _refresh(self):
        if not self.role_rows:
            return
        folder = self.folder_var.get().strip()
        roles = [(r['label'], r['pattern'], r['optional']) for r in self.role_rows]
        for r, (_, picked, other, optional) in zip(self.role_rows, match_files(folder, roles)):
            if picked:
                r['path'].set(picked)
                r['widget'].configure(text=os.path.basename(picked), foreground='#2a7')
            else:
                r['path'].set('')
                extra = f'  (另有{len(other)}个候选, 点"手选")' if other else ''
                r['widget'].configure(
                    text='未匹配到' + extra + ('  [可选, 跳过]' if optional else '') + ('  ← 请选择' if not optional else ''),
                    foreground='#c00' if not optional else '#888')
        if folder and os.path.isdir(folder):
            miss = [r['label'] for r in self.role_rows if not r['path'].get() and not r['optional']]
            self.status.set('还缺: ' + ', '.join(miss) if miss else '文件已齐, 可以生成')
        else:
            self.status.set('请先选择源文件夹')

    def _manual_pick(self, idx):
        r = self.role_rows[idx]
        f = filedialog.askopenfilename(title=f'选择 {r["label"]}',
                                       filetypes=[('Excel 文件', '*.xlsx'), ('全部文件', '*.*')])
        if f:
            r['path'].set(f)
            r['widget'].configure(text=os.path.basename(f), foreground='#2a7')
            self._refresh_status_only()

    def _refresh_status_only(self):
        miss = [r['label'] for r in self.role_rows if not r['path'].get() and not r['optional']]
        self.status.set('还缺: ' + ', '.join(miss) if miss else '文件已齐, 可以生成')

    # ---------- 运行 ----------
    def _run(self):
        if self.busy:
            messagebox.showinfo('运行中', '上一次还在跑, 请稍等')
            return
        if not self.core_path:
            messagebox.showerror('缺少核心脚本', f'未找到 {CORE_NAME}, 无法运行')
            return
        files = []
        for r in self.role_rows:
            p = r['path'].get().strip()
            if not p:
                if r['optional']:
                    continue
                messagebox.showerror('缺少文件', f'{r["label"]} 还没选, 请点它右边的"手选"')
                return
            files.append(p)
        mode = self._mode_key()
        for f in files:
            if not os.path.exists(f):
                messagebox.showerror('文件不存在', f)
                return
        self.busy = True
        self.run_btn.configure(state='disabled')
        self.status.set('运行中...')
        self._log('\n' + '=' * 60 + f'\n开始运行 [{mode}] 模式\n' + '=' * 60 + '\n')
        threading.Thread(target=self._worker, args=(mode, files), daemon=True).start()

    def _worker(self, mode, files):
        func_name = MODES[mode]['func']
        old_out, old_err = sys.stdout, sys.stderr
        try:
            core = load_core(self.core_path)
            with_stdout = QueueWriter(self.log_q)
            sys.stdout = with_stdout
            sys.stderr = with_stdout
            getattr(core, func_name)(*files)
            self.log_q.put('\n[完成] 正常结束\n')
        except SystemExit as e:
            self.log_q.put(f'\n[结束] 脚本主动退出: {e}\n')
        except Exception:
            self.log_q.put('\n[ERROR] 运行失败:\n' + traceback.format_exc() + '\n')
        finally:
            sys.stdout, sys.stderr = old_out, old_err
            self.log_q.put('__DONE__')

    def _drain(self):
        try:
            while True:
                msg = self.log_q.get_nowait()
                if msg == '__DONE__':
                    self.busy = False
                    self.run_btn.configure(state='normal')
                    self.status.set('已结束, 结果见上方日志')
                    continue
                self.log_text.configure(state='normal')
                self.log_text.insert('end', msg)
                self.log_text.see('end')
                self.log_text.configure(state='disabled')
        except queue.Empty:
            pass
        self.after(120, self._drain)

    def _log(self, s):
        self.log_q.put(s)


if __name__ == '__main__':
    App().mainloop()
