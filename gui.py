#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""필러 아울렛 경쟁사 가격 조사 - 그래픽 창(GUI). 열리면 자동으로 조사 시작."""
import os, sys, glob, threading, queue, importlib.util, subprocess

HERE = os.path.dirname(os.path.abspath(__file__))

def ensure_deps(log=print):
    need = []
    for mod, pkg in [("requests","requests"),("bs4","beautifulsoup4"),
                     ("lxml","lxml"),("openpyxl","openpyxl"),("cloudscraper","cloudscraper"),
                     ("curl_cffi","curl_cffi")]:
        if importlib.util.find_spec(mod) is None:
            need.append(pkg)
    if need:
        log("필요한 부품 설치 중: " + ", ".join(need) + " ...")
        subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", *need])
        log("설치 완료.\n")

import tkinter as tk
from tkinter import scrolledtext, messagebox

class StdoutRedirect:
    def __init__(self, q): self.q = q
    def write(self, s):
        if s: self.q.put(s)
    def flush(self): pass

class App:
    def __init__(self, root):
        self.root = root
        root.title("필러 아울렛 경쟁사 가격 조사")
        root.geometry("660x480")
        root.configure(bg="#f4f6fb")
        tk.Label(root, text="경쟁사 가격 조사", font=("맑은 고딕", 16, "bold"),
                 bg="#f4f6fb", fg="#22305a").pack(pady=(16, 2))
        tk.Label(root, text="우리 제품을 자동으로 불러와 7개 경쟁사 가격을 조사합니다.",
                 font=("맑은 고딕", 9), bg="#f4f6fb", fg="#556").pack()
        tk.Label(root, text="※ 봇 차단 사이트는 '차단됨'으로 표시됩니다. 미국 VPN 권장.",
                 font=("맑은 고딕", 9), bg="#f4f6fb", fg="#a05").pack(pady=(0, 8))
        self.btn = tk.Button(root, text="  다시 조사  ", font=("맑은 고딕", 12, "bold"),
                             bg="#3b5bdb", fg="white", activebackground="#2f49b0",
                             relief="flat", padx=10, pady=8, command=self.start)
        self.btn.pack(pady=4)
        self.log = scrolledtext.ScrolledText(root, height=15, font=("Consolas", 9),
                                             bg="#0f1424", fg="#d7e0ff", insertbackground="white")
        self.log.pack(fill="both", expand=True, padx=14, pady=10)
        self.open_btn = tk.Button(root, text="결과 엑셀 열기", font=("맑은 고딕", 10),
                                  state="disabled", command=self.open_result,
                                  relief="flat", bg="#e7ecff", fg="#22305a", padx=8, pady=4)
        self.open_btn.pack(pady=(0, 12))
        self.q = queue.Queue()
        self.result_path = None
        self.root.after(100, self.drain)
        self.root.after(700, self.start)   # 창 열리면 자동 시작

    def write(self, s):
        self.log.insert("end", s); self.log.see("end")

    def drain(self):
        try:
            while True:
                self.write(self.q.get_nowait())
        except queue.Empty:
            pass
        self.root.after(100, self.drain)

    def start(self):
        self.btn.config(state="disabled", text="  조사 중...  ")
        self.open_btn.config(state="disabled")
        threading.Thread(target=self.worker, daemon=True).start()

    def worker(self):
        old = sys.stdout
        sys.stdout = StdoutRedirect(self.q)
        try:
            ensure_deps(log=lambda m: self.q.put(m + "\n"))
            spec = importlib.util.spec_from_file_location(
                "cpc", os.path.join(HERE, "competitor_price_check.py"))
            cpc = importlib.util.module_from_spec(spec); spec.loader.exec_module(cpc)
            cpc.run()
            files = [f for f in glob.glob(os.path.join(HERE, "경쟁사_가격비교_*.xlsx"))
                     if "TEST" not in f and "SMOKE" not in f]
            self.result_path = max(files, key=os.path.getmtime) if files else None
        except Exception as e:
            self.q.put(f"\n[오류] {e}\n")
        finally:
            sys.stdout = old
            self.root.after(0, self.done)

    def done(self):
        self.btn.config(state="normal", text="  다시 조사  ")
        if self.result_path:
            self.open_btn.config(state="normal")
            messagebox.showinfo("완료", "조사가 끝났어요!\n'결과 엑셀 열기' 버튼으로 확인하세요.")

    def open_result(self):
        if self.result_path and os.path.exists(self.result_path):
            try:
                os.startfile(self.result_path)
            except Exception as e:
                messagebox.showerror("열기 실패", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
