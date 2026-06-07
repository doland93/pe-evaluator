import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os, re, threading
from collections import defaultdict

import pdfplumber, fitz
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════
#  PARSER
# ═══════════════════════════════════════════════

def _pdf_text(path):
    text = ""
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t: text += t + "\n"
    except: pass
    if not text.strip():
        try:
            doc = fitz.open(path)
            for page in doc: text += page.get_text() + "\n"
        except: pass
    return text

def _parse_docx(path):
    doc = Document(path)
    result = {"evaluator": None, "team": None, "scores": []}
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) < 2: continue
            if re.match(r'Name\s*:', cells[0], re.I):
                nm = re.sub(r'Name\s*:\s*', '', cells[0], flags=re.I).strip()
                if nm: result["evaluator"] = nm
            if re.match(r'Team\s*:', cells[-1], re.I):
                tm = re.sub(r'Team\s*:\s*', '', cells[-1], flags=re.I).strip()
                m = re.search(r'(\d+)', tm)
                result["team"] = m.group(1) if m else tm
            if re.match(r'^[1-5]\s*[.。]', cells[0]):
                name_raw = re.sub(r'^[1-5]\s*[.。]\s*', '', cells[0]).strip()
                name_raw = re.sub(r'\d{7,10}', '', name_raw).strip()
                pts_raw = cells[-1].strip()
                if name_raw and re.match(r'^\d+$', pts_raw):
                    result["scores"].append((name_raw, int(pts_raw)))
    return result

def _parse_pdf(path):
    text = _pdf_text(path)
    result = {"evaluator": None, "team": None, "scores": []}
    m = re.search(r"Name\s*[:：]\s*([가-힣]+(?:\s+[가-힣]+)?)", text, re.I)
    if m: result["evaluator"] = m.group(1).strip()
    m = re.search(r"Team\s*[:：]\s*(\d+)", text, re.I)
    if m: result["team"] = m.group(1)
    else:
        m = re.search(r"T\s*e\s*a\s*m\s*[:：]\s*(\d+)", text, re.I)
        if m: result["team"] = m.group(1)
    for line in text.split("\n"):
        line = line.strip()
        m = re.match(
            r"^([1-5])\s*[.。]\s*(?:\d{7,10}\s+)?([가-힣]{2,6})\s+(\d{1,3})\s*$", line)
        if m: result["scores"].append((m.group(2).strip(), int(m.group(3))))
    return result

def parse_file(path):
    return _parse_docx(path) if path.lower().endswith('.docx') else _parse_pdf(path)

# ═══════════════════════════════════════════════
#  팀 자동 추정 (Union-Find)
# ═══════════════════════════════════════════════

def infer_teams(records):
    """
    상호 평가 관계로 팀 클러스터 자동 추정.
    반환: {name: inferred_team_str}
    """
    eval_graph = defaultdict(set)
    for rec in records:
        ev = rec["evaluator"]
        if not ev: continue
        for name, _ in rec["scores"]:
            eval_graph[ev].add(name)

    all_names = set(r["evaluator"] for r in records if r["evaluator"])
    parent = {n: n for n in all_names}

    def find(x):
        if parent[x] != x: parent[x] = find(parent[x])
        return parent[x]
    def union(x, y):
        parent[find(x)] = find(y)

    for a in all_names:
        for b in eval_graph[a]:
            if b in eval_graph and a in eval_graph[b]:
                union(a, b)

    # 클러스터별 다수결 팀 번호
    clusters = defaultdict(set)
    for name in all_names:
        clusters[find(name)].add(name)

    name_to_declared = {r["evaluator"]: r["team"] for r in records if r["evaluator"]}
    inferred = {}
    for root, members in clusters.items():
        votes = defaultdict(int)
        for name in members:
            t = name_to_declared.get(name)
            if t: votes[t] += 1
        best = max(votes, key=votes.get) if votes else "?"
        for name in members:
            inferred[name] = best

    return inferred

def detect_issues(records, inferred):
    """
    각 record에 대해 이슈를 감지.
    반환: list of issue strings per record index
    """
    issues = []
    for rec in records:
        ev = rec["evaluator"]
        rec_issues = []

        # ① 팀 번호 오기입
        if ev and inferred.get(ev) and rec["team"] != inferred[ev]:
            rec_issues.append(f"팀 오기입 의심 (기입:{rec['team']} → 추정:{inferred[ev]})")

        # ② 점수 합 체크
        total = sum(p for _, p in rec["scores"])
        if rec["scores"]:
            if total == 0:
                rec_issues.append("점수 합 = 0")
            elif total != 100:
                # 개인별 100점 만점 방식 감지: 모든 점수가 독립적으로 0~100
                all_leq100 = all(p <= 100 for _, p in rec["scores"])
                if all_leq100 and total > 100:
                    rec_issues.append(f"⚠ 개인별 100점 방식 의심 (합={total}) → 정규화 필요")
                else:
                    rec_issues.append(f"점수 합 = {total} (100이 아님)")

        issues.append(rec_issues)
    return issues

# ═══════════════════════════════════════════════
#  EXCEL BUILDER
# ═══════════════════════════════════════════════

_NAVY="1F4E79"; _LBLUE="D6E4F0"; _LBLUE2="EBF5FB"
_GREEN="E8F5E9"; _WARN="FFF9C4"; _ERR="FFCDD2"; _WHITE="FFFFFF"

def _fill(c): return PatternFill("solid", fgColor=c)
def _side(w="thin", c="BBBBBB"): return Side(style=w, color=c)
def _hdr(ws, r, col, val):
    c = ws.cell(row=r, column=col, value=val)
    c.font = Font(bold=True, color="FFFFFF", name="Malgun Gothic", size=10)
    c.fill = _fill(_NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = Border(left=_side(), right=_side(), top=_side(), bottom=_side())

def build_excel(records, out_path):
    team_members = defaultdict(set)
    for rec in records:
        t = rec["team"] or "?"
        if rec["evaluator"]: team_members[t].add(rec["evaluator"])
        for name, _ in rec["scores"]: team_members[t].add(name)

    eval_map = {}
    for rec in records:
        ev = rec["evaluator"]
        for name, pts in rec["scores"]:
            if ev: eval_map[(ev, name)] = pts

    teams = sorted(team_members.keys(), key=lambda x: (not x.isdigit(), int(x) if x.isdigit() else x))
    student_order = [(t, n) for t in teams for n in sorted(team_members[t])]
    max_peers = max(len(m)-1 for m in team_members.values())

    wb = Workbook()
    ws = wb.active; ws.title = "PE"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "C2"

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 7
    for i in range(max_peers):
        ws.column_dimensions[get_column_letter(3+i)].width = 9
    ws.column_dimensions[get_column_letter(3+max_peers)].width = 7
    ws.column_dimensions[get_column_letter(4+max_peers)].width = 9

    hdrs = ["학생명","Group"] + [f"평가자{i+1}" for i in range(max_peers)] + ["합계","Weight"]
    for col, h in enumerate(hdrs, 1): _hdr(ws, 1, col, h)
    ws.row_dimensions[1].height = 22

    team_rows = {}; row = 2
    for (team, name) in student_order:
        if team not in team_rows: team_rows[team] = {"start": row}
        team_rows[team]["end"] = row
        bf = _fill(_LBLUE2 if int(team) % 2 == 0 else _LBLUE) if team.isdigit() else _fill(_WHITE)

        for col, val in [(1, name), (2, int(team) if team.isdigit() else team)]:
            c = ws.cell(row=row, column=col, value=val)
            c.font = Font(name="Malgun Gothic", size=10)
            c.fill = bf; c.alignment = Alignment(horizontal="center")
            c.border = Border(left=_side(), right=_side(), top=_side(), bottom=_side())

        teammates = sorted(team_members[team] - {name})
        for i, ev in enumerate(teammates):
            pts = eval_map.get((ev, name))
            c = ws.cell(row=row, column=3+i, value=pts)
            c.font = Font(name="Malgun Gothic", size=10)
            c.fill = _fill(_WARN) if pts is None else _fill(_WHITE)
            c.alignment = Alignment(horizontal="center")
            c.border = Border(left=_side(), right=_side(), top=_side(), bottom=_side())
        for i in range(len(teammates), max_peers):
            c = ws.cell(row=row, column=3+i)
            c.fill = _fill("F0F0F0")
            c.border = Border(left=_side(), right=_side(), top=_side(), bottom=_side())

        sc = 3+max_peers
        sr = f"{get_column_letter(3)}{row}:{get_column_letter(2+max_peers)}{row}"
        c = ws.cell(row=row, column=sc, value=f"=SUM({sr})")
        c.font = Font(bold=True, name="Malgun Gothic", size=10)
        c.fill = _fill(_GREEN); c.alignment = Alignment(horizontal="center")
        c.border = Border(left=_side(), right=_side(), top=_side(), bottom=_side())

        wc = 4+max_peers
        c = ws.cell(row=row, column=wc, value=f"={get_column_letter(sc)}{row}/100")
        c.font = Font(bold=True, name="Malgun Gothic", size=10)
        c.fill = _fill(_GREEN); c.alignment = Alignment(horizontal="center")
        c.number_format = "0.00"
        c.border = Border(left=_side(), right=_side(), top=_side(), bottom=_side())

        ws.row_dimensions[row].height = 18; row += 1

    total_cols = 4+max_peers
    for team, rng in team_rows.items():
        s, e = rng["start"], rng["end"]
        for r in range(s, e+1):
            for col in range(1, total_cols+1):
                cell = ws.cell(row=r, column=col)
                old = cell.border
                cell.border = Border(left=old.left, right=old.right,
                    top=_side("medium",_NAVY) if r==s else old.top,
                    bottom=_side("medium",_NAVY) if r==e else old.bottom)

    # 제출현황 시트
    ws2 = wb.create_sheet("제출현황")
    ws2.sheet_view.showGridLines = False
    for col, (h, w) in enumerate(zip(
            ["이름","팀","제출여부","부여점수합","비고"],
            [12, 8, 14, 12, 20]), 1):
        _hdr(ws2, 1, col, h)
        ws2.column_dimensions[get_column_letter(col)].width = w
    ws2.row_dimensions[1].height = 22

    submitted = {rec["evaluator"] for rec in records if rec["evaluator"]}
    for i, (team, name) in enumerate(student_order, 2):
        ok = name in submitted
        pts_given = sum(p for (ev, _), p in eval_map.items() if ev == name)
        note = f"⚠ 합계 {pts_given}점" if ok and pts_given != 100 else ""
        rf = _fill(_WHITE) if ok else _fill(_ERR)
        for col, v in enumerate([name, team, "✔ 제출" if ok else "✘ 미제출",
                                  pts_given if ok else "-", note], 1):
            c2 = ws2.cell(row=i, column=col, value=v)
            c2.font = Font(name="Malgun Gothic", size=10,
                           color="006400" if ok else "C62828")
            c2.fill = rf; c2.alignment = Alignment(horizontal="center")
            c2.border = Border(left=_side(), right=_side(), top=_side(), bottom=_side())
        ws2.row_dimensions[i].height = 17

    wb.save(out_path)

# ═══════════════════════════════════════════════
#  GUI — 메인 창
# ═══════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PE 집계기  —  Peer Evaluation Processor")
        self.geometry("740x600")
        self.minsize(620, 500)
        self.configure(bg="#F0F4F8")
        self.files = []
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=f"#{_NAVY}", pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📋  Peer Evaluation 자동 집계기",
                 font=("Malgun Gothic",14,"bold"), fg="white", bg=f"#{_NAVY}").pack()
        tk.Label(hdr, text="PDF / DOCX 파일을 추가 → 파싱 확인 → 엑셀 생성",
                 font=("Malgun Gothic",9), fg="#BDD7EE", bg=f"#{_NAVY}").pack()

        lf = tk.LabelFrame(self, text=" 📂 추가된 파일 ",
                           font=("Malgun Gothic",10,"bold"),
                           bg="#F0F4F8", fg=f"#{_NAVY}", bd=2, relief="groove",
                           padx=8, pady=6)
        lf.pack(fill="both", expand=True, padx=16, pady=(14,4))

        self.lb = tk.Listbox(lf, font=("Consolas",9), selectmode="extended",
                             height=10, bg="white", fg="#222",
                             selectbackground=f"#{_NAVY}", selectforeground="white",
                             bd=0, highlightthickness=1, highlightbackground="#CCC")
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.lb.yview)
        self.lb.configure(yscrollcommand=sb.set)
        self.lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        bf = tk.Frame(self, bg="#F0F4F8")
        bf.pack(fill="x", padx=16, pady=4)
        S = dict(font=("Malgun Gothic",10,"bold"), relief="flat", padx=12, pady=6, cursor="hand2")
        tk.Button(bf, text="+ 파일 추가", bg="#2E86C1", fg="white",
                  command=self._add, **S).pack(side="left", padx=(0,6))
        tk.Button(bf, text="🗑 선택 삭제", bg="#7F8C8D", fg="white",
                  command=self._remove, **S).pack(side="left", padx=(0,6))
        tk.Button(bf, text="전체 초기화", bg="#C0392B", fg="white",
                  command=self._clear, **S).pack(side="left")

        pf = tk.LabelFrame(self, text=" 💾 저장 위치 ",
                           font=("Malgun Gothic",10,"bold"),
                           bg="#F0F4F8", fg=f"#{_NAVY}", bd=2, relief="groove",
                           padx=8, pady=6)
        pf.pack(fill="x", padx=16, pady=(6,4))
        pi = tk.Frame(pf, bg="#F0F4F8")
        pi.pack(fill="x")
        self.path_var = tk.StringVar(
            value=os.path.join(os.path.expanduser("~"), "Desktop", "PE_결과.xlsx"))
        tk.Entry(pi, textvariable=self.path_var, font=("Consolas",9),
                 bg="white", relief="solid", bd=1).pack(side="left", fill="x",
                                                         expand=True, padx=(0,6))
        tk.Button(pi, text="찾기", bg="#5D6D7E", fg="white",
                  font=("Malgun Gothic",9,"bold"), relief="flat",
                  padx=10, pady=4, cursor="hand2", command=self._browse).pack(side="right")

        self.run_btn = tk.Button(self, text="🔍  파싱 & 확인",
                                 bg=f"#{_NAVY}", fg="white",
                                 font=("Malgun Gothic",13,"bold"),
                                 relief="flat", padx=24, pady=10,
                                 cursor="hand2", command=self._run)
        self.run_btn.pack(pady=(10,4))

        self.status = tk.StringVar(value="파일을 추가하고 '파싱 & 확인' 버튼을 눌러주세요.")
        tk.Label(self, textvariable=self.status, font=("Malgun Gothic",9),
                 bg="#DDE3EA", fg="#333", anchor="w", padx=10, pady=5
                 ).pack(fill="x", side="bottom")

    def _add(self):
        paths = filedialog.askopenfilenames(
            title="PDF 또는 DOCX 파일 선택",
            filetypes=[("지원 파일","*.pdf *.docx"),("PDF","*.pdf"),("Word","*.docx")])
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.lb.insert("end", f"  {os.path.basename(p)}")
        self.status.set(f"총 {len(self.files)}개 파일 로드됨")

    def _remove(self):
        for i in reversed(self.lb.curselection()):
            self.lb.delete(i); self.files.pop(i)
        self.status.set(f"총 {len(self.files)}개 파일")

    def _clear(self):
        self.files.clear(); self.lb.delete(0, "end")
        self.status.set("초기화되었습니다.")

    def _browse(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx")], initialfile="PE_결과.xlsx")
        if p: self.path_var.set(p)

    def _run(self):
        if not self.files:
            messagebox.showwarning("파일 없음", "파일을 먼저 추가해주세요.")
            return
        self.run_btn.config(state="disabled", text="⏳  파싱 중...")
        threading.Thread(target=self._parse_all, daemon=True).start()

    def _parse_all(self):
        records, errors = [], []
        for path in self.files:
            try:
                rec = parse_file(path)
                rec["_file"] = os.path.basename(path)
                records.append(rec)
                self.after(0, lambda n=os.path.basename(path):
                           self.status.set(f"파싱: {n}"))
            except Exception as e:
                errors.append(f"{os.path.basename(path)}: {e}")

        if not records:
            self.after(0, lambda: messagebox.showerror("오류","파싱된 데이터가 없습니다."))
            self.after(0, lambda: self.run_btn.config(state="normal", text="🔍  파싱 & 확인"))
            return

        inferred = infer_teams(records)
        # 추정 팀으로 team 필드 업데이트 (원본 보존)
        for rec in records:
            ev = rec["evaluator"]
            rec["_team_declared"] = rec["team"]
            rec["_team_inferred"] = inferred.get(ev, rec["team"])
            rec["team"] = rec["_team_inferred"]  # 추정 팀을 기본값으로

        issues = detect_issues(records, inferred)
        self.after(0, lambda: self._show_review(records, issues, errors, self.path_var.get()))
        self.after(0, lambda: self.run_btn.config(state="normal", text="🔍  파싱 & 확인"))

    def _show_review(self, records, issues, errors, out_path):
        ReviewWindow(self, records, issues, errors, out_path)


# ═══════════════════════════════════════════════
#  GUI — 검토/수정 창
# ═══════════════════════════════════════════════

COL_NAME   = 0
COL_TEAM   = 1
COL_SCORES = 2
COL_TOTAL  = 3
COL_ISSUE  = 4
COL_FILE   = 5

class ReviewWindow(tk.Toplevel):
    def __init__(self, parent, records, issues, errors, out_path):
        super().__init__(parent)
        self.title("파싱 결과 확인 및 수정")
        self.geometry("1000x620")
        self.minsize(800, 500)
        self.configure(bg="#F0F4F8")
        self.records = records
        self.issues  = issues
        self.errors  = errors
        self.out_path = out_path
        self._build()
        self._populate()

    def _build(self):
        # 상단 안내
        top = tk.Frame(self, bg=f"#{_NAVY}", pady=10)
        top.pack(fill="x")
        tk.Label(top, text="📝  파싱 결과 확인  —  수정 후 엑셀 생성",
                 font=("Malgun Gothic",12,"bold"), fg="white", bg=f"#{_NAVY}").pack()
        tk.Label(top,
                 text="팀 번호를 더블클릭하면 수정 가능 | 노란색=경고, 빨간색=오류",
                 font=("Malgun Gothic",9), fg="#BDD7EE", bg=f"#{_NAVY}").pack()

        # 테이블 프레임
        tf = tk.Frame(self, bg="#F0F4F8")
        tf.pack(fill="both", expand=True, padx=14, pady=(10,4))

        cols = ("이름", "팀(기입→추정)", "점수 목록", "합계", "이슈", "파일명")
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", height=18)

        widths = [90, 120, 340, 60, 200, 180]
        for col, w in zip(cols, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="center" if w < 200 else "w")

        vsb = ttk.Scrollbar(tf, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(tf, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tf.rowconfigure(0, weight=1); tf.columnconfigure(0, weight=1)

        # 태그 색상
        self.tree.tag_configure("ok",   background="#FFFFFF")
        self.tree.tag_configure("warn", background="#FFF9C4")  # 노란: 개인별 100점 방식 의심
        self.tree.tag_configure("err",  background="#FFCDD2")  # 빨간: 심각한 오류
        self.tree.tag_configure("team_err", background="#FCE4EC")  # 연분홍: 팀 오기입

        self.tree.bind("<Double-1>", self._on_double_click)

        # 범례
        legend = tk.Frame(self, bg="#F0F4F8")
        legend.pack(fill="x", padx=14, pady=(0,4))
        for color, text in [("#FFF9C4","개인별 100점 방식 의심 (정규화 필요)"),
                             ("#FFCDD2","합계 오류 / 파싱 실패"),
                             ("#FCE4EC","팀 번호 오기입 의심 (자동 수정됨)")]:
            f = tk.Frame(legend, bg=color, width=16, height=16, relief="solid", bd=1)
            f.pack(side="left", padx=(8,3), pady=2)
            tk.Label(legend, text=text, font=("Malgun Gothic",8),
                     bg="#F0F4F8", fg="#555").pack(side="left", padx=(0,14))

        # 하단 버튼
        bf = tk.Frame(self, bg="#F0F4F8")
        bf.pack(fill="x", padx=14, pady=(4,10))

        # 요약 정보
        self.summary_var = tk.StringVar()
        tk.Label(bf, textvariable=self.summary_var,
                 font=("Malgun Gothic",9), bg="#F0F4F8", fg="#555"
                 ).pack(side="left")

        tk.Button(bf, text="⚡  엑셀 생성", bg="#1B5E20", fg="white",
                  font=("Malgun Gothic",12,"bold"), relief="flat",
                  padx=20, pady=8, cursor="hand2",
                  command=self._generate).pack(side="right")
        tk.Button(bf, text="닫기", bg="#7F8C8D", fg="white",
                  font=("Malgun Gothic",10,"bold"), relief="flat",
                  padx=14, pady=8, cursor="hand2",
                  command=self.destroy).pack(side="right", padx=(0,8))

    def _populate(self):
        self.tree.delete(*self.tree.get_children())

        warn_cnt = err_cnt = team_cnt = 0
        for i, rec in enumerate(self.records):
            ev    = rec.get("evaluator") or "?"
            team  = rec.get("team") or "?"
            tdecl = rec.get("_team_declared", team)
            tinf  = rec.get("_team_inferred", team)
            team_col = f"{tdecl} → {tinf}" if tdecl != tinf else team
            scores_str = "  ".join(f"{n}:{p}" for n, p in rec["scores"]) or "(없음)"
            total = sum(p for _, p in rec["scores"])
            issue_str = " | ".join(self.issues[i]) if self.issues[i] else "✓ 정상"
            fname = rec.get("_file","")

            # 태그 결정
            issue_texts = " ".join(self.issues[i])
            if "개인별 100점" in issue_texts:
                tag = "warn"; warn_cnt += 1
            elif self.issues[i] and "오기입" not in issue_texts:
                tag = "err"; err_cnt += 1
            elif "오기입" in issue_texts:
                tag = "team_err"; team_cnt += 1
            else:
                tag = "ok"

            iid = self.tree.insert("", "end",
                values=(ev, team_col, scores_str, total, issue_str, fname),
                tags=(tag,))
            self.tree.item(iid, tags=(tag,))

        total_rec = len(self.records)
        ok_cnt = total_rec - warn_cnt - err_cnt - team_cnt
        self.summary_var.set(
            f"총 {total_rec}명  |  ✓ 정상 {ok_cnt}  |  "
            f"⚠ 정규화필요 {warn_cnt}  |  팀오기입 {team_cnt}  |  ❌ 오류 {err_cnt}"
        )

    def _on_double_click(self, event):
        """팀 번호 셀 더블클릭 → 인라인 수정"""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        col_id = self.tree.identify_column(event.x)
        col_idx = int(col_id.replace("#","")) - 1
        if col_idx != 1: return  # 팀 컬럼만 수정 허용

        iid = self.tree.identify_row(event.y)
        if not iid: return
        row_idx = self.tree.index(iid)

        # 팝업 입력창
        cur_val = self.records[row_idx].get("team","")
        new_val = self._ask_team(cur_val)
        if new_val is None: return

        self.records[row_idx]["team"] = new_val
        self.records[row_idx]["_team_declared"] = new_val
        self.records[row_idx]["_team_inferred"] = new_val

        # 이슈 재계산
        inferred = {r["evaluator"]: r["team"] for r in self.records if r["evaluator"]}
        self.issues = detect_issues(self.records, inferred)
        self._populate()

    def _ask_team(self, current):
        dlg = tk.Toplevel(self)
        dlg.title("팀 번호 수정")
        dlg.geometry("280x120")
        dlg.resizable(False, False)
        dlg.configure(bg="#F0F4F8")
        dlg.grab_set()

        tk.Label(dlg, text="새 팀 번호를 입력하세요:",
                 font=("Malgun Gothic",10), bg="#F0F4F8").pack(pady=(16,4))
        var = tk.StringVar(value=current)
        entry = tk.Entry(dlg, textvariable=var, font=("Malgun Gothic",12),
                         justify="center", width=10)
        entry.pack(); entry.select_range(0, "end"); entry.focus()

        result = [None]
        def ok(e=None):
            result[0] = var.get().strip()
            dlg.destroy()
        def cancel():
            dlg.destroy()

        entry.bind("<Return>", ok)
        bf2 = tk.Frame(dlg, bg="#F0F4F8"); bf2.pack(pady=8)
        tk.Button(bf2, text="확인", command=ok, bg=f"#{_NAVY}", fg="white",
                  font=("Malgun Gothic",9,"bold"), relief="flat",
                  padx=12, pady=4).pack(side="left", padx=4)
        tk.Button(bf2, text="취소", command=cancel, bg="#7F8C8D", fg="white",
                  font=("Malgun Gothic",9,"bold"), relief="flat",
                  padx=12, pady=4).pack(side="left", padx=4)
        dlg.wait_window()
        return result[0]

    def _generate(self):
        try:
            build_excel(self.records, self.out_path)
            msg = f"✅  저장 완료!\n\n📁 {self.out_path}"
            if self.errors:
                msg += f"\n\n⚠️ 파싱 실패:\n" + "\n".join(self.errors)
            messagebox.showinfo("완료", msg, parent=self)
        except Exception as e:
            messagebox.showerror("오류", str(e), parent=self)


if __name__ == "__main__":
    App().mainloop()
