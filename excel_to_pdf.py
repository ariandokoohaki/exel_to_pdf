# -*- coding: utf-8 -*-
"""
Excel → Persian Payslip PDF (enhanced)
Updated: 2025-07-27
---------------------------------------------------------------
• Drag an Excel file containing a «نام» (Name) column.
• One PDF is produced per unique name.
• Layout follows the “three-block” design with totals + net pay.
"""

import os
import sys
import threading
import traceback
import uuid
from datetime import datetime, timedelta

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES

from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Flowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Optional BiDi support
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    BIDI_SUPPORT = True
except ImportError:
    BIDI_SUPPORT = False

def setup_persian_font() -> str:
    candidates = [
        "Vazir.ttf", "Vazir-Regular.ttf", "BNazanin.ttf", "B-Nazanin.ttf",
        "Sahel.ttf", "IRANSans.ttf", "XB Niloofar.ttf"
    ]
    paths = [
        "", "fonts/", "C:/Windows/Fonts/",
        "/usr/share/fonts/truetype/", "/usr/share/fonts/truetype/vazir-font/",
        "/usr/local/share/fonts/", os.path.expanduser("~/.fonts/"),
        "/System/Library/Fonts/", "/Library/Fonts/"
    ]
    for fn in candidates:
        for base in paths:
            fp = os.path.join(base, fn)
            if os.path.isfile(fp):
                try:
                    pdfmetrics.registerFont(TTFont("Persian", fp))
                    return "Persian"
                except:
                    pass
    if not getattr(setup_persian_font, "_prompted", False):
        setup_persian_font._prompted = True
        root = tk.Tk(); root.withdraw()
        if messagebox.askyesno(
            "Font not found",
            "هیچ فونت فارسی مناسبی پیدا نشد.\n"
            "Would you like to locate a TTF font file manually?"
        ):
            fp = filedialog.askopenfilename(
                title="Select a Persian TTF file",
                filetypes=[("TrueType font", "*.ttf")])
            if fp:
                try:
                    pdfmetrics.registerFont(TTFont("Persian", fp))
                    return "Persian"
                except:
                    messagebox.showerror("Error", "Cannot load that font.")
    return "Helvetica"

PERSIAN_FONT = setup_persian_font()

def fix_rtl(text: str) -> str:
    if text is None: return ""
    text = str(text)
    if not BIDI_SUPPORT: return text
    try:
        return get_display(arabic_reshaper.reshape(text))
    except:
        return text

def fmt(v):
    """Format integers with commas, no decimals."""
    try:
        f = float(v)
        return f"{f:,.0f}"
    except:
        return v

def format_hm(v):
    """Format a duration (timedelta or hours float) as H:MM."""
    if isinstance(v, (timedelta, pd.Timedelta)):
        total_s = v.total_seconds()
    else:
        try:
            total_s = float(v) * 3600
        except:
            return fmt(v)
    h = int(total_s // 3600)
    m = int((total_s % 3600) // 60)
    return f"{h}:{m:02d}"

def read_excel(fp: str) -> pd.DataFrame:
    ext = os.path.splitext(fp)[1].lower()
    if ext == ".xlsb":
        try:
            return pd.read_excel(fp, engine="pyxlsb")
        except ImportError:
            raise ImportError("pip install pyxlsb")
    if ext in (".xlsx", ".xlsm", ".xlsb"):
        return pd.read_excel(fp, engine="openpyxl")
    return pd.read_excel(fp, engine="xlrd")

def get(df: pd.DataFrame, col: str, default="-"):
    try:
        v = df[col].iloc[0]
        return default if pd.isna(v) else v
    except KeyError:
        return default

class HRLine(Flowable):
    def __init__(self, width=17*cm):
        super().__init__(); self.width = width
    def draw(self):
        self.canv.setStrokeColor(colors.black)
        self.canv.setLineWidth(0.3)
        self.canv.line(0, 0, self.width, 0)

def make_block(title: str, rows: list, total_label: str=None, cell_w: float=3.2*cm) -> Table:
    cell_style = ParagraphStyle(
        "tbl", fontName=PERSIAN_FONT, fontSize=9,
        alignment=TA_RIGHT, leading=11
    )
    def p(text):
        return Paragraph(fix_rtl(text), cell_style)

    data = []
    for k, v in rows:
        data.append([p(str(v)), p(k)])
    if total_label:
        data.append([p(""), p(total_label)])

    tbl = Table(data, colWidths=[cell_w, cell_w], hAlign="CENTER")
    tbl.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),0.4,colors.black),
        ('FONTNAME',(0,0),(-1,-1),PERSIAN_FONT),
        ('FONTSIZE',(0,0),(-1,-1),9),
        ('ALIGN',(0,0),(-1,-1),'RIGHT'),
        ('ROWBACKGROUNDS',(0,0),(-1,-1),[colors.whitesmoke,colors.beige]),
    ]))
    heading = Table([[Paragraph(fix_rtl(title), ParagraphStyle(
        "heading", fontName=PERSIAN_FONT, fontSize=10,
        alignment=TA_CENTER, textColor=colors.white))]],
        colWidths=[cell_w*2]
    )
    heading.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),colors.grey),
        ('BOX',(0,0),(-1,-1),0.4,colors.black),
    ]))
    return Table([[heading],[tbl]])

def create_payslip(person_df: pd.DataFrame, out_dir: str, opts: dict, name_col="نام") -> str:
    person_name = str(person_df[name_col].iloc[0])
    safe = "".join(c for c in person_name if c.isalnum() or c in (" ","_","-")).strip()
    filename = f"{safe}-{uuid.uuid4().hex[:4]}.pdf"
    pdf_path = os.path.join(out_dir, filename)

    doc = SimpleDocTemplate(pdf_path,
        pagesize=landscape(A4),
        leftMargin=18, rightMargin=18, topMargin=18, bottomMargin=18
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("title", parent=styles["Title"],
        fontName=PERSIAN_FONT, fontSize=16, alignment=TA_CENTER, leading=22
    )
    normal_style = ParagraphStyle("normal", parent=styles["Normal"],
        fontName=PERSIAN_FONT, fontSize=9, alignment=TA_RIGHT
    )

    elems = [
        Paragraph(fix_rtl(opts["company"]), title_style),
        Paragraph(fix_rtl(opts["period"]),  title_style),
        Spacer(1, 0.4*cm)
    ]

    # Mini-header
    raw_phone = get(person_df, "شماره تماس", "")
    phone = raw_phone if raw_phone in ("-", "") else str(int(float(raw_phone)))
    mini_data = [
        (phone,           "شماره تماس"),
        (person_name,     "نام کامل"),
    ]
    mini_tbl = Table([[Paragraph(fix_rtl(v), normal_style), Paragraph(fix_rtl(k), normal_style)]
                      for v,k in mini_data],
                     colWidths=[3.5*cm, 3.5*cm]*len(mini_data))
    mini_tbl.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),0.4,colors.black),
        ('FONTNAME',(0,0),(-1,-1),PERSIAN_FONT),
        ('FONTSIZE',(0,0),(-1,-1),9),
        ('ALIGN',(0,0),(-1,-1),'RIGHT'),
    ]))
    elems += [mini_tbl, Spacer(1,0.4*cm)]

    # Durations
    raw_work = get(person_df, "مجموع ساعت کاری", 0)
    raw_delay = get(person_df, "تاخیر غیر مجاز", 0)
    work_str  = f"{format_hm(raw_work)} ساعت"
    delay_str = f"{format_hm(raw_delay)} ساعت"

    # Three blocks
    work_rows = [
        ("حقوق ساعتی پایه", fmt(get(person_df, "حقوق پایه"))),
        ("کارکرد ساعتی",    work_str),
        ("تعداد روز کارکرد", fmt(get(person_df, "روز کارکرد"))),
        ("تاخیر غیر مجاز",  delay_str),
    ]
    block_work = make_block("کارکرد", work_rows, cell_w=3.2*cm)

    benefit_rows = [
        ("بن مصرفی خواربار", fmt(get(person_df, "بن مصرفی"))),
        ("پاداش وقت شناسی",  fmt(get(person_df, "پاداش وقت شناسی"))),
        ("پاداش عملکرد",     fmt(get(person_df, "پاداش"))),
        ("ماموریت",          fmt(get(person_df, "مازاد مسکن"))),
    ]
    total_benefit = get(person_df, "جمع مزایا", 0)
    benefit_rows += [
        ("جمع مزایا", fmt(total_benefit)),
        ("جمع حقوق",  fmt(get(person_df, "جمع حقوق", 0))),
    ]
    block_benefit = make_block("مزایا", benefit_rows, cell_w=3.2*cm)

    deduction_rows = [
        ("بیمه",               fmt(get(person_df, "بیمه"))),
        ("مساعده",             fmt(get(person_df, "مساعده"))),
        ("اتلاف بن مصرفی",     fmt(get(person_df, "اتلاف بن مصرفی"))),
        ("مصرف ماه",           fmt(get(person_df, "مصرف ماه", 0))),
        ("جریمه تاخیر",        fmt(get(person_df, "جریمه تاخیر"))),
        ("بازپرداخت وام قرض الحسنه", fmt(get(person_df, "وام"))),
    ]
    total_deduction = get(person_df, "جمع کسور", 0)
    deduction_rows.append(("جمع کسور", fmt(total_deduction)))
    block_deduction = make_block("کسور", deduction_rows, cell_w=3.2*cm)

    three_tbl = Table([[block_deduction, block_benefit, block_work]],
                      colWidths=[7.2*cm]*3, hAlign="CENTER")
    three_tbl.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP')]))
    elems += [three_tbl, Spacer(1,0.3*cm)]

    # --- NET PAY (“جمع پرداختی”) from your EXCEL's “پرداختی” column ---
    raw_pay = get(person_df, "پرداختی", 0)
    pay_str = f"{fmt(raw_pay)} ریال"
    net_tbl = Table([[
        Paragraph(fix_rtl(pay_str), normal_style),
        Paragraph(fix_rtl("جمع پرداختی:"), normal_style)
    ]], colWidths=[4.0*cm, 3.0*cm])
    net_tbl.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),0.4,colors.black),
        ('FONTNAME',(0,0),(-1,-1),PERSIAN_FONT),
        ('FONTSIZE',(0,0),(-1,-1),9),
        ('ALIGN',(0,0),(-1,-1),'RIGHT'),
    ]))
    elems += [net_tbl, Spacer(1,0.2*cm)]

    # Disclaimer + timestamp
    elems += [
        Paragraph(fix_rtl(
            "★ حقوق ساعتی پایه با احتساب حق مسکن، سنوات و حق تاهل می‌باشد."
        ), normal_style),
        Spacer(1,0.2*cm),
        Paragraph(fix_rtl(
            f"تاریخ تولید گزارش : {datetime.now().strftime('%H:%M  %d-%m-%Y')}"
        ), normal_style),
    ]

    doc.build(elems)
    return pdf_path

class ExcelToPDFConverter:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel → Payslip PDF (Persian)")
        self.root.geometry("720x680")
        self.root.minsize(650, 600)
        self.file_path = ""
        self.output_dir = os.path.join(os.getcwd(), "payslips")
        self.name_column = "نام"
        self._stop_requested = False
        self._current_thread = None
        self._build_ui()

    def _build_ui(self):
        f = ttk.Frame(self.root, padding=14)
        f.grid(sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        ttk.Label(f, text="Excel → Payslip PDF", font=("Arial",18,"bold"))\
            .grid(row=0, column=0, columnspan=3)
        stat = ("✓ Persian shaping enabled" if BIDI_SUPPORT
                else "⚠ Install arabic-reshaper + python-bidi")
        col = "green" if BIDI_SUPPORT else "darkorange"
        ttk.Label(f, text=stat, foreground=col)\
            .grid(row=1, column=0, columnspan=3, pady=(4,12))

        ttk.Label(f, text="Company:").grid(row=2, column=0, sticky="e")
        self.ent_company = ttk.Entry(f, width=25)
        self.ent_company.insert(0, "کافه رود")
        self.ent_company.grid(row=2, column=1, sticky="w", padx=(0,8))

        ttk.Label(f, text="Period:").grid(row=3, column=0, sticky="e")
        self.ent_period = ttk.Entry(f, width=25)
        self.ent_period.grid(row=3, column=1, sticky="w", padx=(0,8))

        ttk.Label(f, text="Disclaimer (optional):").grid(row=4, column=0, sticky="e")
        self.ent_disclaimer = ttk.Entry(f, width=40)
        self.ent_disclaimer.insert(
            0,
            "حقوق ساعتی پایه با احتساب حق مسکن، سنوات و حق تاهل می‌باشد."
        )
        self.ent_disclaimer.grid(row=4, column=1, columnspan=2,
                                 sticky="we", pady=(0,12))

        self.drop = tk.Frame(f, height=120, width=480,
                             bg="lightgray", relief=tk.RIDGE, bd=4)
        self.drop.grid(row=5, column=0, columnspan=3, sticky="nsew")
        self.drop.pack_propagate(False)
        ttk.Label(self.drop,
                  text="Drag & Drop Excel here\nor press «Browse»",
                  background="lightgray").pack(expand=True)
        self.drop.drop_target_register(DND_FILES)
        self.drop.dnd_bind("<<Drop>>", self._on_drop)

        ttk.Button(f, text="Browse", command=self._browse)\
            .grid(row=6, column=0, columnspan=3, pady=6)

        self.file_lbl = ttk.Label(f, text="No file selected", foreground="gray")
        self.file_lbl.grid(row=7, column=0, columnspan=3, pady=(0,8))

        s = ttk.Style(); s.configure("Accent.TButton", foreground="blue")
        self.convert_btn = ttk.Button(f, text="Convert to PDF",
                                      style="Accent.TButton",
                                      command=self._convert,
                                      state="disabled")
        self.convert_btn.grid(row=8, column=0, columnspan=3, pady=10)

        self.cancel_btn = ttk.Button(f, text="Cancel",
                                     command=self._cancel,
                                     state="disabled")
        self.cancel_btn.grid(row=9, column=0, columnspan=3, pady=2)

        self.prog = ttk.Progressbar(f, mode="determinate", length=400)
        self.prog.grid(row=10, column=0, columnspan=3, sticky="ew", pady=4)

        self.stat_lbl = ttk.Label(f, text="")
        self.stat_lbl.grid(row=11, column=0, columnspan=3, pady=6)

    def _on_drop(self, ev):
        fp = ev.data.strip("{}")
        if fp.lower().endswith((".xls", ".xlsx", ".xlsm", ".xlsb")):
            self._load_file(fp)
        else:
            messagebox.showerror("Invalid file", "Please drop an Excel file.")

    def _browse(self):
        fp = filedialog.askopenfilename(
            filetypes=[("Excel files","*.xls *.xlsx *.xlsm *.xlsb")])
        if fp:
            self._load_file(fp)

    def _load_file(self, path):
        self.file_path = path
        self.ent_period.delete(0, tk.END)
        self.ent_period.insert(0, os.path.splitext(os.path.basename(path))[0])
        self.file_lbl.configure(text=f"Selected: {os.path.basename(path)}",
                                foreground="green")
        self.convert_btn["state"] = "normal"

    def _convert(self):
        if not self.file_path: return
        self.convert_btn["state"]  = "disabled"
        self.cancel_btn["state"]   = "normal"
        self.prog["value"]         = 0
        self.stat_lbl["text"]      = ""
        self._stop_requested       = False
        self._current_thread       = threading.Thread(target=self._worker, daemon=True)
        self._current_thread.start()

    def _cancel(self):
        self._stop_requested = True

    def _worker(self):
        try:
            df = read_excel(self.file_path)
            if self.name_column not in df.columns:
                raise ValueError(f"Column «{self.name_column}» not found.")
            os.makedirs(self.output_dir, exist_ok=True)

            names = df[self.name_column].dropna().unique()
            total = len(names)
            ok = bad = 0
            self.prog["maximum"] = total

            opts = {
                "company": self.ent_company.get().strip() or "-",
                "period":  self.ent_period.get().strip()  or "-",
            }

            for idx, nm in enumerate(names, 1):
                if self._stop_requested: break
                self.stat_lbl.after(0, lambda n=nm, i=idx, t=total:
                                    self.stat_lbl.configure(text=f"{i}/{t} → {n}"))
                self.prog.after(0, lambda v=idx: self.prog.configure(value=v))
                try:
                    create_payslip(df[df[self.name_column]==nm], self.output_dir, opts)
                    ok += 1
                except Exception as e:
                    bad += 1
                    with open(os.path.join(self.output_dir,"converter_error.log"),
                              "a", encoding="utf-8") as logf:
                        logf.write(f"\n\n[{datetime.now()}] {nm} ➜ {e}\n")
                        traceback.print_exc(file=logf)

            msg = f"✅ {ok} PDF(s) created"
            if bad: msg += f"  –  ❌ {bad} failed"
            if self._stop_requested: msg = "⏹ Operation cancelled.\n" + msg

            messagebox.showinfo("Finished", msg + f"\n\n{self.output_dir}")
            if ok and sys.platform.startswith("win") and not self._stop_requested:
                os.startfile(self.output_dir)

        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror("Error", str(exc))
        finally:
            self.prog.after(0, lambda: self.prog.configure(value=0))
            self.convert_btn.after(0, lambda: self.convert_btn.configure(state="normal"))
            self.cancel_btn.after(0, lambda: self.cancel_btn.configure(state="disabled"))
            self._current_thread = None

def check_requirements():
    missing = []
    for pkg in ("pandas","openpyxl","xlrd","reportlab","tkinterdnd2"):
        try:
            __import__(pkg.split("-")[0])
        except ImportError:
            missing.append(pkg)
    if missing:
        messagebox.showerror(
            "Missing libs",
            "Required but not installed:\n\n  " + ", ".join(missing) +
            "\n\nPlease install via pip."
        )
        sys.exit(1)

def main():
    check_requirements()
    root = TkinterDnD.Tk()
    ExcelToPDFConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()
