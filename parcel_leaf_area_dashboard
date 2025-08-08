import os, re
import pandas as pd
import numpy as np
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt

DATA_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_PATTERNS = [
    "19..07.2019 parsel6(2).xlsx",
    "19.07.2019 parsel 2.xlsx",
    "19.07.2019 parsel 10.xlsx",
    "19.07.2019 parsel1(1).xlsx",
    "19.07.2019 parsel1(2).xlsx",
    "19.07.2019 parsel3(1).xlsx",
    "19.07.2019 parsel3(2).xlsx",
    "19.07.2019 parsel4.xlsx",
    "19.07.2019 parsel5.xlsx",
    "19.07.2019 parsel6 (1).xlsx",
]

def normalize_date_str(s: str) -> str:
    return s.replace("..",".")

def parse_date_from_filename(fn: str):
    s = normalize_date_str(fn)
    m = re.search(r"(\d{2}\.\d{2}\.\d{4})", s)
    if m:
        try:
            return datetime.strptime(m.group(1), "%d.%m.%Y").date()
        except Exception:
            return None
    return None

def parse_parcel_from_filename(fn: str):
    m = re.search(r"parsel\s*([0-9]+)", fn.lower())
    if m:
        return int(m.group(1))
    m = re.search(r"parsel\s*([0-9]+)", fn.lower().replace(" ",""))
    if m:
        return int(m.group(1))
    m = re.search(r"([0-9]+)", fn)
    if m:
        return int(m.group(1))
    return None

def preferred_value_columns(df: pd.DataFrame):
    cols = [str(c).strip() for c in df.columns]
    name_hits = [c for c in cols if any(k in c.lower() for k in ["alan","yaprak","leaf","area","pixel","piksel","değer","deger","value"])]
    if name_hits:
        name_hits = [c for c in name_hits if pd.api.types.is_numeric_dtype(df[c]) or pd.to_numeric(df[c], errors="coerce").notna().any()]
        if name_hits:
            return name_hits
    numeric_cols = []
    for c in cols:
        ser = pd.to_numeric(df[c], errors="coerce")
        if ser.notna().mean() >= 0.5:
            numeric_cols.append(c)
    ignore = set()
    for c in numeric_cols:
        if re.search(r"\b(id|index| sıra | sira |no|num)\b", c.lower()):
            ignore.add(c)
    numeric_cols = [c for c in numeric_cols if c not in ignore]
    return numeric_cols

def load_all():
    parts = []
    for fn in FILE_PATTERNS:
        path = os.path.join(DATA_DIR, fn)
        if not os.path.exists(path):
            continue
        try:
            df = pd.read_excel(path)
        except Exception as e:
            print(f"Okuma hatası: {fn} -> {e}")
            continue
        df.columns = [str(c).strip() for c in df.columns]
        value_cols = preferred_value_columns(df)
        if not value_cols:
            continue
        melted = df[value_cols].copy().dropna(how="all")
        for c in melted.columns:
            melted[c] = pd.to_numeric(melted[c], errors="coerce")
        long_df = melted.melt(var_name="Ölçüm", value_name="Deger").dropna(subset=["Deger"])
        long_df["Dosya"] = fn
        long_df["Tarih"] = parse_date_from_filename(fn)
        long_df["Parsel"] = parse_parcel_from_filename(fn)
        parts.append(long_df)
    if parts:
        d = pd.concat(parts, ignore_index=True)
        d = d[~d["Parsel"].isna()]
        return d
    return pd.DataFrame(columns=["Ölçüm","Deger","Dosya","Tarih","Parsel"])

class Dashboard(tk.Tk):
    def __init__(self, df: pd.DataFrame):
        super().__init__()
        self.title("Parsel Yaprak Alanı Analiz Paneli v2")
        self.geometry("1280x780")
        self.df_raw = df.copy()  
        self.df = df.copy()

        
        top = ttk.Frame(self); top.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)

        # Parsel seçimi
        ttk.Label(top, text="Parsel:").pack(side=tk.LEFT, padx=(0,4))
        self.parsel_var = tk.StringVar(value="(Hepsi)")
        parsels = ["(Hepsi)"] + [str(int(x)) for x in sorted(self.df["Parsel"].dropna().unique())]
        self.parsel_cb = ttk.Combobox(top, textvariable=self.parsel_var, values=parsels, state="readonly", width=10)
        self.parsel_cb.pack(side=tk.LEFT, padx=4)
        self.parsel_cb.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        # Metrik seçimi
        ttk.Label(top, text="Metrik:").pack(side=tk.LEFT, padx=(12,4))
        self.metric_var = tk.StringVar(value="Ortalama")
        self.metric_cb = ttk.Combobox(top, textvariable=self.metric_var, values=["Ortalama","Medyan","Min","Max","StdSapma","Ölçüm_Sayısı"], state="readonly", width=14)
        self.metric_cb.pack(side=tk.LEFT, padx=4)
        self.metric_cb.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        # Birim dönüştürme
        ttk.Label(top, text="Ölçek (çarpan):").pack(side=tk.LEFT, padx=(12,4))
        self.scale_var = tk.StringVar(value="1.0")
        ttk.Entry(top, textvariable=self.scale_var, width=10).pack(side=tk.LEFT, padx=4)

        ttk.Label(top, text="Birim etiketi:").pack(side=tk.LEFT, padx=(12,4))
        self.unit_var = tk.StringVar(value="(birim)")
        ttk.Entry(top, textvariable=self.unit_var, width=12).pack(side=tk.LEFT, padx=4)

        # Uygula butonu
        ttk.Button(top, text="Ölçek/Birim Uygula", command=self.apply_scale_and_refresh).pack(side=tk.LEFT, padx=(10,4))

        filt = ttk.LabelFrame(self, text="Aykırı/Nümerik Filtreleme"); filt.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)

        # Z-score filtresi
        self.use_z_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(filt, text="Z-score filtresi", variable=self.use_z_var, command=self.refresh).pack(side=tk.LEFT, padx=8)
        ttk.Label(filt, text="k (±k·std):").pack(side=tk.LEFT, padx=(4,2))
        self.z_k_var = tk.StringVar(value="3.0")
        ttk.Entry(filt, textvariable=self.z_k_var, width=6).pack(side=tk.LEFT, padx=(0,12))

        # Değer aralığı
        ttk.Label(filt, text="Alt sınır:").pack(side=tk.LEFT, padx=(8,2))
        self.min_var = tk.StringVar(value="")
        ttk.Entry(filt, textvariable=self.min_var, width=10).pack(side=tk.LEFT, padx=(0,8))

        ttk.Label(filt, text="Üst sınır:").pack(side=tk.LEFT, padx=(4,2))
        self.max_var = tk.StringVar(value="")
        ttk.Entry(filt, textvariable=self.max_var, width=10).pack(side=tk.LEFT, padx=(0,8))

        ttk.Button(filt, text="Filtreleri Uygula", command=self.refresh).pack(side=tk.LEFT, padx=8)

    
        self.fig = plt.Figure(figsize=(8.5,4.2), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

       
        bottom = ttk.Frame(self); bottom.pack(side=tk.BOTTOM, fill=tk.BOTH, padx=8, pady=8)

        cols = ("Parsel","Ölçüm_Sayısı","Ortalama","Medyan","StdSapma","Min","Max","Birim")
        self.tree = ttk.Treeview(bottom, columns=cols, show="headings", height=10)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120 if c!="Birim" else 80, anchor=tk.CENTER)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(bottom, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=sb.set); sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Export buttons
        exp = ttk.Frame(self); exp.pack(side=tk.BOTTOM, fill=tk.X, padx=8, pady=(0,8))
        ttk.Button(exp, text="CSV olarak kaydet", command=self.export_csv).pack(side=tk.LEFT, padx=6)
        ttk.Button(exp, text="Excel olarak kaydet", command=self.export_xlsx).pack(side=tk.LEFT, padx=6)

  
        self.apply_scale_and_refresh()

    
    def get_scaled_filtered_df(self):
   
        dff = self.df_raw.copy()
     
        try:
            scale = float(self.scale_var.get().strip())
        except Exception:
            scale = 1.0
        dff["Deger"] = pd.to_numeric(dff["Deger"], errors="coerce") * scale
        # aralık filtresi
        mn_txt = self.min_var.get().strip()
        mx_txt = self.max_var.get().strip()
        if mn_txt != "":
            try:
                mn = float(mn_txt)
                dff = dff[dff["Deger"] >= mn]
            except Exception:
                pass
        if mx_txt != "":
            try:
                mx = float(mx_txt)
                dff = dff[dff["Deger"] <= mx]
            except Exception:
                pass
        
        if self.use_z_var.get():
            try:
                k = float(self.z_k_var.get().strip())
            except Exception:
                k = 3.0
            keep_idx = []
            for p, sub in dff.groupby("Parsel"):
                vals = sub["Deger"].astype(float)
                mu = vals.mean()
                sd = vals.std(ddof=0)
                if sd == 0 or np.isnan(sd):
                    keep_idx.extend(sub.index.tolist())
                else:
                    z = (vals - mu) / sd
                    mask = z.abs() <= k
                    keep_idx.extend(sub[mask].index.tolist())
            dff = dff.loc[sorted(set(keep_idx))]
        return dff

    def compute_summary(self, dff: pd.DataFrame):
        if dff.empty:
            out = pd.DataFrame(columns=["Parsel","Ölçüm_Sayısı","Ortalama","Medyan","StdSapma","Min","Max"])
        else:
            g = dff.groupby(["Parsel"])["Deger"]
            out = g.agg(["count","mean","median","std","min","max"]).reset_index()
            out = out.rename(columns={
                "count":"Ölçüm_Sayısı","mean":"Ortalama","median":"Medyan","std":"StdSapma","min":"Min","max":"Max"
            }).sort_values("Parsel")
        # Birim sütunu
        out["Birim"] = self.unit_var.get().strip()
        return out

    def refresh(self):
        dff = self.get_scaled_filtered_df()

        # parsel filtresi
        p = self.parsel_var.get()
        if p and p != "(Hepsi)":
            try:
                pv = int(p)
                dff = dff[dff["Parsel"] == pv]
            except Exception:
                pass

        summ = self.compute_summary(dff)

        # tablo
        for i in self.tree.get_children():
            self.tree.delete(i)
        for _, r in summ.iterrows():
            vals = (
                int(r["Parsel"]) if pd.notna(r["Parsel"]) else "",
                int(r["Ölçüm_Sayısı"]) if pd.notna(r["Ölçüm_Sayısı"]) else "",
                round(r["Ortalama"],3) if pd.notna(r["Ortalama"]) else "",
                round(r["Medyan"],3) if pd.notna(r["Medyan"]) else "",
                round(r["StdSapma"],3) if pd.notna(r["StdSapma"]) else "",
                round(r["Min"],3) if pd.notna(r["Min"]) else "",
                round(r["Max"],3) if pd.notna(r["Max"]) else "",
                r["Birim"]
            )
            self.tree.insert("", tk.END, values=vals)

        # grafik
        self.ax.clear()
        if not summ.empty:
            metric = self.metric_var.get()
            y = summ[metric] if metric in summ.columns else summ["Ortalama"]
            self.ax.bar(summ["Parsel"].astype(str).values, y.values)
            self.ax.set_title(f"Parsel Bazlı {metric} ({self.unit_var.get().strip()})")
            self.ax.set_xlabel("Parsel")
            self.ax.set_ylabel(f"{metric}")
            self.fig.tight_layout()
        else:
            self.ax.set_title("Veri bulunamadı")
        self.canvas.draw_idle()

    def apply_scale_and_refresh(self):
        self.refresh()

    def get_current_summary_df(self):
        dff = self.get_scaled_filtered_df()
        # parsel filtresi uygulansın
        p = self.parsel_var.get()
        if p and p != "(Hepsi)":
            try:
                pv = int(p)
                dff = dff[dff["Parsel"] == pv]
            except Exception:
                pass
        summ = self.compute_summary(dff)
        return summ

    def export_csv(self):
        df = self.get_current_summary_df()
        if df.empty:
            messagebox.showwarning("Uyarı", "Kaydedilecek veri yok.")
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        default = os.path.join(DATA_DIR, f"parsel_ozet_{ts}.csv")
        path = filedialog.asksaveasfilename(title="CSV olarak kaydet", defaultextension=".csv", initialfile=os.path.basename(default), filetypes=[("CSV","*.csv")])
        if not path:
            return
        df.to_csv(path, index=False, encoding="utf-8-sig")
        messagebox.showinfo("Bilgi", f"CSV kaydedildi:\n{path}")

    def export_xlsx(self):
        df = self.get_current_summary_df()
        if df.empty:
            messagebox.showwarning("Uyarı", "Kaydedilecek veri yok.")
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        default = os.path.join(DATA_DIR, f"parsel_ozet_{ts}.xlsx")
        path = filedialog.asksaveasfilename(title="Excel olarak kaydet", defaultextension=".xlsx", initialfile=os.path.basename(default), filetypes=[("Excel","*.xlsx")])
        if not path:
            return
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Özet", index=False)
        messagebox.showinfo("Bilgi", f"Excel kaydedildi:\n{path}")

def main():
    df = load_all()
    if df.empty:
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Hata", "Veri yüklenemedi. Excel dosyalarının script ile aynı klasörde olduğundan emin olun.")
        return
    app = Dashboard(df)
    app.mainloop()

if __name__ == "__main__":
    main()
