# -*- coding: utf-8 -*-
"""
Monte Carlo giá điều WW320 – bản đã điều chỉnh
================================================
Các thay đổi so với code gốc:
  1. Đơn vị đồng nhất: data theo THÁNG → N_PERIODS=12, DT=1/12
     (code gốc dùng DT=1/252 ngày với returns tháng → sigma bị thu nhỏ ~3.5x)
  2. Lưới quét có ý nghĩa kinh tế hơn: thay vì nhân hệ số ±50% quanh
     mu/sigma lịch sử, giờ quét mu theo năm (0%, -2%, -4%) ×
     sigma theo năm (±10%, ±25%, ±50% quanh sigma lịch sử).
  3. Fan chart dùng sigma_base (kịch bản trung tâm = lịch sử).
  4. Bổ sung AR(1) mean-reversion check và in cảnh báo nếu
     autocorrelation đủ lớn để cân nhắc dùng Ornstein–Uhlenbeck.
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")          # không cần màn hình
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

# =================================================================
# 1. ĐƯỜNG DẪN  –  BẠN TỰ CHỈNH
# =================================================================
INPUT_PATH  = r"C:\Users\Acer\Desktop\Dan on food\Code\monte carlo\input.xlsx"
OUTPUT_PATH = r"C:\Users\Acer\Desktop\Dan on food\Code\monte carlo\ouput_fixed.xlsx"

RETURN_COLUMN = "% change "    # cột returns (giá trị thập phân, VD -0.0597)
PRICE_COLUMN  = "Average value"
INPUT_SHEET   = None           # None = sheet đầu tiên

# =================================================================
# 2. THAM SỐ MONTE CARLO
# =================================================================
N_SIMS    = 10_000
N_PERIODS = 12          # ← 12 tháng = 1 năm (khớp với tần suất data)
DT        = 1 / 12      # ← bước thời gian theo NĂM (1/12 năm mỗi tháng)

SEED = 42

# =================================================================
# 3. LƯỚI QUÉT  (theo đơn vị NĂM)
# =================================================================
# mu (drift năm): quét 3 mức – không có xu hướng, giảm nhẹ, giảm vừa
MU_ANNUAL_GRID  = [0.00, -0.02, -0.04, -0.06, -0.08]

# sigma (vol năm): quét 5 mức xung quanh sigma lịch sử
# hệ số nhân: 0.75, 0.875, 1.0, 1.25, 1.5
SIG_FACTORS = [0.75, 0.875, 1.0, 1.25, 1.5]

# Kịch bản trung tâm cho fan chart: mu=0 (index 2), sig=1.0x (index 2)
CENTER_MU_IDX  = 0   # mu_annual = 0.00
CENTER_SIG_IDX = 2   # sig_factor = 1.0

# =================================================================
# 4. ĐỌC DỮ LIỆU
# =================================================================
def load_data(path, sheet, ret_col, price_col):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls", ".xlsm"):
        df = pd.read_excel(path, sheet_name=sheet if sheet else 0)
    elif ext == ".csv":
        df = pd.read_csv(path)
    else:
        raise ValueError(f"Định dạng không hỗ trợ: {ext}")

    if ret_col not in df.columns:
        raise ValueError(f"Không tìm thấy cột returns: '{ret_col}'. Các cột hiện có: {list(df.columns)}")
    if price_col not in df.columns:
        raise ValueError(f"Không tìm thấy cột giá: '{price_col}'. Các cột hiện có: {list(df.columns)}")

    returns = df[ret_col].dropna().astype(float).values
    prices  = df[price_col].dropna().astype(float).values

    if len(returns) < 12:
        raise ValueError("Cần ít nhất 12 điểm dữ liệu lịch sử.")
    return returns, prices

# =================================================================
# 5. KIỂM TRA MEAN REVERSION (AR1)
# =================================================================
def check_mean_reversion(returns):
    """
    Tính AR(1) autocorrelation của returns.
    Nếu |rho| > 0.15 thì in cảnh báo cân nhắc Ornstein–Uhlenbeck.
    """
    if len(returns) < 4:
        return
    rho = np.corrcoef(returns[:-1], returns[1:])[0, 1]
    print(f"[INFO] AR(1) autocorr returns: {rho:.3f}")
    if abs(rho) > 0.15:
        print(
            f"[WARN] |rho|={abs(rho):.3f} > 0.15 → chuỗi có xu hướng mean-reversion.\n"
            f"       GBM giả định bước độc lập. Kết quả vẫn dùng được nhưng\n"
            f"       mô hình Ornstein–Uhlenbeck sẽ phản ánh sát hơn đặc trưng giá điều."
        )

# =================================================================
# 6. GBM MỘT KỊCH BẢN
# =================================================================
def run_gbm(n_sims, n_periods, s0, mu_annual, sigma_annual, dt, rng):
    """
    GBM theo tháng:
        S_{t+1} = S_t * exp((mu - 0.5*sigma^2)*dt + sigma*sqrt(dt)*z)
    mu_annual, sigma_annual: tham số theo NĂM
    dt = 1/12 (năm/tháng)
    """
    drift = (mu_annual - 0.5 * sigma_annual ** 2) * dt
    diff  = sigma_annual * np.sqrt(dt)

    z        = rng.standard_normal((n_periods, n_sims))
    log_inc  = drift + diff * z
    log_path = np.vstack([np.zeros(n_sims), np.cumsum(log_inc, axis=0)])
    paths    = s0 * np.exp(log_path)          # shape (n_periods+1, n_sims)
    term_ret = paths[-1] / s0 - 1.0
    return paths, term_ret

# =================================================================
# 7. THỐNG KÊ
# =================================================================
def compute_stats(rets, conf_levels=(0.95, 0.99)):
    sorted_r = np.sort(rets)
    n = len(sorted_r)
    out = {
        "E[Return]"  : float(np.mean(rets)),
        "Std[Return]": float(np.std(rets, ddof=1)),
        "Median"     : float(np.median(rets)),
        "Min"        : float(sorted_r[0]),
        "Max"        : float(sorted_r[-1]),
    }
    for cl in conf_levels:
        alpha = 1 - cl
        idx   = max(1, int(np.floor(n * alpha)))
        pct   = int(round(cl * 100))
        out[f"VaR {pct}%"]  = float(sorted_r[idx - 1])
        out[f"CVaR {pct}%"] = float(np.mean(sorted_r[:idx]))
    return out

# =================================================================
# 8. MAIN
# =================================================================
def main():
    rng = np.random.default_rng(SEED)

    # ── 8.1 Đọc & kiểm tra dữ liệu ──────────────────────────────
    hist_rets, hist_prices = load_data(
        INPUT_PATH, INPUT_SHEET, RETURN_COLUMN, PRICE_COLUMN
    )
    check_mean_reversion(hist_rets)

    s0 = float(hist_prices[-1])

    # Tính mu, sigma THEO THÁNG từ lịch sử, rồi quy ra NĂM
    mu_m    = float(np.mean(hist_rets))
    sigma_m = float(np.std(hist_rets, ddof=1))
    mu_annual_base    = mu_m    * 12            # nhân 12  (cộng gộp đơn giản)
    sigma_annual_base = sigma_m * np.sqrt(12)   # nhân √12

    print(f"\n{'='*55}")
    print(f"[INFO] Số điểm lịch sử    : {len(hist_rets)}")
    print(f"[INFO] S0 (giá bắt đầu)   : {s0:.4f} USD/lb")
    print(f"[INFO] mu  tháng / năm     : {mu_m:.6f} / {mu_annual_base:.6f}")
    print(f"[INFO] sig tháng / năm     : {sigma_m:.6f} / {sigma_annual_base:.6f}")
    print(f"[INFO] N_PERIODS={N_PERIODS} tháng, DT={DT:.5f} năm")
    print(f"{'='*55}\n")

    # ── 8.2 Lưới kịch bản ────────────────────────────────────────
    stat_rows        = []
    mean_paths_store = {}
    fan_data         = None

    sc = 0
    for mi, mu_a in enumerate(MU_ANNUAL_GRID):
        for si, sf in enumerate(SIG_FACTORS):
            sc += 1
            sigma_a = sigma_annual_base * sf

            paths, rets = run_gbm(
                N_SIMS, N_PERIODS, s0, mu_a, sigma_a, DT, rng
            )

            mean_path   = paths.mean(axis=1)
            scen_label  = f"S{sc:02d} (mu_a={mu_a:+.3f}, sig_a={sigma_a:.4f})"
            mean_paths_store[scen_label] = mean_path

            st = compute_stats(rets)
            st.update({"Scenario": scen_label,
                        "Mu_annual": mu_a,
                        "Sigma_annual": sigma_a})
            stat_rows.append(st)

            # Fan chart dùng kịch bản trung tâm
            if mi == CENTER_MU_IDX and si == CENTER_SIG_IDX:
                fan_data = {
                    "period": np.arange(N_PERIODS + 1),
                    "P5" : np.percentile(paths,  5, axis=1),
                    "P25": np.percentile(paths, 25, axis=1),
                    "P50": np.percentile(paths, 50, axis=1),
                    "P75": np.percentile(paths, 75, axis=1),
                    "P95": np.percentile(paths, 95, axis=1),
                }

            print(f"[OK] {scen_label}  |  E[R]={st['E[Return]']:+.4f}  "
                  f"VaR95={st['VaR 95%']:.4f}  "
                  f"P5_price={s0*(1+st['VaR 95%'])**(1):.3f}")

    # ── 8.3 Xuất Excel ───────────────────────────────────────────
    stat_df    = pd.DataFrame(stat_rows)[[
        "Scenario", "Mu_annual", "Sigma_annual",
        "E[Return]", "Std[Return]",
        "VaR 95%", "VaR 99%", "CVaR 95%", "CVaR 99%",
        "Median", "Min", "Max",
    ]]
    mean_df = pd.DataFrame({"Period": np.arange(N_PERIODS + 1),
                              **mean_paths_store})
    fan_df  = pd.DataFrame(fan_data)

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        stat_df.to_excel(writer, sheet_name="Statistics", index=False)
        mean_df.to_excel(writer, sheet_name="MeanPaths",  index=False)
        fan_df.to_excel( writer, sheet_name="FanChart",   index=False)

    print(f"\n[INFO] Đã xuất: {OUTPUT_PATH}")

    # ── 8.4 Chart ────────────────────────────────────────────────
    months = np.arange(N_PERIODS + 1)

    # Chart 1: Mean paths tất cả scenario
    fig1, ax1 = plt.subplots(figsize=(13, 6))
    for label, mp in mean_paths_store.items():
        ax1.plot(months, mp, linewidth=1.2, label=label)
    ax1.axhline(s0, color="black", linewidth=1.5, linestyle="--",
                label=f"S0 = {s0:.3f}")
    ax1.set_title("Mean path – tất cả kịch bản quét (đơn vị tháng)")
    ax1.set_xlabel("Tháng")
    ax1.set_ylabel("Giá WW320 (USD/lb)")
    ax1.set_xticks(months)
    ax1.grid(True, alpha=0.3)
    ax1.legend(fontsize=7, loc="upper right", ncol=2)

    # Chart 2: Fan chart kịch bản trung tâm
    fig2, ax2 = plt.subplots(figsize=(13, 6))
    x = fan_df["period"]
    ax2.fill_between(x, fan_df["P5"],  fan_df["P95"],
                     alpha=0.15, color="steelblue", label="P5–P95 (90%)")
    ax2.fill_between(x, fan_df["P25"], fan_df["P75"],
                     alpha=0.30, color="steelblue", label="P25–P75 (50%)")
    ax2.plot(x, fan_df["P50"], color="steelblue",
             linewidth=2.5, label="Median (P50)")
    ax2.axhline(s0, color="black", linewidth=1.5, linestyle="--",
                label=f"S0 = {s0:.3f}")
    ax2.axhline(3.0, color="red", linewidth=1.2, linestyle=":",
                label="Ngưỡng rủi ro 3.0 USD/lb")

    # In giá trị tại tháng 12
    for col, pct in [("P5","5%"),("P50","50%"),("P95","95%")]:
        val = fan_df[col].iloc[-1]
        ax2.annotate(f"{pct}: {val:.3f}",
                     xy=(12, val), xytext=(10.5, val),
                     fontsize=9,
                     arrowprops=dict(arrowstyle="-", color="gray"))

    ax2.set_title(
        f"Fan chart – kịch bản trung tâm\n"
        f"mu_annual=0.00, sigma_annual={sigma_annual_base:.4f} "
        f"(σ lịch sử {len(hist_rets)} tháng quy năm)"
    )
    ax2.set_xlabel("Tháng")
    ax2.set_ylabel("Giá WW320 (USD/lb)")
    ax2.set_xticks(months)
    ax2.grid(True, alpha=0.3)
    ax2.legend(fontsize=9)

    # Lưu file
    out_dir    = os.path.dirname(OUTPUT_PATH) or "."
    chart1_pth = os.path.join(out_dir, "mc_mean_paths_fixed.png")
    chart2_pth = os.path.join(out_dir, "mc_fan_chart_fixed.png")
    pdf_pth    = os.path.join(out_dir, "mc_charts_fixed.pdf")
    fig1.tight_layout(); fig1.savefig(chart1_pth, dpi=150)
    fig2.tight_layout(); fig2.savefig(chart2_pth, dpi=150)
    with PdfPages(pdf_pth) as pdf:
        pdf.savefig(fig1); pdf.savefig(fig2)

    print(f"[INFO] Chart 1 : {chart1_pth}")
    print(f"[INFO] Chart 2 : {chart2_pth}")
    print(f"[INFO] PDF     : {pdf_pth}")

    # ── 8.5 Tóm tắt kịch bản trung tâm in ra console ────────────
    fan_end = {k: fan_df[k].iloc[-1] for k in ["P5","P25","P50","P75","P95"]}
    print(f"\n{'='*55}")
    print(f"Kịch bản trung tâm – giá tháng 12 (USD/lb):")
    for k, v in fan_end.items():
        chg = (v / s0 - 1) * 100
        print(f"  {k:4s}: {v:.3f}  ({chg:+.1f}% vs S0)")
    # Xác suất giá < 3.0 và > 3.5
    center_sc = [s for s in stat_rows
                 if s["Mu_annual"] == MU_ANNUAL_GRID[CENTER_MU_IDX]
                 and abs(s["Sigma_annual"] - sigma_annual_base * SIG_FACTORS[CENTER_SIG_IDX]) < 1e-8]
    if center_sc:
        st_c = center_sc[0]
        print(f"\n  E[Return 12 tháng] : {st_c['E[Return]']:+.4f}")
        print(f"  VaR 95%            : {st_c['VaR 95%']:.4f}  "
              f"(~{s0*(1+st_c['VaR 95%']):.3f} USD/lb)")
        print(f"  VaR 99%            : {st_c['VaR 99%']:.4f}  "
              f"(~{s0*(1+st_c['VaR 99%']):.3f} USD/lb)")
        print(f"  CVaR 95%           : {st_c['CVaR 95%']:.4f}")
    print(f"{'='*55}\n")

if __name__ == "__main__":
    main()