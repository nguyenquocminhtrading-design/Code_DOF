# -*- coding: utf-8 -*-
"""
Monte Carlo theo vùng quét
- Đọc dữ liệu lịch sử returns từ INPUT_PATH
- Quét lưới (mu, sigma) quanh giá trị ước lượng
- Xuất bảng thống kê (E[R], Std, VaR, CVaR, Median, Min, Max) ra OUTPUT_PATH
- Vẽ chart mean path các scenario + fan chart percentile bands
"""

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

# =================================================================
# 1. CẤU HÌNH ĐƯỜNG DẪN  (BẠN TỰ SETUP)
# =================================================================
INPUT_PATH  = r"C:\Users\Acer\Desktop\Dan on food\Code\monte carlo\input.xlsx"   # file dữ liệu lịch sử
OUTPUT_PATH = r"C:\Users\Acer\Desktop\Dan on food\Code\monte carlo\ouput.xlsx"    # file kết quả

# Cột chứa returns trong input (nếu là CSV/Excel nhiều cột)
RETURN_COLUMN = "% change "   # đổi tên cột cho khớp file của bạn

# Cột chứa giá trong input để lấy giá trị ban đầu (S0)
PRICE_COLUMN = "Average value"

# Sheet name nếu là Excel (để None nếu là sheet đầu tiên)
INPUT_SHEET = None

# =================================================================
# 2. THAM SỐ MONTE CARLO
# =================================================================
N_SIMS    = 10000       # số path mỗi scenario
N_PERIODS = 12        # số kỳ mỗi path (1 năm)
# S0 sẽ được lấy tự động từ file (xem hàm main)
DT        = 1 / 12    # bước thời gian

# Lưới quét: lệch ±50%, ±25%, 0 quanh giá trị base
MU_FACTORS  = [-0.50, -0.25, 0.0, 0.25, 0.50]
SIG_FACTORS = [-0.50, -0.25, 0.0, 0.25, 0.50]

SEED = 42              # None để random mỗi lần chạy

# =================================================================
# 3. ĐỌC DỮ LIỆU LỊCH SỬ
# =================================================================
def load_data(path, sheet=None, ret_col=None, price_col=None):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls", ".xlsm"):
        df = pd.read_excel(path, sheet_name=sheet if sheet else 0)
    elif ext == ".csv":
        df = pd.read_csv(path)
    else:
        raise ValueError(f"Định dạng không hỗ trợ: {ext}")

    if ret_col and ret_col in df.columns:
        returns = df[ret_col].dropna().astype(float).values
    else:
        raise ValueError(f"Không tìm thấy cột returns: {ret_col}")
        
    if price_col and price_col in df.columns:
        prices = df[price_col].dropna().astype(float).values
    else:
        raise ValueError(f"Không tìm thấy cột giá: {price_col}")

    if len(returns) < 5:
        raise ValueError("Cần ít nhất 5 điểm dữ liệu lịch sử.")
    return returns, prices


# =================================================================
# 4. CHẠY 1 KỊCH BẢN MONTE CARLO
# =================================================================
def run_one_mc(n_sims, n_periods, s0, mu, sigma, dt, rng):
    """
    Sinh paths theo GBM:
        S_{t+1} = S_t * exp((mu - 0.5*sigma^2)*dt + sigma*sqrt(dt)*z)
    Trả về:
        paths        : (n_periods+1, n_sims)
        terminal_ret : (n_sims,)  return cuối kỳ so với S0
    """
    drift = (mu - 0.5 * sigma ** 2) * dt
    diff  = sigma * np.sqrt(dt)

    z = rng.standard_normal((n_periods, n_sims))
    log_inc = drift + diff * z
    log_paths = np.vstack([np.zeros(n_sims), np.cumsum(log_inc, axis=0)])
    paths = s0 * np.exp(log_paths)           # (n_periods+1, n_sims)

    terminal_ret = paths[-1] / s0 - 1.0
    return paths, terminal_ret


# =================================================================
# 5. THỐNG KÊ: VaR, CVaR, E[R]
# =================================================================
def stats_from_returns(rets, conf_levels=(0.95, 0.99)):
    """
    Trả về dict:
        E[R], Std, Median, Min, Max,
        VaR_95, VaR_99, CVaR_95, CVaR_99
    """
    sorted_r = np.sort(rets)
    n = len(sorted_r)

    out = {
        "E[Return]": float(np.mean(rets)),
        "Std[Return]": float(np.std(rets, ddof=1)),
        "Median": float(np.median(rets)),
        "Min": float(sorted_r[0]),
        "Max": float(sorted_r[-1]),
    }

    for cl in conf_levels:
        alpha = 1 - cl
        idx = max(1, int(np.floor(n * alpha)))
        var = float(sorted_r[idx - 1])
        cvar = float(np.mean(sorted_r[:idx]))
        pct = int(round(cl * 100))
        out[f"VaR {pct}%"]  = var
        out[f"CVaR {pct}%"] = cvar

    return out


# =================================================================
# 6. MAIN: SWEEP + XUẤT KẾT QUẢ
# =================================================================
def main():
    rng = np.random.default_rng(SEED)

    # 6.1 Đọc dữ liệu
    hist_returns, hist_prices = load_data(INPUT_PATH, sheet=INPUT_SHEET, ret_col=RETURN_COLUMN, price_col=PRICE_COLUMN)
    
    # Tính S0
    # NẾU BẠN MUỐN DỰ PHÓNG TƯƠNG LAI TỪ HIỆN TẠI: Dùng giá trị CUỐI CÙNG (hist_prices[-1])
    # NẾU BẠN MUỐN CHẠY LẠI TỪ QUÁ KHỨ: Dùng giá trị ĐẦU TIÊN (hist_prices[0])
    s0 = hist_prices[-1]  # <--- Ở ĐÂY ĐANG DÙNG GIÁ TRỊ CUỐI CÙNG LÀM S0
    
    mu_base    = float(np.mean(hist_returns))
    sigma_base = float(np.std(hist_returns, ddof=1))

    print(f"[INFO] Số điểm lịch sử : {len(hist_returns)}")
    print(f"[INFO] S0 (Giá bắt đầu) : {s0:.4f}")
    print(f"[INFO] mu_base          : {mu_base:.6f}")
    print(f"[INFO] sigma_base       : {sigma_base:.6f}")

    # 6.2 Xác định scenario trung tâm
    center_mu_idx  = MU_FACTORS.index(0.0)
    center_sig_idx = SIG_FACTORS.index(0.0)

    # 6.3 Chuẩn bị lưu trữ
    stat_rows        = []
    mean_paths_store = {}     # scenario_id -> mean path (mảng)
    fan_data         = None   # percentile bands của scenario trung tâm

    sc = 0
    for mi, mf in enumerate(MU_FACTORS):
        for si, sf in enumerate(SIG_FACTORS):
            sc += 1
            mu    = mu_base * (1 + mf)
            sigma = sigma_base * (1 + sf)

            paths, rets = run_one_mc(
                N_SIMS, N_PERIODS, s0, mu, sigma, DT, rng
            )

            # mean path
            mean_path = paths.mean(axis=1)
            scen_label = f"S{sc} (mu={mu:.4f}, sig={sigma:.4f})"
            mean_paths_store[scen_label] = mean_path

            # thống kê
            st = stats_from_returns(rets)
            st.update({
                "Scenario": scen_label,
                "Mu": mu,
                "Sigma": sigma,
            })
            stat_rows.append(st)

            # fan chart cho scenario trung tâm
            if mi == center_mu_idx and si == center_sig_idx:
                p5  = np.percentile(paths, 5,  axis=1)
                p25 = np.percentile(paths, 25, axis=1)
                p50 = np.percentile(paths, 50, axis=1)
                p75 = np.percentile(paths, 75, axis=1)
                p95 = np.percentile(paths, 95, axis=1)
                fan_data = {
                    "period": np.arange(N_PERIODS + 1),
                    "P5": p5, "P25": p25, "P50": p50,
                    "P75": p75, "P95": p95,
                }

            print(f"[OK] {scen_label}")

    # 6.4 Bảng thống kê
    stat_df = pd.DataFrame(stat_rows)
    cols_order = [
        "Scenario", "Mu", "Sigma",
        "E[Return]", "Std[Return]",
        "VaR 95%", "VaR 99%",
        "CVaR 95%", "CVaR 99%",
        "Median", "Min", "Max",
    ]
    stat_df = stat_df[cols_order]

    # 6.5 DataFrame mean paths
    mean_path_df = pd.DataFrame(
        {"Period": np.arange(N_PERIODS + 1), **mean_paths_store}
    )

    # 6.6 DataFrame fan chart
    fan_df = pd.DataFrame(fan_data)

    # 6.7 Xuất Excel
    ext = os.path.splitext(OUTPUT_PATH)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
            stat_df.to_excel(writer,     sheet_name="Statistics", index=False)
            mean_path_df.to_excel(writer, sheet_name="MeanPaths",  index=False)
            fan_df.to_excel(writer,      sheet_name="FanChart",   index=False)
    elif ext == ".csv":
        stat_df.to_csv(OUTPUT_PATH, index=False)
    else:
        raise ValueError(f"Định dạng output không hỗ trợ: {ext}")

    print(f"\n[INFO] Đã xuất bảng thống kê: {OUTPUT_PATH}")

    # =============================================================
    # 7. VẼ CHART
    # =============================================================
    # 7.1 Chart 1: mean path của tất cả scenario
    fig1, ax1 = plt.subplots(figsize=(12, 6))
    for label, mp in mean_paths_store.items():
        ax1.plot(np.arange(N_PERIODS + 1), mp, linewidth=1.2, label=label)
    ax1.set_title("Mean path - tất cả kịch bản quét")
    ax1.set_xlabel("Kỳ")
    ax1.set_ylabel("Giá trị trung bình")
    ax1.grid(True, alpha=0.3)
    ax1.legend(fontsize=7, loc="upper left", ncol=2)

    # 7.2 Chart 2: fan chart percentile bands (scenario trung tâm)
    fig2, ax2 = plt.subplots(figsize=(12, 6))
    x = fan_df["period"]
    ax2.fill_between(x, fan_df["P5"],  fan_df["P95"], alpha=0.15, label="P5–P95")
    ax2.fill_between(x, fan_df["P25"], fan_df["P75"], alpha=0.30, label="P25–P75")
    ax2.plot(x, fan_df["P50"], linewidth=2, label="Median (P50)")
    ax2.set_title("Fan chart - Percentile bands (scenario trung tâm)")
    ax2.set_xlabel("Kỳ")
    ax2.set_ylabel("Giá trị")
    ax2.grid(True, alpha=0.3)
    ax2.legend()

    # 7.3 Lưu chart
    chart_dir = os.path.dirname(OUTPUT_PATH)
    chart1_path = os.path.join(chart_dir, "mc_mean_paths.png")
    chart2_path = os.path.join(chart_dir, "mc_fan_chart.png")
    fig1.tight_layout(); fig1.savefig(chart1_path, dpi=150)
    fig2.tight_layout(); fig2.savefig(chart2_path, dpi=150)
    print(f"[INFO] Chart 1: {chart1_path}")
    print(f"[INFO] Chart 2: {chart2_path}")

    # PDF gộp 2 chart (tuỳ chọn)
    pdf_path = os.path.join(chart_dir, "mc_charts.pdf")
    with PdfPages(pdf_path) as pdf:
        pdf.savefig(fig1)
        pdf.savefig(fig2)
    print(f"[INFO] PDF charts: {pdf_path}")

    plt.show()


if __name__ == "__main__":
    main()