# -*- coding: utf-8 -*-
"""
生成《商用統計學》期中考精選 40 題全套專業統計圖表 (Matplotlib 300 DPI)
"""

import os
import matplotlib
import matplotlib.pyplot as plt
import numpy as np

# 設定微軟正黑體與負號正常顯示
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'Segoe UI', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 10

NAVY = "#1B365D"
SLATE = "#4A5568"
BLUE = "#2B6CB0"
LIGHT_BLUE = "#EBF3FA"
ACCENT = "#D9534F"
GREEN = "#2E7D32"
AMBER = "#D97706"

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)

def generate_all_figures(fig_dir):
    ensure_dir(fig_dir)
    print(f"Generating charts in: {fig_dir}")

    # ==================== CHAPTER 2 ====================
    # Ch2 Q1: 百事可樂圓形圖
    fig, ax = plt.subplots(figsize=(5, 4.2), dpi=300)
    sizes = [25, 40, 20, 15]
    labels = ['百事可樂\n(90°，25%)', '可口可樂\n(40%)', 'Cola-Plus\n(20%)', '其他\n(15%)']
    colors = [BLUE, '#3182CE', '#63B3ED', '#BEE3F8']
    explode = (0.08, 0, 0, 0)
    wedges, texts, autotexts = ax.pie(sizes, explode=explode, labels=labels, autopct='%1.0f%%',
                                      startangle=90, colors=colors, textprops=dict(color="black", fontsize=10))
    for at in autotexts:
        at.set_color('white')
        at.set_weight('bold')
    ax.set_title('第 2 章 習題 1：可樂相對市占率圓形圖（百事可樂 90° 占 25%）', fontsize=11, fontweight='bold', color=NAVY, pad=12)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch2_q1.png"))
    plt.close()

    # Ch2 Q9: 組距與組數 2^k 判定圖
    fig, ax = plt.subplots(figsize=(5.5, 3.8), dpi=300)
    k_vals = np.array([5, 6, 7, 8, 9])
    two_k = 2 ** k_vals
    bars = ax.bar(k_vals, two_k, color=BLUE, width=0.5, alpha=0.85, label='2^k 涵蓋容量')
    ax.axhline(230, color=ACCENT, linestyle='--', linewidth=1.5, label='樣本數 n = 230')
    bars[3].set_color(NAVY) # k=8
    ax.text(8, 260, 'k=8: 256 >= 230\n(建議組數)', ha='center', va='bottom', fontweight='bold', color=NAVY, fontsize=9.5)
    ax.set_xlabel('分組組數 (k)', fontsize=10)
    ax.set_ylabel('2^k 次方值', fontsize=10)
    ax.set_title('第 2 章 習題 9：2^k 法則決定組數 (k = 8 組)', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xticks(k_vals)
    ax.legend(frameon=True, facecolor=LIGHT_BLUE)
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch2_q9.png"))
    plt.close()

    # Ch2 Q11: Wachesaw 工廠生產量直方圖
    fig, ax = plt.subplots(figsize=(5.5, 3.8), dpi=300)
    bins = [24.0, 25.5, 27.0, 28.5, 30.0, 31.5]
    counts = [2, 4, 8, 0, 2]
    centers = [(bins[i] + bins[i+1])/2 for i in range(len(counts))]
    widths = [1.5] * len(counts)
    bars = ax.bar(centers, counts, width=widths, color=BLUE, edgecolor='white', linewidth=1.2, alpha=0.9)
    bars[2].set_color(NAVY) # 眾數組
    ax.set_title('第 2 章 習題 11：Wachesaw 工廠日生產量直方圖', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('生產數量區間（組距 1.5）', fontsize=10)
    ax.set_ylabel('天數 (次數 f)', fontsize=10)
    ax.set_xticks(bins)
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    for c, cnt in zip(centers, counts):
        if cnt > 0:
            ax.text(c, cnt + 0.2, f'{cnt}天\n({cnt/16:.1%})', ha='center', va='bottom', fontsize=9, fontweight='bold')
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch2_q11.png"))
    plt.close()

    # Ch2 Q15: Sun Ray Candles 蠟燭出貨直方圖
    fig, ax = plt.subplots(figsize=(6, 3.8), dpi=300)
    bins = [0, 5, 10, 15, 20, 25, 30, 35]
    counts = [5, 13, 28, 23, 18, 10, 3]
    centers = [(bins[i] + bins[i+1])/2 for i in range(len(counts))]
    widths = [5] * len(counts)
    bars = ax.bar(centers, counts, width=widths, color=BLUE, edgecolor='white', linewidth=1.2, alpha=0.9)
    bars[2].set_color(NAVY) # 眾數組 10~15
    for c, cnt in zip(centers, counts):
        ax.text(c, cnt + 0.6, str(cnt), ha='center', va='bottom', fontsize=9.5, fontweight='bold')
    ax.set_title('第 2 章 習題 15：Sun Ray Candles 最近 100 天出貨數直方圖', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('出貨數量（包）', fontsize=10)
    ax.set_ylabel('天數 (次數 f)', fontsize=10)
    ax.set_xticks(bins)
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch2_q15.png"))
    plt.close()

    # Ch2 Q17: Brumley 顧問直方圖與次數多邊形
    fig, ax = plt.subplots(figsize=(6, 4), dpi=300)
    bins = [0, 3, 6, 9, 12, 15]
    counts = [5, 12, 23, 8, 2]
    midpoints = [1.5, 4.5, 7.5, 10.5, 13.5]
    widths = [3] * len(counts)
    ax.bar(midpoints, counts, width=widths, color=LIGHT_BLUE, edgecolor=BLUE, linewidth=1.2, label='直方圖 (柱寬 = 組距 3)')
    poly_x = [-1.5] + midpoints + [16.5]
    poly_y = [0] + counts + [0]
    ax.plot(poly_x, poly_y, color=NAVY, marker='o', linewidth=2, markersize=6, label='次數多邊形 (連接各組中點)')
    for x, y in zip(midpoints, counts):
        ax.text(x, y + 0.8, f'({x}, {y})', ha='center', va='bottom', fontsize=8.5, fontweight='bold', color=NAVY)
    ax.set_title('第 2 章 習題 17：員工飛行里程直方圖與次數多邊形圖', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('飛行里程數（千英里）', fontsize=10)
    ax.set_ylabel('員工數量 (f)', fontsize=10)
    ax.set_xticks(bins)
    ax.legend(frameon=True, facecolor='white')
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch2_q17.png"))
    plt.close()

    # Ch2 Q19: 焊接工人時薪肩峰圖
    fig, ax = plt.subplots(figsize=(6, 4), dpi=300)
    wage_x = [0, 5, 10, 15, 20, 25, 30]
    cum_y = [0, 2, 11, 24, 31, 37, 40]
    ax.plot(wage_x, cum_y, color=NAVY, marker='s', linewidth=2.2, markersize=6, label='累加次數曲線')
    ax.axhline(30, color=ACCENT, linestyle=':', label='75% 累積門檻 (30人, 時薪約 $18)')
    ax.axvline(18, color=ACCENT, linestyle=':')
    ax.axhline(11, color=AMBER, linestyle=':', label='時薪低於 $10 人數 (約 11~12人)')
    ax.axvline(10, color=AMBER, linestyle=':')
    ax.set_title('第 2 章 習題 19：焊接工時薪之累加次數多邊形圖（肩峰圖）', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('時薪 ($)', fontsize=10)
    ax.set_ylabel('累加人數（上限 40 人）', fontsize=10)
    ax.set_xticks(wage_x)
    ax.set_yticks([0, 10, 20, 30, 40])
    ax.legend(frameon=True, facecolor='white', fontsize=8.5)
    ax.grid(True, linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch2_q19.png"))
    plt.close()

    # Ch2 Q21: Brumley 顧問飛行里程肩峰圖
    fig, ax = plt.subplots(figsize=(6, 4), dpi=300)
    mile_x = [0, 3, 6, 9, 12, 15]
    cum_m = [0, 5, 17, 40, 48, 50]
    ax.plot(mile_x, cum_m, color=BLUE, marker='D', linewidth=2.2, markersize=6, label='累加次數曲線 (Ogive)')
    ax.axhline(37.5, color=ACCENT, linestyle='--', label='75% 人數 (37.5人)')
    ax.axvline(8.7, color=ACCENT, linestyle='--', label='對應里程約 8.7 千英里 (8,700 英里)')
    ax.scatter([8.7], [37.5], color=ACCENT, s=80, zorder=5)
    ax.set_title('第 2 章 習題 21：搭機里程之累加次數多邊形圖（肩峰圖）', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('飛行里程上限（千英里）', fontsize=10)
    ax.set_ylabel('累加員工數 (CF)', fontsize=10)
    ax.set_xticks(mile_x)
    ax.set_yticks([0, 10, 20, 30, 40, 50])
    ax.legend(frameon=True, facecolor=LIGHT_BLUE, fontsize=8.5)
    ax.grid(True, linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch2_q21.png"))
    plt.close()

    # Ch2 Q23: 變數分類與測量尺度架構圖
    fig, ax = plt.subplots(figsize=(6.5, 3.5), dpi=300)
    ax.axis('off')
    # 繪製樹狀階層框
    boxes = [
        (0.5, 0.85, '統計變數 (Variables)', NAVY),
        (0.25, 0.55, '質變數 (定性/類別)\nQualitative', BLUE),
        (0.75, 0.55, '量變數 (定量/數值)\nQuantitative', BLUE),
        (0.12, 0.2, '名目尺度\n(Nominal)', SLATE),
        (0.38, 0.2, '順序尺度\n(Ordinal)', SLATE),
        (0.62, 0.2, '區間尺度\n(Interval)', SLATE),
        (0.88, 0.2, '比例尺度\n(Ratio)', SLATE)
    ]
    for x, y, text, col in boxes:
        ax.text(x, y, text, ha='center', va='center', bbox=dict(boxstyle='round,pad=0.5', facecolor=LIGHT_BLUE, edgecolor=col, linewidth=1.5), fontsize=9, fontweight='bold', color=col)
    lines = [
        ((0.5, 0.77), (0.25, 0.65)), ((0.5, 0.77), (0.75, 0.65)),
        ((0.25, 0.45), (0.12, 0.28)), ((0.25, 0.45), (0.38, 0.28)),
        ((0.75, 0.45), (0.62, 0.28)), ((0.75, 0.45), (0.88, 0.28))
    ]
    for p1, p2 in lines:
        ax.annotate('', xy=p2, xytext=p1, arrowprops=dict(arrowstyle='->', color=SLATE, lw=1.2))
    ax.set_title('第 2 章 習題 23：質變數與量變數及其測量尺度體系', fontsize=11, fontweight='bold', color=NAVY, pad=10)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch2_q23.png"))
    plt.close()

    # Ch2 Q25: 度假村活動偏好長條圖與圓形圖 (雙圖)
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(7.2, 3.6), dpi=300)
    cats = ['有計畫', '無計畫', '不確定', '沒答覆']
    vals = [63, 135, 78, 24]
    cols = ['#4299E1', NAVY, '#90CDF4', '#CBD5E0']
    ax1.bar(cats, vals, color=cols, width=0.6)
    for i, v in enumerate(vals):
        ax1.text(i, v + 3, str(v), ha='center', va='bottom', fontweight='bold', fontsize=9)
    ax1.set_title('(b) 長條圖（人數）', fontsize=10, fontweight='bold', color=NAVY)
    ax1.set_ylabel('顧客人數', fontsize=9)
    ax1.grid(axis='y', linestyle=':', alpha=0.6)
    
    ax2.pie(vals, labels=cats, autopct='%1.1f%%', colors=cols, startangle=140, textprops=dict(fontsize=8.5))
    ax2.set_title('(c) 圓形圖（結構比）', fontsize=10, fontweight='bold', color=NAVY)
    fig.suptitle('第 2 章 習題 25：Myrtle Beach 度假村顧客休閒活動偏好', fontsize=11, fontweight='bold', color=NAVY)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch2_q25.png"))
    plt.close()

    # Ch2 Q27: 83 筆資料之 2^k 組數判定圖
    fig, ax = plt.subplots(figsize=(5.5, 3.8), dpi=300)
    k_vals27 = np.array([4, 5, 6, 7, 8])
    two_k27 = 2 ** k_vals27
    bars27 = ax.bar(k_vals27, two_k27, color=BLUE, width=0.5, alpha=0.85, label='2^k 涵蓋容量')
    ax.axhline(83, color=ACCENT, linestyle='--', linewidth=1.5, label='樣本數 n = 83')
    bars27[3].set_color(NAVY) # k=7
    ax.text(7, 135, 'k=7: 128 >= 83\n(建議組數 7 組)', ha='center', va='bottom', fontweight='bold', color=NAVY, fontsize=9.5)
    ax.set_xlabel('分組組數 (k)', fontsize=10)
    ax.set_ylabel('2^k 次方值', fontsize=10)
    ax.set_title('第 2 章 習題 27：2^k 法則決定組數 (k = 7 組)', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xticks(k_vals27)
    ax.legend(frameon=True, facecolor=LIGHT_BLUE)
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch2_q27.png"))
    plt.close()

    # Ch2 Q29: 通勤時間次數分配直方圖
    fig, ax = plt.subplots(figsize=(6, 3.8), dpi=300)
    bins29 = [15, 22, 29, 36, 43, 50]
    counts29 = [3, 8, 7, 5, 2]
    mid29 = [(bins29[i]+bins29[i+1])/2 for i in range(len(counts29))]
    bars29 = ax.bar(mid29, counts29, width=7, color=BLUE, edgecolor='white', linewidth=1.2, alpha=0.9)
    bars29[1].set_color(NAVY)
    for m, c in zip(mid29, counts29):
        ax.text(m, c + 0.2, f'{c}人\n({c/25:.0%})', ha='center', va='bottom', fontsize=9, fontweight='bold')
    ax.set_title('第 2 章 習題 29：員工通勤時間次數分配直方圖 (n = 25)', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('通勤時間（分鐘，組距 7）', fontsize=10)
    ax.set_ylabel('人數 (f)', fontsize=10)
    ax.set_xticks(bins29)
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch2_q29.png"))
    plt.close()

    # Ch2 Q31: 大學生看電影次數鐘形分配直方圖
    fig, ax = plt.subplots(figsize=(6, 3.8), dpi=300)
    bins31 = [0, 2, 4, 6, 8, 10, 12]
    counts31 = [1, 5, 12, 17, 8, 2]
    mid31 = [(bins31[i]+bins31[i+1])/2 for i in range(len(counts31))]
    bars31 = ax.bar(mid31, counts31, width=2, color=BLUE, edgecolor='white', linewidth=1.2, alpha=0.85)
    bars31[3].set_color(NAVY) # 眾數組 6~8
    # 擬合鐘形曲線
    x_curve = np.linspace(0, 12, 100)
    mean_val, std_val = 6.0, 2.2
    y_curve = 45 * 2 * (1 / (std_val * np.sqrt(2 * np.pi))) * np.exp(-0.5 * ((x_curve - mean_val) / std_val) ** 2)
    ax.plot(x_curve, y_curve, color=ACCENT, linewidth=2, linestyle='-', label='對稱鐘形擬合曲線')
    for m, c in zip(mid31, counts31):
        ax.text(m, c + 0.3, str(c), ha='center', va='bottom', fontsize=9.5, fontweight='bold')
    ax.set_title('第 2 章 習題 31：大學生觀影次數鐘形對稱分配直方圖 (n = 45)', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('看電影次數（組距 2）', fontsize=10)
    ax.set_ylabel('學生人數 (f)', fontsize=10)
    ax.set_xticks(bins31)
    ax.legend(frameon=True, facecolor='white')
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch2_q31.png"))
    plt.close()

    # ==================== CHAPTER 3 ====================
    # Ch3 Q43: 5 數值之平均數與離差圖
    fig, ax = plt.subplots(figsize=(5.5, 3.5), dpi=300)
    x_pts = [8, 3, 7, 3, 4]
    y_pts = [1, 2, 3, 4, 5]
    ax.scatter(x_pts, y_pts, color=NAVY, s=100, zorder=5, label='觀測數值 x')
    ax.axvline(5.0, color=ACCENT, linestyle='--', linewidth=2, label='母體平均數 μ = 5.0')
    for x, y in zip(x_pts, y_pts):
        ax.plot([5.0, x], [y, y], color=BLUE, linestyle='-', linewidth=1.5)
        diff = int(x - 5)
        ax.text((5.0 + x)/2, y + 0.15, f'{diff:+d}', ha='center', va='bottom', fontsize=9, fontweight='bold', color=BLUE)
    ax.set_yticks(y_pts)
    ax.set_yticklabels(['值 1: 8', '值 2: 3', '值 3: 7', '值 4: 3', '值 5: 4'])
    ax.set_title('第 3 章 習題 43：各數值對平均數 μ = 5 之離差與變異數分析', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('數值尺度', fontsize=10)
    ax.legend(loc='lower right', frameon=True, facecolor=LIGHT_BLUE)
    ax.grid(True, linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch3_q43.png"))
    plt.close()

    # Ch3 Q53: 10 家旅館房價分布與平均數標準差
    fig, ax = plt.subplots(figsize=(6, 3.6), dpi=300)
    prices = [101, 97, 103, 110, 78, 87, 101, 80, 106, 88]
    ax.scatter(prices, [1]*len(prices), color=BLUE, s=90, alpha=0.8, label='旅館房價樣本 (n=10)')
    ax.axvline(95.1, color=NAVY, linewidth=2, label='樣本平均數 x̄ = $95.1')
    ax.axvspan(95.1 - 11.12, 95.1 + 11.12, color=LIGHT_BLUE, alpha=0.6, label='標準差區間 x̄ ± 1s ($84.0 ~ $106.2)')
    ax.set_title('第 3 章 習題 53：休斯頓汽車旅館平日房價樣本與離散尺度', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('房價 ($)', fontsize=10)
    ax.set_yticks([])
    ax.legend(frameon=True, facecolor='white', loc='upper left')
    ax.grid(axis='x', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch3_q53.png"))
    plt.close()

    # Ch3 Q55: 柴比雪夫定理 k=1.8 比例圖
    fig, ax = plt.subplots(figsize=(6, 3.5), dpi=300)
    x = np.linspace(-3, 3, 200)
    # 任意形狀分配（如雙峰）
    y = 0.5 * np.exp(-0.5 * (x - 1)**2) + 0.5 * np.exp(-0.5 * (x + 1)**2)
    ax.plot(x, y, color=NAVY, linewidth=1.8, label='任意分配曲線 (Arbitrary Distribution)')
    ax.axvline(-1.8, color=ACCENT, linestyle='--', label='μ - 1.8σ')
    ax.axvline(1.8, color=ACCENT, linestyle='--', label='μ + 1.8σ')
    mask = (x >= -1.8) & (x <= 1.8)
    ax.fill_between(x[mask], y[mask], color=LIGHT_BLUE, alpha=0.7, label='柴比雪夫至少涵蓋 >= 69.14%')
    ax.set_title('第 3 章 習題 55：柴比雪夫定理 (k = 1.8，至少 1 - 1/k² ≈ 69%)', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xticks([-1.8, 0, 1.8])
    ax.set_xticklabels(['μ - 1.8σ', 'μ', 'μ + 1.8σ'])
    ax.legend(frameon=True, facecolor='white', fontsize=8.5)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch3_q55.png"))
    plt.close()

    # Ch3 Q57: 經驗法則 68-95-99.7% 與 47.5%、2.5%
    fig, ax = plt.subplots(figsize=(6.5, 3.8), dpi=300)
    x = np.linspace(-3.5, 3.5, 300)
    y = (1 / np.sqrt(2 * np.pi)) * np.exp(-0.5 * x**2)
    ax.plot(x, y, color=NAVY, linewidth=2)
    # 塗色 0 到 2: 47.5%
    m1 = (x >= 0) & (x <= 2)
    ax.fill_between(x[m1], y[m1], color=BLUE, alpha=0.6, label='x̄ 至 x̄ + 2s 占 47.5%')
    # 塗色 > 2: 2.5%
    m2 = (x >= 2)
    ax.fill_between(x[m2], y[m2], color=ACCENT, alpha=0.7, label='高於 x̄ + 2s 占 2.5%')
    # 塗色 -2 到 0: 47.5%
    m3 = (x >= -2) & (x < 0)
    ax.fill_between(x[m3], y[m3], color=LIGHT_BLUE, alpha=0.5, label='x̄ - 2s 至 x̄ 占 47.5%')
    ax.set_title('第 3 章 習題 57：常態鐘形分配經驗法則面積分割', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xticks([-3, -2, -1, 0, 1, 2, 3])
    ax.set_xticklabels(['x̄-3s', 'x̄-2s', 'x̄-1s', 'x̄', 'x̄+1s', 'x̄+2s', 'x̄+3s'])
    ax.legend(frameon=True, facecolor='white', fontsize=8.5)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch3_q57.png"))
    plt.close()

    # Ch3 Q59: 離差和恆為零平衡槓桿圖
    fig, ax = plt.subplots(figsize=(6, 3), dpi=300)
    ax.plot([2, 8], [0, 0], color=SLATE, linewidth=3)
    ax.plot([5, 4.8, 5.2, 5], [0, -0.4, -0.4, 0], color=NAVY, linewidth=2) # 支點
    ax.scatter([3, 4, 5, 6, 7], [0, 0, 0, 0, 0], color=BLUE, s=120, zorder=5)
    ax.text(5, -0.6, '平衡支點：平均數 μ = 5', ha='center', fontweight='bold', color=NAVY, fontsize=9.5)
    ax.text(3, 0.15, '3 (-2)', ha='center', color=ACCENT, fontweight='bold')
    ax.text(4, 0.15, '4 (-1)', ha='center', color=ACCENT, fontweight='bold')
    ax.text(5, 0.15, '5 (0)', ha='center', color=NAVY, fontweight='bold')
    ax.text(6, 0.15, '6 (+1)', ha='center', color=GREEN, fontweight='bold')
    ax.text(7, 0.15, '7 (+2)', ha='center', color=GREEN, fontweight='bold')
    ax.set_xlim(2, 8)
    ax.set_ylim(-0.8, 0.6)
    ax.axis('off')
    ax.set_title('第 3 章 習題 59：算術平均數離差總和恆為零 Σ(x - μ) = 0 槓桿平衡', fontsize=11, fontweight='bold', color=NAVY)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch3_q59.png"))
    plt.close()

    # Ch3 Q65: 加權平均數條圖
    fig, ax = plt.subplots(figsize=(5.5, 3.6), dpi=300)
    grades = ['A 等級 ($5.00)', 'B 等級 ($6.50)', 'C 等級 ($8.00)']
    costs = [5.00, 6.50, 8.00]
    lengths = [270, 300, 100]
    bars = ax.bar(grades, costs, color=BLUE, width=0.55, alpha=0.85)
    ax.axhline(6.12, color=ACCENT, linestyle='--', linewidth=2, label='加權平均每英尺成本 = $6.12')
    for b, l in zip(bars, lengths):
        ax.text(b.get_x() + b.get_width()/2, b.get_height() + 0.15, f'長度 {l} 呎', ha='center', va='bottom', fontsize=9)
    ax.set_title('第 3 章 習題 65：Split-A-Rail 各等級籬笆單價與加權平均成本', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_ylabel('每英尺成本 ($)', fontsize=10)
    ax.set_ylim(0, 9.5)
    ax.legend(frameon=True, facecolor=LIGHT_BLUE)
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch3_q65.png"))
    plt.close()

    # Ch3 Q75: 阿波羅計畫任務時間
    fig, ax = plt.subplots(figsize=(6.5, 3.6), dpi=300)
    missions = [f'M{i+1}' for i in range(13)]
    times = [9, 195, 241, 301, 216, 260, 7, 244, 192, 147, 10, 295, 142]
    ax.plot(missions, times, marker='o', color=NAVY, linewidth=1.5, markersize=5)
    ax.axhline(173.77, color=ACCENT, linestyle='--', label='母體平均數 μ = 173.77 小時')
    ax.axhline(195.0, color=GREEN, linestyle=':', label='中位數 = 195.00 小時')
    ax.set_title('第 3 章 習題 75：阿波羅太空計畫 13 次任務時間 (全距 294 小時)', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_ylabel('任務時間（小時）', fontsize=10)
    ax.legend(frameon=True, facecolor='white', fontsize=8.5)
    ax.grid(True, linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch3_q75.png"))
    plt.close()

    # Ch3 Q77: 大學生娛樂開銷直方圖與 95% 區間
    fig, ax = plt.subplots(figsize=(6.2, 3.8), dpi=300)
    spendings = [684, 710, 688, 711, 722, 698, 723, 743, 738, 722, 696, 721, 685, 763, 681, 731, 736, 771, 693, 701, 737, 717, 752, 710, 697]
    ax.hist(spendings, bins=8, color=BLUE, edgecolor='white', alpha=0.85)
    ax.axvline(717.20, color=NAVY, linewidth=2, label='樣本平均數 x̄ = $717.20')
    ax.axvspan(667.46, 766.94, color=LIGHT_BLUE, alpha=0.5, label='95% 經驗法則區間 [$667.46, $766.94]')
    ax.set_title('第 3 章 習題 77：大學生娛樂開銷樣本分配與 95% 區間', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('開銷金額 ($)', fontsize=10)
    ax.set_ylabel('人數 (f)', fontsize=10)
    ax.legend(frameon=True, facecolor='white', fontsize=8.5)
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch3_q77.png"))
    plt.close()

    # ==================== CHAPTER 4 ====================
    # Ch4 Q5: 數值軸與四分位數位置
    fig, ax = plt.subplots(figsize=(6, 2.8), dpi=300)
    vals45 = [46, 47, 49, 49, 51, 53, 54, 54, 55, 55, 59]
    ax.scatter(vals45, [1]*len(vals45), color=SLATE, s=80, alpha=0.8)
    ax.axvline(49, color=BLUE, linestyle='--', linewidth=1.8, label='Q1 = 49 (25% 門檻)')
    ax.axvline(53, color=NAVY, linestyle='-', linewidth=2.2, label='中位數 = 53 (50% 門檻)')
    ax.axvline(55, color=GREEN, linestyle='--', linewidth=1.8, label='Q3 = 55 (75% 門檻)')
    ax.set_yticks([])
    ax.set_title('第 4 章 習題 5：四分位數 (Q1=49, Median=53, Q3=55) 位置標記', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('觀測數值 (n = 11)', fontsize=10)
    ax.legend(frameon=True, facecolor=LIGHT_BLUE, fontsize=8.5, loc='upper left')
    ax.grid(axis='x', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch4_q5.png"))
    plt.close()

    # Ch4 Q9: 書籍花費箱形圖
    fig, ax = plt.subplots(figsize=(6.5, 3.2), dpi=300)
    # 構建水平箱形圖
    ax.plot([200, 500], [1, 1], color=NAVY, linewidth=1.5)
    ax.plot([1200, 1900], [1, 1], color=NAVY, linewidth=1.5)
    # 箱體
    rect = plt.Rectangle((500, 0.7), 700, 0.6, facecolor=LIGHT_BLUE, edgecolor=NAVY, linewidth=2)
    ax.add_patch(rect)
    ax.plot([800, 800], [0.7, 1.3], color=ACCENT, linewidth=2.5, label='中位數 = $800')
    ax.axvline(2250, color=ACCENT, linestyle=':', linewidth=1.5, label='離群值上限 = $2,250 (Q3 + 1.5×IQR)')
    ax.text(500, 0.6, 'Q1=$500', ha='center', fontsize=9, fontweight='bold')
    ax.text(1200, 0.6, 'Q3=$1,200', ha='center', fontsize=9, fontweight='bold')
    ax.text(200, 0.6, 'Min=200', ha='center', fontsize=8.5)
    ax.text(1900, 0.6, 'Max=1900', ha='center', fontsize=8.5)
    ax.set_xlim(0, 2500)
    ax.set_ylim(0.4, 1.6)
    ax.set_yticks([])
    ax.set_title('第 4 章 習題 9：大學生書籍文具花費箱形圖（正偏，無離群值）', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('每年花費金額 ($)', fontsize=10)
    ax.legend(frameon=True, facecolor='white', loc='upper right', fontsize=8.5)
    ax.grid(axis='x', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch4_q9.png"))
    plt.close()

    # Ch4 Q13: 起薪點圖與偏態
    fig, ax = plt.subplots(figsize=(5.5, 3.2), dpi=300)
    salaries = [26.0, 28.0, 31.0, 33.0, 36.0]
    ax.scatter(salaries, [1]*len(salaries), color=BLUE, s=100, zorder=5)
    ax.axvline(30.8, color=NAVY, linewidth=2, label='平均數 x̄ = 30.8')
    ax.axvline(31.0, color=ACCENT, linestyle='--', linewidth=2, label='中位數 = 31.0')
    ax.set_yticks([])
    ax.set_title('第 4 章 習題 13：會計畢業生起薪分布與偏態 (sk_Pearson = -0.15)', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('起薪（千美元）', fontsize=10)
    ax.legend(frameon=True, facecolor=LIGHT_BLUE, fontsize=8.5)
    ax.grid(axis='x', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch4_q13.png"))
    plt.close()

    # Ch4 Q19: 用餐時間與點甜點分組長條圖
    fig, ax = plt.subplots(figsize=(5.5, 3.8), dpi=300)
    x_pos = np.arange(2)
    width = 0.35
    ax.bar(x_pos - width/2, [32, 85], width, label='點甜點 (是)', color=BLUE)
    ax.bar(x_pos + width/2, [68, 15], width, label='不點甜點 (否)', color='#CBD5E0')
    ax.text(0 - width/2, 34, '32%', ha='center', fontweight='bold', fontsize=9.5)
    ax.text(1 - width/2, 87, '85%', ha='center', fontweight='bold', fontsize=9.5)
    ax.set_xticks(x_pos)
    ax.set_xticklabels(['午餐時段 (Lunch)', '晚餐時段 (Dinner)'])
    ax.set_ylabel('比率 (%)', fontsize=10)
    ax.set_title('第 4 章 習題 19：用餐時段與點甜點列聯關聯長條圖', fontsize=11, fontweight='bold', color=NAVY)
    ax.legend(frameon=True, facecolor='white')
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch4_q19.png"))
    plt.close()

    # Ch4 Q23: 房貸再融資箱形圖
    fig, ax = plt.subplots(figsize=(6.5, 3.2), dpi=300)
    ax.plot([59.2, 68.175], [1, 1], color=NAVY, linewidth=1.5)
    ax.plot([89.425, 100.7], [1, 1], color=NAVY, linewidth=1.5)
    rect = plt.Rectangle((68.175, 0.7), 89.425 - 68.175, 0.6, facecolor=LIGHT_BLUE, edgecolor=NAVY, linewidth=2)
    ax.add_patch(rect)
    ax.plot([84.65, 84.65], [0.7, 1.3], color=ACCENT, linewidth=2.5, label='中位數 = 84.65')
    ax.text(68.175, 0.6, 'Q1=68.18', ha='center', fontsize=8.5, fontweight='bold')
    ax.text(89.425, 0.6, 'Q3=89.43', ha='center', fontsize=8.5, fontweight='bold')
    ax.text(84.65, 1.35, 'Median=84.65', ha='center', fontsize=8.5, color=ACCENT, fontweight='bold')
    ax.set_xlim(55, 105)
    ax.set_ylim(0.4, 1.6)
    ax.set_yticks([])
    ax.set_title('第 4 章 習題 23：房屋貸款再融資金額箱形圖 (千歐元)', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('再融資金額（千歐元）', fontsize=10)
    ax.legend(frameon=True, facecolor='white', loc='upper left', fontsize=8.5)
    ax.grid(axis='x', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch4_q23.png"))
    plt.close()

    # Ch4 Q25: 雙組並列箱形圖 (大眾運輸 vs 自行開車)
    fig, ax = plt.subplots(figsize=(6.5, 3.8), dpi=300)
    transit = [23, 25, 25, 30, 31, 31, 32, 33, 35, 36, 37, 42]
    drive = [32, 32, 33, 34, 37, 37, 38, 38, 38, 39, 40, 44]
    bp = ax.boxplot([transit, drive], patch_artist=True, orientation='horizontal',
                    boxprops=dict(facecolor=LIGHT_BLUE, color=NAVY, linewidth=1.5),
                    medianprops=dict(color=ACCENT, linewidth=2.2))
    ax.set_yticklabels(['搭乘大眾運輸\n(中位數 31.5分, IQR 9.5)', '自行開車\n(中位數 37.5分, IQR 5.5)'])
    ax.set_title('第 4 章 習題 25：上班通勤時間雙組並列箱形圖比較', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('通勤時間（分鐘）', fontsize=10)
    ax.grid(axis='x', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch4_q25.png"))
    plt.close()

    # Ch4 Q31: 汽車失竊數點圖
    fig, ax = plt.subplots(figsize=(5.5, 3.2), dpi=300)
    thefts = [3, 3, 7, 8, 8, 12, 13]
    ax.scatter(thefts, [1]*len(thefts), color=BLUE, s=100, zorder=5)
    ax.axvline(7.71, color=NAVY, linewidth=2, label='平均數 x̄ = 7.71')
    ax.axvline(8.0, color=ACCENT, linestyle='--', linewidth=2, label='中位數 = 8.00')
    ax.set_yticks([])
    ax.set_title('第 4 章 習題 31：7 大城市汽車失竊數分布（接近對稱）', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('汽車失竊數', fontsize=10)
    ax.legend(frameon=True, facecolor=LIGHT_BLUE, fontsize=8.5)
    ax.grid(axis='x', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch4_q31.png"))
    plt.close()

    # Ch4 Q35: 男女失業率對比長條圖
    fig, ax = plt.subplots(figsize=(5.5, 3.6), dpi=300)
    gender = ['男性 (Male)', '女性 (Female)', '全體平均']
    rates = [5.64, 5.12, 5.40]
    bars = ax.bar(gender, rates, color=[NAVY, BLUE, SLATE], width=0.5, alpha=0.85)
    for b in bars:
        ax.text(b.get_x() + b.get_width()/2, b.get_height() + 0.1, f'{b.get_height():.2f}%', ha='center', va='bottom', fontweight='bold', fontsize=9.5)
    ax.set_title('第 4 章 習題 35：美國 20 歲以上性別勞動力之失業率比較', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_ylabel('失業率 (%)', fontsize=10)
    ax.set_ylim(0, 7)
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch4_q35.png"))
    plt.close()

    # ==================== CHAPTER 5 ====================
    # Ch5 Q29: 劍尾魚文氏圖 (Venn Diagram)
    fig, ax = plt.subplots(figsize=(5.5, 3.8), dpi=300)
    circle_green = plt.Circle((0.38, 0.5), 0.32, facecolor='#68D391', edgecolor=GREEN, alpha=0.5, linewidth=2)
    circle_male = plt.Circle((0.62, 0.5), 0.32, facecolor='#63B3ED', edgecolor=BLUE, alpha=0.5, linewidth=2)
    ax.add_patch(circle_green)
    ax.add_patch(circle_male)
    ax.text(0.24, 0.5, '綠色且雌\n44 隻', ha='center', va='center', fontweight='bold', fontsize=9.5)
    ax.text(0.50, 0.5, '綠色且雄\n36 隻', ha='center', va='center', fontweight='bold', fontsize=9.5, color=NAVY)
    ax.text(0.76, 0.5, '橙色且雄\n24 隻', ha='center', va='center', fontweight='bold', fontsize=9.5)
    ax.text(0.50, 0.08, '橙色且雌：36 隻 ｜ 全體總數 N = 140 隻', ha='center', fontsize=9, color=SLATE)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')
    ax.set_title('第 5 章 習題 29：水族箱魚群分類文氏圖 P(A 或 B) = 104/140', fontsize=11, fontweight='bold', color=NAVY)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch5_q29.png"))
    plt.close()

    # Ch5 Q35: 雙向列聯表熱力圖
    fig, ax = plt.subplots(figsize=(5.2, 3.2), dpi=300)
    data35 = np.array([[2, 1, 3], [1, 2, 1]])
    cax = ax.matshow(data35, cmap='Blues', alpha=0.7)
    for i in range(2):
        for j in range(3):
            ax.text(j, i, f'{data35[i, j]} (P={data35[i,j]/10:.1f})', ha='center', va='center', fontweight='bold', fontsize=10)
    ax.set_xticks([0, 1, 2])
    ax.set_xticklabels(['A1', 'A2', 'A3'])
    ax.set_yticks([0, 1])
    ax.set_yticklabels(['B1', 'B2'])
    ax.set_title('第 5 章 習題 35：事件交叉機率分佈網格圖 (總數 10)', fontsize=11, fontweight='bold', color=NAVY, pad=12)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch5_q35.png"))
    plt.close()

    # Ch5 Q41: 貝氏定理樹狀圖
    fig, ax = plt.subplots(figsize=(6.2, 3.6), dpi=300)
    ax.axis('off')
    # 樹狀節點
    ax.text(0.05, 0.5, '起點', ha='center', va='center', bbox=dict(boxstyle='circle', facecolor=LIGHT_BLUE, edgecolor=NAVY))
    ax.text(0.4, 0.75, 'A1 (0.60)', ha='center', va='center', bbox=dict(boxstyle='round', facecolor=LIGHT_BLUE, edgecolor=BLUE))
    ax.text(0.4, 0.25, 'A2 (0.40)', ha='center', va='center', bbox=dict(boxstyle='round', facecolor=LIGHT_BLUE, edgecolor=BLUE))
    ax.text(0.85, 0.85, 'B1 (0.05) → 聯合: 0.030', ha='left', va='center', fontsize=9, fontweight='bold', color=NAVY)
    ax.text(0.85, 0.65, 'B2 (0.95) → 聯合: 0.570', ha='left', va='center', fontsize=9, color=SLATE)
    ax.text(0.85, 0.35, 'B1 (0.10) → 聯合: 0.040', ha='left', va='center', fontsize=9, fontweight='bold', color=NAVY)
    ax.text(0.85, 0.15, 'B2 (0.90) → 聯合: 0.360', ha='left', va='center', fontsize=9, color=SLATE)
    
    # 箭頭
    ax.annotate('', xy=(0.33, 0.72), xytext=(0.1, 0.53), arrowprops=dict(arrowstyle='->', lw=1.5, color=BLUE))
    ax.annotate('', xy=(0.33, 0.28), xytext=(0.1, 0.47), arrowprops=dict(arrowstyle='->', lw=1.5, color=BLUE))
    ax.annotate('', xy=(0.82, 0.85), xytext=(0.48, 0.77), arrowprops=dict(arrowstyle='->', lw=1.2, color=NAVY))
    ax.annotate('', xy=(0.82, 0.65), xytext=(0.48, 0.73), arrowprops=dict(arrowstyle='->', lw=1.2, color=SLATE))
    ax.annotate('', xy=(0.82, 0.35), xytext=(0.48, 0.27), arrowprops=dict(arrowstyle='->', lw=1.2, color=NAVY))
    ax.annotate('', xy=(0.82, 0.15), xytext=(0.48, 0.23), arrowprops=dict(arrowstyle='->', lw=1.2, color=SLATE))
    
    ax.text(0.5, 0.02, '事後機率 P(A1|B1) = 0.030 / (0.030 + 0.040) = 0.4286', ha='center', fontweight='bold', color=NAVY, fontsize=9.5)
    ax.set_title('第 5 章 習題 41：貝氏定理機率樹狀圖推導', fontsize=11, fontweight='bold', color=NAVY)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch5_q41.png"))
    plt.close()

    # Ch5 Q43: 球隊贏球時段貝氏樹狀圖
    fig, ax = plt.subplots(figsize=(6.2, 3.6), dpi=300)
    ax.axis('off')
    ax.text(0.05, 0.5, '全季比賽', ha='center', va='center', bbox=dict(boxstyle='circle', facecolor=LIGHT_BLUE, edgecolor=NAVY))
    ax.text(0.4, 0.75, '夜間 (0.70)', ha='center', va='center', bbox=dict(boxstyle='round', facecolor=LIGHT_BLUE, edgecolor=BLUE))
    ax.text(0.4, 0.25, '白天 (0.30)', ha='center', va='center', bbox=dict(boxstyle='round', facecolor=LIGHT_BLUE, edgecolor=BLUE))
    ax.text(0.85, 0.85, '贏球 (0.50) → 聯合: 0.35', ha='left', va='center', fontsize=9, fontweight='bold', color=NAVY)
    ax.text(0.85, 0.65, '輸球 (0.50) → 聯合: 0.35', ha='left', va='center', fontsize=9, color=SLATE)
    ax.text(0.85, 0.35, '贏球 (0.90) → 聯合: 0.27', ha='left', va='center', fontsize=9, fontweight='bold', color=NAVY)
    ax.text(0.85, 0.15, '輸球 (0.10) → 聯合: 0.03', ha='left', va='center', fontsize=9, color=SLATE)
    
    ax.annotate('', xy=(0.33, 0.72), xytext=(0.1, 0.53), arrowprops=dict(arrowstyle='->', lw=1.5, color=BLUE))
    ax.annotate('', xy=(0.33, 0.28), xytext=(0.1, 0.47), arrowprops=dict(arrowstyle='->', lw=1.5, color=BLUE))
    ax.annotate('', xy=(0.82, 0.85), xytext=(0.48, 0.77), arrowprops=dict(arrowstyle='->', lw=1.2, color=NAVY))
    ax.annotate('', xy=(0.82, 0.65), xytext=(0.48, 0.73), arrowprops=dict(arrowstyle='->', lw=1.2, color=SLATE))
    ax.annotate('', xy=(0.82, 0.35), xytext=(0.48, 0.27), arrowprops=dict(arrowstyle='->', lw=1.2, color=NAVY))
    ax.annotate('', xy=(0.82, 0.15), xytext=(0.48, 0.23), arrowprops=dict(arrowstyle='->', lw=1.2, color=SLATE))
    
    ax.text(0.5, 0.02, '事後機率 P(夜間|贏) = 0.35 / (0.35 + 0.27) = 0.5645', ha='center', fontweight='bold', color=NAVY, fontsize=9.5)
    ax.set_title('第 5 章 習題 43：小聯盟球隊贏球時段貝氏機率樹狀圖', fontsize=11, fontweight='bold', color=NAVY)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch5_q43.png"))
    plt.close()

    # Ch5 Q45: 百貨支付方式貝氏樹狀圖
    fig, ax = plt.subplots(figsize=(6.5, 4.0), dpi=300)
    ax.axis('off')
    ax.text(0.05, 0.5, '付款方式', ha='center', va='center', bbox=dict(boxstyle='circle', facecolor=LIGHT_BLUE, edgecolor=NAVY))
    ax.text(0.35, 0.8, '現金 (0.30)', ha='center', va='center', bbox=dict(boxstyle='round', facecolor=LIGHT_BLUE, edgecolor=BLUE))
    ax.text(0.35, 0.5, '信用卡 (0.30)', ha='center', va='center', bbox=dict(boxstyle='round', facecolor=LIGHT_BLUE, edgecolor=BLUE))
    ax.text(0.35, 0.2, '簽帳卡 (0.40)', ha='center', va='center', bbox=dict(boxstyle='round', facecolor=LIGHT_BLUE, edgecolor=BLUE))
    
    ax.text(0.8, 0.85, '>$50 (0.20) → 0.060', ha='left', va='center', fontsize=9, fontweight='bold', color=NAVY)
    ax.text(0.8, 0.55, '>$50 (0.90) → 0.270', ha='left', va='center', fontsize=9, fontweight='bold', color=NAVY)
    ax.text(0.8, 0.25, '>$50 (0.60) → 0.240', ha='left', va='center', fontsize=9, fontweight='bold', color=NAVY)
    
    ax.annotate('', xy=(0.28, 0.77), xytext=(0.1, 0.55), arrowprops=dict(arrowstyle='->', lw=1.3, color=BLUE))
    ax.annotate('', xy=(0.28, 0.50), xytext=(0.1, 0.50), arrowprops=dict(arrowstyle='->', lw=1.3, color=BLUE))
    ax.annotate('', xy=(0.28, 0.23), xytext=(0.1, 0.45), arrowprops=dict(arrowstyle='->', lw=1.3, color=BLUE))
    
    ax.annotate('', xy=(0.77, 0.85), xytext=(0.43, 0.80), arrowprops=dict(arrowstyle='->', lw=1.2, color=NAVY))
    ax.annotate('', xy=(0.77, 0.55), xytext=(0.43, 0.50), arrowprops=dict(arrowstyle='->', lw=1.2, color=NAVY))
    ax.annotate('', xy=(0.77, 0.25), xytext=(0.43, 0.20), arrowprops=dict(arrowstyle='->', lw=1.2, color=NAVY))
    
    ax.text(0.5, 0.04, '全機率 P(>$50) = 0.570 ｜ P(現金|>$50) = 0.060/0.570 = 0.1053', ha='center', fontweight='bold', color=NAVY, fontsize=9.5)
    ax.set_title('第 5 章 習題 45：百貨公司消費金額與支付方式貝氏樹狀圖', fontsize=11, fontweight='bold', color=NAVY)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch5_q45.png"))
    plt.close()

    # Ch5 Q59: 女籃假日賽機率樹狀圖
    fig, ax = plt.subplots(figsize=(6.2, 3.6), dpi=300)
    ax.axis('off')
    ax.text(0.05, 0.5, '半決賽', ha='center', va='center', bbox=dict(boxstyle='circle', facecolor=LIGHT_BLUE, edgecolor=NAVY))
    ax.text(0.4, 0.75, 'A 隊贏 (2/3)', ha='center', va='center', bbox=dict(boxstyle='round', facecolor=LIGHT_BLUE, edgecolor=BLUE))
    ax.text(0.4, 0.25, 'A 隊輸 (1/3)', ha='center', va='center', bbox=dict(boxstyle='round', facecolor=LIGHT_BLUE, edgecolor=SLATE))
    
    ax.text(0.85, 0.85, 'B 贏 (3/4) → 兩隊皆勝: 1/2 (0.50)', ha='left', va='center', fontsize=8.5, fontweight='bold', color=NAVY)
    ax.text(0.85, 0.65, 'B 輸 (1/4) → 僅 A 勝: 2/12', ha='left', va='center', fontsize=8.5, color=SLATE)
    ax.text(0.85, 0.35, 'B 贏 (3/4) → 僅 B 勝: 3/12', ha='left', va='center', fontsize=8.5, color=SLATE)
    ax.text(0.85, 0.15, 'B 輸 (1/4) → 兩隊皆輸: 1/12 (0.083)', ha='left', va='center', fontsize=8.5, fontweight='bold', color=ACCENT)
    
    ax.annotate('', xy=(0.33, 0.72), xytext=(0.1, 0.53), arrowprops=dict(arrowstyle='->', lw=1.3, color=BLUE))
    ax.annotate('', xy=(0.33, 0.28), xytext=(0.1, 0.47), arrowprops=dict(arrowstyle='->', lw=1.3, color=SLATE))
    ax.annotate('', xy=(0.82, 0.85), xytext=(0.48, 0.77), arrowprops=dict(arrowstyle='->', lw=1.2, color=NAVY))
    ax.annotate('', xy=(0.82, 0.65), xytext=(0.48, 0.73), arrowprops=dict(arrowstyle='->', lw=1.2, color=SLATE))
    ax.annotate('', xy=(0.82, 0.35), xytext=(0.48, 0.27), arrowprops=dict(arrowstyle='->', lw=1.2, color=SLATE))
    ax.annotate('', xy=(0.82, 0.15), xytext=(0.48, 0.23), arrowprops=dict(arrowstyle='->', lw=1.2, color=ACCENT))
    
    ax.text(0.5, 0.02, '至少一隊勝 = 1 - P(兩隊皆輸) = 1 - 1/12 = 11/12 ≈ 0.9167', ha='center', fontweight='bold', color=NAVY, fontsize=9.5)
    ax.set_title('第 5 章 習題 59：女子籃球賽獨立勝負機率樹狀圖', fontsize=11, fontweight='bold', color=NAVY)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch5_q59.png"))
    plt.close()

    # Ch5 Q61: 二項存活機率分布柱狀圖
    fig, ax = plt.subplots(figsize=(5.5, 3.6), dpi=300)
    n61, p61 = 5, 0.98
    k_vals = np.arange(6)
    import math
    probs = [math.comb(5, k) * (p61**k) * ((1-p61)**(5-k)) for k in k_vals]
    bars = ax.bar(k_vals, probs, color=BLUE, width=0.55, alpha=0.85)
    bars[5].set_color(NAVY)
    ax.text(5, probs[5] + 0.03, f'P(5人全活)\n={probs[5]:.4f}', ha='center', va='bottom', fontweight='bold', color=NAVY, fontsize=9)
    ax.set_title('第 5 章 習題 61：5 位男性隔年存活人數機率分布', fontsize=11, fontweight='bold', color=NAVY)
    ax.set_xlabel('存活人數 (k)', fontsize=10)
    ax.set_ylabel('機率 P(X = k)', fontsize=10)
    ax.set_xticks(k_vals)
    ax.set_ylim(0, 1.05)
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch5_q61.png"))
    plt.close()

    # Ch5 Q67: 廣告記憶度機率樹狀圖
    fig, ax = plt.subplots(figsize=(6.2, 3.5), dpi=300)
    ax.axis('off')
    ax.text(0.05, 0.5, '全體客戶', ha='center', va='center', bbox=dict(boxstyle='circle', facecolor=LIGHT_BLUE, edgecolor=NAVY))
    ax.text(0.4, 0.75, '讀報 S (0.60)', ha='center', va='center', bbox=dict(boxstyle='round', facecolor=LIGHT_BLUE, edgecolor=BLUE))
    ax.text(0.4, 0.25, '不讀報 S\' (0.40)', ha='center', va='center', bbox=dict(boxstyle='round', facecolor=LIGHT_BLUE, edgecolor=SLATE))
    
    ax.text(0.85, 0.85, '記得廣告 R (0.85) → 聯合: 0.51 (51%)', ha='left', va='center', fontsize=9, fontweight='bold', color=NAVY)
    ax.text(0.85, 0.65, '不記得 R\' (0.15) → 聯合: 0.09 (9%)', ha='left', va='center', fontsize=9, fontweight='bold', color=ACCENT)
    
    ax.annotate('', xy=(0.33, 0.72), xytext=(0.1, 0.53), arrowprops=dict(arrowstyle='->', lw=1.3, color=BLUE))
    ax.annotate('', xy=(0.33, 0.28), xytext=(0.1, 0.47), arrowprops=dict(arrowstyle='->', lw=1.3, color=SLATE))
    ax.annotate('', xy=(0.82, 0.85), xytext=(0.48, 0.77), arrowprops=dict(arrowstyle='->', lw=1.2, color=NAVY))
    ax.annotate('', xy=(0.82, 0.65), xytext=(0.48, 0.73), arrowprops=dict(arrowstyle='->', lw=1.2, color=ACCENT))
    
    ax.set_title('第 5 章 習題 67：讀報習慣與廣告記憶之條件機率樹狀圖', fontsize=11, fontweight='bold', color=NAVY)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch5_q67.png"))
    plt.close()

    # Ch5 Q69: SD 卡品質結構圖
    fig, ax = plt.subplots(figsize=(5.2, 3.8), dpi=300)
    sizes69 = [857, 112, 31]
    labels69 = ['良品 (Good)\n857張 (85.7%)', '小瑕疵品\n112張 (11.2%)', '完全不能用\n31張 (3.1%)']
    cols69 = [BLUE, AMBER, ACCENT]
    explode69 = (0, 0.08, 0.12)
    ax.pie(sizes69, explode=explode69, labels=labels69, autopct='%1.1f%%', colors=cols69, startangle=140, textprops=dict(fontsize=9))
    ax.set_title('第 5 章 習題 69：1,000 張 SD 卡品質分級比例圓形圖', fontsize=11, fontweight='bold', color=NAVY)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch5_q69.png"))
    plt.close()

    # Ch5 Q71: 服飾店前景評估貝氏樹狀圖
    fig, ax = plt.subplots(figsize=(6.5, 4.0), dpi=300)
    ax.axis('off')
    ax.text(0.05, 0.5, '前景評估', ha='center', va='center', bbox=dict(boxstyle='circle', facecolor=LIGHT_BLUE, edgecolor=NAVY))
    ax.text(0.35, 0.8, '好 (0.60)', ha='center', va='center', bbox=dict(boxstyle='round', facecolor=LIGHT_BLUE, edgecolor=BLUE))
    ax.text(0.35, 0.5, '一般 (0.30)', ha='center', va='center', bbox=dict(boxstyle='round', facecolor=LIGHT_BLUE, edgecolor=BLUE))
    ax.text(0.35, 0.2, '劣 (0.10)', ha='center', va='center', bbox=dict(boxstyle='round', facecolor=LIGHT_BLUE, edgecolor=ACCENT))
    
    ax.text(0.8, 0.85, '獲利 (0.80) → 0.480 (70.6%)', ha='left', va='center', fontsize=9, color=NAVY)
    ax.text(0.8, 0.55, '獲利 (0.60) → 0.180 (26.5%)', ha='left', va='center', fontsize=9, color=NAVY)
    ax.text(0.8, 0.25, '獲利 (0.20) → 0.020 (2.9%)', ha='left', va='center', fontsize=9, fontweight='bold', color=ACCENT)
    
    ax.annotate('', xy=(0.28, 0.77), xytext=(0.1, 0.55), arrowprops=dict(arrowstyle='->', lw=1.3, color=BLUE))
    ax.annotate('', xy=(0.28, 0.50), xytext=(0.1, 0.50), arrowprops=dict(arrowstyle='->', lw=1.3, color=BLUE))
    ax.annotate('', xy=(0.28, 0.23), xytext=(0.1, 0.45), arrowprops=dict(arrowstyle='->', lw=1.3, color=ACCENT))
    
    ax.annotate('', xy=(0.77, 0.85), xytext=(0.43, 0.80), arrowprops=dict(arrowstyle='->', lw=1.2, color=NAVY))
    ax.annotate('', xy=(0.77, 0.55), xytext=(0.43, 0.50), arrowprops=dict(arrowstyle='->', lw=1.2, color=NAVY))
    ax.annotate('', xy=(0.77, 0.25), xytext=(0.43, 0.20), arrowprops=dict(arrowstyle='->', lw=1.2, color=ACCENT))
    
    ax.text(0.5, 0.04, '全機率 P(獲利) = 0.680 ｜ P(劣|獲利) = 0.020/0.680 = 0.0294 (2.94%)', ha='center', fontweight='bold', color=NAVY, fontsize=9.5)
    ax.set_title('第 5 章 習題 71：服飾店發展前景與第一年獲利貝氏樹狀圖', fontsize=11, fontweight='bold', color=NAVY)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch5_q71.png"))
    plt.close()

    # Ch5 Q73: 刮刮樂機率樹狀圖
    fig, ax = plt.subplots(figsize=(6, 3.6), dpi=300)
    visits = np.arange(1, 4)
    p_none = [(0.88)**i for i in visits]
    p_at_least_one = [1 - p for p in p_none]
    ax.plot(visits, p_none, marker='o', color=ACCENT, linewidth=2, label='連去 n 次皆未中獎 P=(0.88)^n')
    ax.plot(visits, p_at_least_one, marker='s', color=GREEN, linewidth=2, label='連去 n 次至少一次中獎 P=1-(0.88)^n')
    for v, pn, pa in zip(visits, p_none, p_at_least_one):
        ax.text(v, pn - 0.06, f'{pn:.4f}', ha='center', fontsize=9, fontweight='bold', color=ACCENT)
        ax.text(v, pa + 0.04, f'{pa:.4f}', ha='center', fontsize=9, fontweight='bold', color=GREEN)
    ax.set_xticks(visits)
    ax.set_xticklabels(['去 1 次', '去 2 次', '連去 3 次'])
    ax.set_ylabel('發生機率', fontsize=10)
    ax.set_ylim(-0.05, 1.05)
    ax.set_title('第 5 章 習題 73：連續訪問披薩店刮刮樂中獎機率趨勢圖', fontsize=11, fontweight='bold', color=NAVY)
    ax.legend(frameon=True, facecolor='white', loc='center left')
    ax.grid(True, linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch5_q73.png"))
    plt.close()

    # Ch5 Q87: 四家供應商瑕疵貝氏事後修正圖
    fig, ax = plt.subplots(figsize=(6.5, 3.8), dpi=300)
    suppliers = ['Tyson\n(供應20%, 瑕疵3%)', 'Fuji\n(供應30%, 瑕疵4%)', 'Kirkpatricks\n(供應25%, 瑕疵7%)', 'Parts\n(供應25%, 瑕疵6.5%)']
    prior = [0.20, 0.30, 0.25, 0.25]
    posterior = [0.1159, 0.2319, 0.3382, 0.3140]
    x_pos = np.arange(4)
    w = 0.35
    ax.bar(x_pos - w/2, [p*100 for p in prior], width=w, color=SLATE, label='事前供應比率 (%)', alpha=0.85)
    bars_post = ax.bar(x_pos + w/2, [p*100 for p in posterior], width=w, color=NAVY, label='事後瑕疵責任比率 (%)', alpha=0.9)
    bars_post[0].set_color(GREEN) # Tyson 瑕疵率最低
    for i in range(4):
        ax.text(x_pos[i] + w/2, posterior[i]*100 + 1, f'{posterior[i]:.1%}', ha='center', fontweight='bold', fontsize=8.5, color=NAVY)
    ax.set_xticks(x_pos)
    ax.set_xticklabels(suppliers, fontsize=8.5)
    ax.set_ylabel('百分比 (%)', fontsize=10)
    ax.set_title('第 5 章 習題 87：四家映像管供應商事前供應 vs 貝氏事後瑕疵機率', fontsize=11, fontweight='bold', color=NAVY)
    ax.legend(frameon=True, facecolor=LIGHT_BLUE)
    ax.grid(axis='y', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "ch5_q87.png"))
    plt.close()

    print("All 40 charts successfully generated!")

if __name__ == "__main__":
    out_dir = r"c:\Users\user\OneDrive\桌面\士鈞的資料夾\統計學期中考\figures"
    generate_all_figures(out_dir)
