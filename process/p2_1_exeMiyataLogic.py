import pandas as pd
import numpy as np
from typing import Dict, List, Optional

def exe_miyata_logic(df_journal: pd.DataFrame, df_bs: pd.DataFrame) -> pd.DataFrame:
    """
    標準化された会計データから、宮田ロジックに基づく11項目の分析を行う。
    """
    results = []
    
    # 日付変換の確認
    df_j = df_journal.copy()
    df_j['date'] = pd.to_datetime(df_j['date'], errors='coerce')
    df_j = df_j.dropna(subset=['date'])
    
    # 期間の把握
    if not df_j.empty:
        min_date = df_j['date'].min()
        max_date = df_j['date'].max()
        months_count = (max_date.year - min_date.year) * 12 + (max_date.month - min_date.month) + 1
    else:
        months_count = 0

    # --- 共通の集計処理 ---
    # 月次売上・原価の集計
    df_j['year_month'] = df_j['date'].dt.to_period('M')
    
    sales_patterns = ["売上", "売上高"]
    cogs_patterns = ["仕入", "売上原価", "外注費"]
    ap_patterns = ["買掛金", "未払金", "未払費用"]

    def is_match(acc_name, patterns):
        return any(p in str(acc_name) for p in patterns)

    df_j['is_sales'] = df_j['debit_account'].apply(lambda x: is_match(x, sales_patterns)) | \
                       df_j['credit_account'].apply(lambda x: is_match(x, sales_patterns))
    df_j['is_cogs'] = df_j['debit_account'].apply(lambda x: is_match(x, cogs_patterns)) | \
                      df_j['credit_account'].apply(lambda x: is_match(x, cogs_patterns))
    df_j['is_ap'] = df_j['debit_account'].apply(lambda x: is_match(x, ap_patterns)) | \
                    df_j['credit_account'].apply(lambda x: is_match(x, ap_patterns))

    # 金額の整理 (売上は貸方、原価は借方が基本だが、標準化時にdebit_amount/credit_amountに分かれている前提)
    # 本来は勘定科目ごとにどちらの金額を取るか決めるべきだが、簡易的に売上科目がある行のdebit/credit合計を考える
    # ここでは「金額」列が標準化されていることを期待（現状はdebit_amount, credit_amount）
    df_j['sales_amt'] = df_j.apply(lambda r: (r['credit_amount'] if pd.notna(r['credit_amount']) else 0) if r['is_sales'] else 0, axis=1)
    df_j['cogs_amt'] = df_j.apply(lambda r: (r['debit_amount'] if pd.notna(r['debit_amount']) else 0) if r['is_cogs'] else 0, axis=1)
    df_j['ap_amt'] = df_j.apply(lambda r: (r['credit_amount'] if pd.notna(r['credit_amount']) else 0) if r['is_ap'] else 0, axis=1)

    monthly_stats = df_j.groupby('year_month').agg({
        'sales_amt': 'sum',
        'cogs_amt': 'sum',
        'ap_amt': 'sum'
    })
    
    annual_sales = monthly_stats['sales_amt'].sum()

    # --- 共通の集計処理 (追加) ---
    # 前年度と当年度の分離
    if months_count >= 13:
        # ちょうど1年前の日付を境界にする
        boundary_date = max_date - pd.DateOffset(years=1)
        df_curr = df_j[df_j['date'] > boundary_date].copy()
        df_prev = df_j[df_j['date'] <= boundary_date].copy()
    else:
        df_curr = df_j.copy()
        df_prev = pd.DataFrame()

    # --- 1. 資金繰り ---
    # 1.1 現金薄さ
    if not df_bs.empty and "期末現預金合計" in df_bs.columns and annual_sales > 0:
        cash_balance = df_bs["期末現預金合計"].iloc[0]
        ratio = (cash_balance / annual_sales) * 100
        results.append(["① 資金繰り", "現金薄さ", f"{ratio:.1f}%", f"年商に対する現預金比率が{ratio:.1f}%です（目安3%以上）", "red" if ratio < 3 else "blue"])
    else:
        results.append(["① 資金繰り", "現金薄さ", "なし", "貸借対照表がない、または売上が0のため判定できません", "grey"])

    # 1.2 買掛・未払残高
    if len(monthly_stats) >= 3:
        ap_trend = monthly_stats['ap_amt'].iloc[-3:]
        is_increasing = ap_trend.iloc[0] < ap_trend.iloc[1] < ap_trend.iloc[2]
        results.append(["① 資金繰り", "買掛・未払残高", "確認" if is_increasing else "安定", "買掛・未払金が3ヶ月連続で増加しています" if is_increasing else "急激な増加は見られません", "red" if is_increasing else "blue"])
    else:
        results.append(["① 資金繰り", "買掛・未払残高", "なし", "データが3ヶ月分に満たないため判定できません", "grey"])

    # --- 2. 会計品質 ---
    # 2.1 仕訳入力遅延
    if 'created_at' in df_j.columns and df_j['created_at'].notna().any():
        df_j['created_at_dt'] = pd.to_datetime(df_j['created_at'], errors='coerce')
        # dateをdatetimeに変換 (created_at_dtは既にdatetime)
        df_j['date_dt'] = pd.to_datetime(df_j['date'])
        df_j['delay'] = (df_j['created_at_dt'] - df_j['date_dt']).dt.days
        delayed_count = (df_j['delay'] > 15).sum()
        total_count = df_j['created_at_dt'].notna().sum()
        delay_rate = (delayed_count / total_count) * 100 if total_count > 0 else 0
        results.append(["② 会計品質", "仕訳入力遅延", f"{delay_rate:.1f}%", f"15日以上の入力遅延が{delay_rate:.1f}%発生しています", "red" if delay_rate >= 20 else "blue"])
    else:
        results.append(["② 会計品質", "仕訳入力遅延", "なし", "CSVに「作成日（登録日）」列がないため判定できません", "grey"])

    # 2.2 粗利率ブレ
    if len(monthly_stats) >= 2:
        # 分母が0の場合の回避
        with np.errstate(divide='ignore', invalid='ignore'):
            monthly_stats['margin'] = (monthly_stats['sales_amt'] - monthly_stats['cogs_amt']) / monthly_stats['sales_amt']
            monthly_stats['margin'] = monthly_stats['margin'].replace([np.inf, -np.inf], np.nan).fillna(0)
        
        monthly_stats['margin_diff'] = monthly_stats['margin'].diff().abs()
        is_volatile = (monthly_stats['margin_diff'] > 0.1).any()
        results.append(["② 会計品質", "粗利率ブレ", "変動あり" if is_volatile else "安定", "月次の粗利率に10%以上の変動が見られます" if is_volatile else "安定した粗利率で推移しています", "red" if is_volatile else "blue"])
    else:
        results.append(["② 会計品質", "粗利率ブレ", "なし", "データが2ヶ月分に満たないため判定できません", "grey"])

    # 2.3 入金サイト延伸
    if months_count >= 13 and not df_prev.empty:
        ar_patterns = ["売掛金"]
        df_j['is_ar'] = df_j['debit_account'].apply(lambda x: is_match(x, ar_patterns)) | \
                       df_j['credit_account'].apply(lambda x: is_match(x, ar_patterns))
        
        def calc_ar_days(target_df):
            sales = target_df.apply(lambda r: (r['credit_amount'] if pd.notna(r['credit_amount']) else 0) if is_match(r['credit_account'], sales_patterns) else 0, axis=1).sum()
            ar_debits = target_df.apply(lambda r: (r['debit_amount'] if pd.notna(r['debit_amount']) else 0) if is_match(r['debit_account'], ar_patterns) else 0, axis=1).sum()
            return (ar_debits / sales * 365) if sales > 0 else 0

        ar_days_curr = calc_ar_days(df_curr)
        ar_days_prev = calc_ar_days(df_prev)
        diff = ar_days_curr - ar_days_prev

        results.append(["② 会計品質", "入金サイト延伸", f"{diff:+.1f}日", f"回収期間が前年比で{diff:+.1f}日変動しています", "red" if diff >= 5 else "blue"])
    else:
        results.append(["② 会計品質", "入金サイト延伸", "なし", "データが12ヶ月分のみのため判定できません", "grey"])

    # --- 3. 売上構造 ---
    # 3.1 新規取引先数
    if months_count >= 13 and not df_prev.empty:
        # 取引先列が credit_partner 側にある前提（売上行）
        partners_curr = set(df_curr[df_curr['is_sales']]['credit_partner'].dropna())
        partners_prev = set(df_prev[df_prev['is_sales']]['credit_partner'].dropna())
        new_partners = partners_curr - partners_prev
        new_partner_count = len(new_partners)
        
        results.append(["③ 売上構造", "新規取引先数", f"{new_partner_count}社", f"直近1年で{new_partner_count}社の新規取引先がありました", "red" if new_partner_count == 0 else "blue"])
    else:
        results.append(["③ 売上構造", "新規取引先数", "なし", "比較対象となる昨年のデータがないため判定できません", "grey"])

    # 3.2 新規継続率
    if months_count >= 13 and 'new_partners' in locals() and len(new_partners) > 0:
        retain_count = 0
        for p in new_partners:
            p_rows = df_curr[df_curr['credit_partner'] == p].sort_values('date')
            if len(p_rows) > 1:
                first_date = p_rows.iloc[0]['date']
                second_date = p_rows.iloc[1]['date']
                if (second_date - first_date).days <= 90:
                    retain_count += 1
        
        retention_rate = (retain_count / len(new_partners)) * 100
        results.append(["③ 売上構造", "新規継続率", f"{retention_rate:.1f}%", f"新規取引先のうち{retention_rate:.1f}%が3ヶ月以内に再取引しています", "red" if retention_rate < 20 else "blue"])
    else:
        results.append(["③ 売上構造", "新規継続率", "なし", "新規取引先がいない、または12ヶ月分のみのため判定できません", "grey"])

    # 3.3 粗利率トレンド
    if len(monthly_stats) >= 3:
        margin_trend = monthly_stats['margin'].iloc[-3:]
        is_declining = margin_trend.iloc[0] > margin_trend.iloc[1] > margin_trend.iloc[2]
        results.append(["③ 売上構造", "粗利率トレンド", "低下中" if is_declining else "安定", "粗利率が3ヶ月連続で低下しています" if is_declining else "粗利率の継続的な低下は見られません", "red" if is_declining else "blue"])
    else:
        results.append(["③ 売上構造", "粗利率トレンド", "なし", "データが3ヶ月分に満たないため判定できません", "grey"])

    # 3.4 上位3社売上集中度
    sales_by_partner = df_j[df_j['is_sales']].groupby('credit_partner')['sales_amt'].sum().sort_values(ascending=False)
    if not sales_by_partner.empty:
        top3_share = (sales_by_partner.head(3).sum() / sales_by_partner.sum()) * 100
        results.append(["③ 売上構造", "上位3社売上集中度", f"{top3_share:.1f}%", f"上位3社への売上集中度が{top3_share:.1f}%です", "red" if top3_share >= 70 else "blue"])
    else:
        results.append(["③ 売上構造", "上位3社売上集中度", "なし", "取引先別の売上データがありません", "grey"])

    # --- 4. 仕入コスト ---
    # 4.1 上位3社仕入集中度
    cogs_by_partner = df_j[df_j['is_cogs']].groupby('debit_partner')['cogs_amt'].sum().sort_values(ascending=False)
    if not cogs_by_partner.empty:
        top3_cogs_share = (cogs_by_partner.head(3).sum() / cogs_by_partner.sum()) * 100
        results.append(["④ 仕入コスト", "上位3社仕入集中度", f"{top3_cogs_share:.1f}%", f"上位3社への仕入集中度が{top3_cogs_share:.1f}%です", "red" if top3_cogs_share >= 70 else "blue"])
    else:
        results.append(["④ 仕入コスト", "上位3社仕入集中度", "なし", "取引先別の仕入データがありません", "grey"])

    # 4.2 単価上昇率
    if months_count >= 13 and 'quantity' in df_j.columns and df_j['quantity'].notna().any():
        # 仕入行の単価（amount / quantity）を計算
        df_curr['unit_price'] = df_curr['cogs_amt'] / df_curr['quantity'].replace(0, np.nan)
        df_prev['unit_price'] = df_prev['cogs_amt'] / df_prev['quantity'].replace(0, np.nan)
        
        avg_price_curr = df_curr['unit_price'].mean()
        avg_price_prev = df_prev['unit_price'].mean()
        
        if avg_price_prev > 0:
            price_increase = (avg_price_curr / avg_price_prev - 1) * 100
            results.append(["④ 仕入コスト", "単価上昇率", f"{price_increase:+.1f}%", f"仕入平均単価が前年比で{price_increase:+.1f}%変動しています", "red" if price_increase >= 10 else "blue"])
        else:
            results.append(["④ 仕入コスト", "単価上昇率", "判定不可", "前年の価格データが不足しています", "grey"])
    else:
        results.append(["④ 仕入コスト", "単価上昇率", "なし", "数量データ（quantity列）がないため判定できません", "grey"])

    # DataFrame化して返却
    return pd.DataFrame(results, columns=["category", "item", "result", "comment", "color"])
