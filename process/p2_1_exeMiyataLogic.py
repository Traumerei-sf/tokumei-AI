import pandas as pd
import numpy as np
from typing import Dict, List, Optional

def exe_miyata_logic(df_journal: pd.DataFrame, df_bs: pd.DataFrame, sales_index_data: Optional[Dict] = None) -> pd.DataFrame:
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
    
    with np.errstate(divide='ignore', invalid='ignore'):
        monthly_stats['margin'] = (monthly_stats['sales_amt'] - monthly_stats['cogs_amt']) / monthly_stats['sales_amt']
        monthly_stats['margin'] = monthly_stats['margin'].replace([np.inf, -np.inf], np.nan).fillna(0)
        
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
        if ratio < 3:
            color = "red"
            comment = f"年商に対する現預金比率が{ratio:.1f}%と、危険水準（3%未満）にあります。手元の資金が極めて薄く、至急の資金手当てが必要です。"
        elif ratio > 10:
            color = "blue"
            comment = f"年商に対する現預金比率が{ratio:.1f}%と、安全水準（10%超）を維持しています。十分なキャッシュが確保されています。"
        else:
            color = "yellow"
            comment = f"年商に対する現預金比率が{ratio:.1f}%と、注意水準（3%以上10%以下）です。不測の事態に備え、手元資金の積み増しを推奨します。"
        results.append(["① 資金繰り", "現金薄さ", f"{ratio:.1f}%", comment, color])
    else:
        results.append(["① 資金繰り", "現金薄さ", "なし", "貸借対照表がない、または売上が0のため判定できません", "grey"])

    # 1.2 買掛・未払残高
    if len(monthly_stats) >= 3:
        ap_trend = monthly_stats['ap_amt'].iloc[-3:]
        v1, v2, v3 = ap_trend.iloc[0], ap_trend.iloc[1], ap_trend.iloc[2]
        if v1 < v2 < v3:
            color = "red"
            res = "3ヶ月連続増加"
            comment = "買掛・未払金残高が3ヶ月連続で増加しています。資金繰りの悪化や支払遅延の予兆である可能性があり、早期の確認が必要です。"
        elif v2 >= v3:
            color = "blue"
            res = "増加なし"
            comment = "買掛・未払金残高の継続的な増加は見られず、安定して推移しています。"
        else:
            color = "yellow"
            res = "単月増加"
            comment = "買掛・未払金残高が前月比で増加しています。一時的な仕入増の可能性もありますが、連続増加に移行しないか注視が必要です。"
        results.append(["① 資金繰り", "買掛・未払残高", res, comment, color])
    else:
        results.append(["① 資金繰り", "買掛・未払残高", "なし", "データが3ヶ月分に満たないため判定できません", "grey"])

    # 1.4 預金体力推移
    if not df_bs.empty and "期末現預金合計" in df_bs.columns:
        results.append(["① 資金繰り", "預金体力推移", "-", "リストに記載", "white"])
    else:
        results.append(["① 資金繰り", "預金体力推移", "なし", "貸借対照表がないため判定できません", "grey"])

    # 1.3 税金・社会保険料 of 納付確認
    # フィルター判定関数の定義
    def is_gensen_credit(r):
        is_acc = pd.notna(r['credit_account']) and "預り金" in str(r['credit_account'])
        is_desc = pd.notna(r['partner']) and any(k in str(r['partner']) for k in ["源泉", "所得税", "給与"])
        return is_acc and is_desc

    def is_gensen_debit(r):
        is_acc = pd.notna(r['debit_account']) and "預り金" in str(r['debit_account'])
        is_desc = pd.notna(r['partner']) and any(k in str(r['partner']) for k in ["源泉", "所得税"])
        return is_acc and is_desc

    def is_juumin_credit(r):
        is_acc = pd.notna(r['credit_account']) and "預り金" in str(r['credit_account'])
        is_desc = pd.notna(r['partner']) and any(k in str(r['partner']) for k in ["住民税", "特別徴収", "市県民税"])
        return is_acc and is_desc

    def is_juumin_debit(r):
        is_acc = pd.notna(r['debit_account']) and "預り金" in str(r['debit_account'])
        is_desc = pd.notna(r['partner']) and any(k in str(r['partner']) for k in ["住民税", "特別徴収"])
        return is_acc and is_desc

    def is_shaho_credit(r):
        is_acc = pd.notna(r['credit_account']) and any(k in str(r['credit_account']) for k in ["預り金", "法定福利費"])
        is_desc = pd.notna(r['partner']) and any(k in str(r['partner']) for k in ["社会保険", "健康保険", "厚生年金", "年金"])
        return is_acc and is_desc

    def is_shaho_debit(r):
        is_acc = pd.notna(r['debit_account']) and any(k in str(r['debit_account']) for k in ["預り金", "法定福利費"])
        is_desc = pd.notna(r['partner']) and any(k in str(r['partner']) for k in ["健康保険", "厚生年金", "日本年金機構", "社会保険"])
        return is_acc and is_desc

    # 全期間で預り金系が全くないかチェック
    has_gensen_any = df_curr.apply(is_gensen_credit, axis=1).any()
    has_juumin_any = df_curr.apply(is_juumin_credit, axis=1).any()
    has_shaho_any = df_curr.apply(is_shaho_credit, axis=1).any()

    if has_gensen_any or has_juumin_any or has_shaho_any:
        # 時系列順の月リスト
        months = sorted(df_curr['year_month'].unique())
        
        unconfirmed_months = 0
        monthly_retained = [] # 各月の滞留フラグ (is_retained, R_next)
        
        for idx, m in enumerate(months):
            df_m = df_curr[df_curr['year_month'] == m]
            
            # 各科目の発生有無
            credit_gensen = df_m.apply(is_gensen_credit, axis=1).any()
            credit_juumin = df_m.apply(is_juumin_credit, axis=1).any()
            credit_shaho = df_m.apply(is_shaho_credit, axis=1).any()
            
            # 翌月15日、翌々月5日の期限算出
            limit_15 = m.start_time + pd.DateOffset(months=1, days=14)
            limit_next_5 = m.start_time + pd.DateOffset(months=2, days=4)
            
            month_unconfirmed = False
            
            # 1. 源泉所得税のチェック
            if credit_gensen:
                # 当月〜翌月15日までの出金仕訳があるか
                df_gensen_debit = df_curr[df_curr.apply(is_gensen_debit, axis=1)]
                has_pay = not df_gensen_debit[(df_gensen_debit['date'] >= m.start_time) & (df_gensen_debit['date'] <= limit_15)].empty
                if not has_pay:
                    month_unconfirmed = True
                    
            # 2. 住民税のチェック
            if credit_juumin:
                df_juumin_debit = df_curr[df_curr.apply(is_juumin_debit, axis=1)]
                has_pay = not df_juumin_debit[(df_juumin_debit['date'] >= m.start_time) & (df_juumin_debit['date'] <= limit_15)].empty
                if not has_pay:
                    month_unconfirmed = True
                    
            # 3. 社会保険料のチェック
            if credit_shaho:
                df_shaho_debit = df_curr[df_curr.apply(is_shaho_debit, axis=1)]
                has_pay = not df_shaho_debit[(df_shaho_debit['date'] >= m.start_time) & (df_shaho_debit['date'] <= limit_next_5)].empty
                if not has_pay:
                    month_unconfirmed = True
                    
            # いずれかの科目の支払チェック対象（発生あり）があり、支払が未確認なら未確認月としてカウント
            if (credit_gensen or credit_juumin or credit_shaho) and month_unconfirmed:
                unconfirmed_months += 1
                
            # 4. 滞留確認（2ヶ月目以降）
            is_retained = False
            R_next = 0.0
            if idx > 0 and (idx < len(months) - 1): # 翌月のデータが存在する月のみ
                next_m = months[idx + 1]
                df_two_months = df_curr[(df_curr['year_month'] == m) | (df_curr['year_month'] == next_m)]
                
                # 当月の天引き額 (A_m)
                A_m = df_m.apply(lambda r: r['credit_amount'] if pd.notna(r['credit_amount']) and (is_gensen_credit(r) or is_juumin_credit(r) or is_shaho_credit(r)) else 0.0, axis=1).sum()
                
                # 当月および翌月の発生額合計
                total_credit = df_two_months.apply(lambda r: r['credit_amount'] if pd.notna(r['credit_amount']) and (is_gensen_credit(r) or is_juumin_credit(r) or is_shaho_credit(r)) else 0.0, axis=1).sum()
                
                # 当月および翌月の支払額合計
                total_debit = df_two_months.apply(lambda r: r['debit_amount'] if pd.notna(r['debit_amount']) and (is_gensen_debit(r) or is_juumin_debit(r) or is_shaho_debit(r)) else 0.0, axis=1).sum()
                
                R_next = total_credit - total_debit
                if R_next < 0.0:
                    R_next = 0.0
                    
                if A_m > 0.0:
                    is_retained = R_next > (A_m * 1.05)
                    
            monthly_retained.append((is_retained, R_next))
            
        # 連続滞留判定（2ヶ月以上連続）
        has_continuous_ret = False
        for i_idx in range(len(monthly_retained) - 1):
            if monthly_retained[i_idx][0] and monthly_retained[i_idx+1][0]:
                has_continuous_ret = True
                break
                
        # 増加傾向判定（3ヶ月連続で5%以上増加）
        has_increasing_trend = False
        for i_idx in range(len(monthly_retained) - 2):
            v1 = monthly_retained[i_idx][1]
            v2 = monthly_retained[i_idx+1][1]
            v3 = monthly_retained[i_idx+2][1]
            if v1 > 1.0 and v2 >= v1 * 1.05 and v3 >= v2 * 1.05:
                has_increasing_trend = True
                break
                
        # 滞留月が1ヶ月でもあるか（単発滞留）
        has_any_ret = any(item[0] for item in monthly_retained)
        
        # 判定決定
        if unconfirmed_months >= 5 or has_continuous_ret or has_increasing_trend:
            color = "red"
            res = "規律再設計期"
            comment = "納付確認が複数月で確認できず、資金規律の再整理が必要となる可能性があります。"
        elif unconfirmed_months >= 3 or has_any_ret:
            color = "yellow"
            res = "規律一部変動"
            comment = "一部月において納付確認が遅延する傾向が見られます。資金規律の整理余地があります。"
        else:
            color = "blue"
            res = "資金規律安定"
            comment = "税金および社会保険料は概ね期日通り納付されています。"
            
        results.append(["① 資金繰り", "税金・社会保険料の納付確認", res, comment, color])

    # --- 2. 会計品質 ---
    # 2.1 仕訳入力遅延
    if 'created_at' in df_j.columns and df_j['created_at'].notna().any():
        df_j['created_at_dt'] = pd.to_datetime(df_j['created_at'], errors='coerce')
        df_j['date_dt'] = pd.to_datetime(df_j['date'])
        df_j['delay'] = (df_j['created_at_dt'] - df_j['date_dt']).dt.days
        delayed_count = (df_j['delay'] > 15).sum()
        total_count = df_j['created_at_dt'].notna().sum()
        delay_rate = (delayed_count / total_count) * 100 if total_count > 0 else 0
        if delay_rate >= 20:
            color = "red"
            comment = f"15日以上の仕訳入力遅延が{delay_rate:.1f}%発生しています。月次決算の早期化や経営状況のリアルタイム把握に支障が出ているため、業務フローの見直しが必要です。"
        elif delay_rate <= 10:
            color = "blue"
            comment = f"15日以上の仕訳入力遅延は{delay_rate:.1f}%と低水準に抑えられており、迅速な記帳が行われています。"
        else:
            color = "yellow"
            comment = f"15日以上の仕訳入力遅延が{delay_rate:.1f}%発生しています。一定の遅延が見られるため、タイムリーな記帳体制への改善が望まれます。"
        results.append(["② 会計品質", "仕訳入力遅延", f"{delay_rate:.1f}%", comment, color])
    else:
        results.append(["② 会計品質", "仕訳入力遅延", "なし", "CSVに「作成日（登録日）」列がないため判定できません", "grey"])

    # 2.2 粗利率ブレ
    if len(monthly_stats) >= 2:
        # 粗利率ブレ用に「未収入金」を含めた売上高を計算する
        sales_patterns_margin = ["売上", "売上高", "未収入金"]
        df_j['is_sales_margin'] = df_j['debit_account'].apply(lambda x: is_match(x, sales_patterns_margin)) | \
                                  df_j['credit_account'].apply(lambda x: is_match(x, sales_patterns_margin))
        df_j['sales_amt_margin'] = df_j.apply(lambda r: (r['credit_amount'] if pd.notna(r['credit_amount']) else 0) if r['is_sales_margin'] else 0, axis=1)
        
        monthly_stats_margin = df_j.groupby('year_month').agg({
            'sales_amt_margin': 'sum',
            'cogs_amt': 'sum'
        })
        
        with np.errstate(divide='ignore', invalid='ignore'):
            monthly_stats_margin['margin'] = (monthly_stats_margin['sales_amt_margin'] - monthly_stats_margin['cogs_amt']) / monthly_stats_margin['sales_amt_margin']
            monthly_stats_margin['margin'] = monthly_stats_margin['margin'].replace([np.inf, -np.inf], np.nan).fillna(0)
        
        monthly_stats_margin['margin_diff'] = monthly_stats_margin['margin'].diff().abs()
        valid_diffs = monthly_stats_margin['margin_diff'].dropna()
        if not valid_diffs.empty:
            max_diff = valid_diffs.max()
            if max_diff >= 0.10:
                color = "red"
                res = "大幅な変動あり"
                comment = f"月次の粗利率に10%以上の大幅な変動（最大 {max_diff*100:.1f}%）が見られます。原価計算のズレや期末の一括調整、または不安定な価格交渉が発生している可能性があります。"
            elif max_diff <= 0.05:
                color = "blue"
                res = "安定"
                comment = f"月次の粗利率変動はすべて5%以内（最大 {max_diff*100:.1f}%）に収まっており、極めて安定しています。"
            else:
                color = "yellow"
                res = "中程度の変動あり"
                comment = f"月次の粗利率に5%以上10%未満の変動（最大 {max_diff*100:.1f}%）が見られます。大きなブレではありませんが、収益性の安定に向けて月次の価格・原価管理を注視してください。"
            results.append(["② 会計品質", "粗利率ブレ", res, comment, color])
        else:
            results.append(["② 会計品質", "粗利率ブレ", "なし", "有効な月次変動データが得られなかったため判定できません", "grey"])
    else:
        results.append(["② 会計品質", "粗利率ブレ", "なし", "データが2ヶ月分に満たないため判定できません", "grey"])

    # 2.3 入金サイト延伸
    if months_count >= 13 and not df_prev.empty:
        ar_patterns = ["売掛金", "未収入金"]
        df_j['is_ar'] = df_j['debit_account'].apply(lambda x: is_match(x, ar_patterns)) | \
                       df_j['credit_account'].apply(lambda x: is_match(x, ar_patterns))
        
        def calc_ar_days(target_df):
            valid_df = target_df[~target_df['partner'].astype(str).str.contains('期首|開始|繰越', na=False)]
            sales_patterns_ar = ["売上", "売上高", "未収入金"]
            sales = valid_df.apply(lambda r: (r['credit_amount'] if pd.notna(r['credit_amount']) else 0) if is_match(r['credit_account'], sales_patterns_ar) else 0, axis=1).sum()
            ar_debits = valid_df.apply(lambda r: (r['debit_amount'] if pd.notna(r['debit_amount']) else 0) if is_match(r['debit_account'], ar_patterns) else 0, axis=1).sum()
            
            if sales <= 1000:
                return 0
            
            days = (ar_debits / sales * 365)
            return min(days, 999.0)

        ar_days_curr = calc_ar_days(df_curr)
        ar_days_prev = calc_ar_days(df_prev)
        
        if ar_days_curr == 0 and ar_days_prev == 0:
            results.append(["② 会計品質", "入金サイト延伸", "なし", "比較可能な有意義な売上データが不足しているため判定できません", "grey"])
        else:
            diff = ar_days_curr - ar_days_prev
            if abs(diff) >= 365:
                results.append(["② 会計品質", "入金サイト延伸", "なし", "存在しない、または解析エラーです", "grey"])
            else:
                if diff >= 5:
                    color = "red"
                    comment = f"回収期間（売掛金回転日数）が前年比で{diff:+.1f}日と、5日以上延伸しています。支払期限の遅延や回収条件の悪化が発生している懸念があり、早期の回収状況確認が必要です。"
                elif diff <= 0:
                    color = "blue"
                    comment = f"回収期間が前年比で{diff:+.1f}日と、維持または短縮されています（0日以下）。回収業務は健全に行われています。"
                else:
                    color = "yellow"
                    comment = f"回収期間が前年比で{diff:+.1f}日と、わずかに延伸しています（0日超5日未満）。大きな変化ではありませんが、売掛金の滞留がないか定期的なチェックをお勧めします。"
                results.append(["② 会計品質", "入金サイト延伸", f"{diff:+.1f}日", comment, color])
    else:
        results.append(["② 会計品質", "入金サイト延伸", "なし", "データが12ヶ月分のみのため判定できません", "grey"])

    # 2.4 売上計上思想指数
    if sales_index_data is not None:
        idx_val = sales_index_data.get("index", 0.0)
        target_cnt = sales_index_data.get("target_count", 0)
        total_cnt = sales_index_data.get("total_count", 0)
        
        if idx_val >= 20.0:
            color = "red"
            comment = f"売上仕訳の中に概算や修正等の曖昧な仕訳（計{target_cnt}件）が全体の{idx_val:.1f}%を占めており、売上計上思想の健全性に重大な疑義があります。"
        elif idx_val >= 5.0:
            color = "yellow"
            comment = f"売上仕訳の中に一部曖昧な仕訳（計{target_cnt}件、比率{idx_val:.1f}%）が確認されます。売上計上ルールの厳格化をお勧めします。"
        else:
            color = "blue"
            comment = f"概算や修正等の仕訳（計{target_cnt}件、比率{idx_val:.1f}%）はごく低水準であり、適切な売上計上が行われています。"
            
        results.append(["② 会計品質", "売上計上思想指数", f"{idx_val:.1f}%", comment, color])
    else:
        results.append(["② 会計品質", "売上計上思想指数", "なし", "データ不足のため判定できません", "grey"])

    # --- 3. 売上構造 ---
    # 3.1 新規取引先数
    if 'partner' not in df_j.columns or df_j['partner'].isna().all():
        results.append(["③ 売上構造", "新規取引先数", "なし", "取引先データが抽出できなかったため判定できません", "grey"])
    elif months_count >= 13 and not df_prev.empty:
        partners_curr = set(df_curr[df_curr['is_sales']]['partner'].dropna())
        partners_prev = set(df_prev[df_prev['is_sales']]['partner'].dropna())
        new_partners = partners_curr - partners_prev
        new_partner_count = len(new_partners)
        
        if new_partner_count == 0:
            color = "red"
            comment = "直近1年間の新規取引先が0社です。既存取引先のみへの依存度が高まっており、顧客の離脱や市場変化に対するリスクが増大しています。"
        elif new_partner_count >= 3:
            color = "blue"
            comment = f"直近1年で{new_partner_count}社の新規取引先を獲得しています（3社以上）。アクティブな顧客開拓が行われており、健全な売上構造の構築が進んでいます。"
        else:
            color = "yellow"
            comment = f"直近1年間の新規取引先は{new_partner_count}社にとどまっています（1〜2社）。新規顧客の開拓ペースが緩やかであるため、さらなる営業活動の強化が望まれます。"
        results.append(["③ 売上構造", "新規取引先数", f"{new_partner_count}社", comment, color])
    else:
        results.append(["③ 売上構造", "新規取引先数", "なし", "比較対象となる昨年のデータがないため判定できません", "grey"])

    # 3.2 新規継続率
    if 'partner' not in df_j.columns or df_j['partner'].isna().all():
        results.append(["③ 売上構造", "新規継続率", "なし", "取引先データが抽出できなかったため判定できません", "grey"])
    elif months_count >= 13 and 'new_partners' in locals() and len(new_partners) > 0:
        retain_count = 0
        for p in new_partners:
            p_rows = df_curr[df_curr['partner'] == p].sort_values('date')
            if len(p_rows) > 1:
                first_date = p_rows.iloc[0]['date']
                second_date = p_rows.iloc[1]['date']
                if (second_date - first_date).days <= 90:
                    retain_count += 1
        
        retention_rate = (retain_count / len(new_partners)) * 100
        if retention_rate < 20:
            color = "red"
            comment = f"新規取引先のうち3ヶ月以内のリピート率（継続率）が{retention_rate:.1f}%と、20%未満の低水準です。初期のアプローチやサービスの満足度に課題がある可能性があり、定着化の仕組み作りが急務です。"
        elif retention_rate >= 50:
            color = "blue"
            comment = f"新規取引先のうち3ヶ月以内のリピート率（継続率）は{retention_rate:.1f}%と、50%以上の高水準です。新規顧客がしっかりと定着しており、良質なサービス提供と関係構築が行われています。"
        else:
            color = "yellow"
            comment = f"新規取引先のうち3ヶ月以内のリピート率（継続率）は{retention_rate:.1f}%と、中程度（20%以上50%未満）です。一定の継続性はありますが、さらなるリピート率向上のための施策が求められます。"
        results.append(["③ 売上構造", "新規継続率", f"{retention_rate:.1f}%", comment, color])
    else:
        results.append(["③ 売上構造", "新規継続率", "なし", "新規取引先がいない、または判定材料が不足しているため判定できません", "grey"])

    # 3.3 粗利率トレンド
    if len(monthly_stats) >= 3:
        margin_trend = monthly_stats['margin'].iloc[-3:]
        v1, v2, v3 = margin_trend.iloc[0], margin_trend.iloc[1], margin_trend.iloc[2]
        if v1 > v2 > v3:
            color = "red"
            res = "3ヶ月連続低下"
            comment = "粗利率が3ヶ月連続で低下しています。コストの急増や価格競争への巻き込まれが強く疑われ、早急な販売価格の見直しや仕入先との交渉が必要です。"
        elif v2 <= v3:
            color = "blue"
            res = "低下なし"
            comment = "粗利率の継続的な低下は見られず、収益性は安定的に推移しています。"
        else:
            color = "yellow"
            res = "単月低下"
            comment = "粗利率が前月比で低下しています。単発的なコスト増の懸念もありますが、このまま低下トレンドが続かないよう注意深く監視する必要があります。"
        results.append(["③ 売上構造", "粗利率トレンド", res, comment, color])
    else:
        results.append(["③ 売上構造", "粗利率トレンド", "なし", "データが3ヶ月分に満たないため判定できません", "grey"])

    # 3.4 上位3社売上集中度
    sales_by_partner = df_j[df_j['is_sales']].groupby('partner')['sales_amt'].sum().sort_values(ascending=False)
    if not sales_by_partner.empty and sales_by_partner.sum() > 0:
        top3_share = (sales_by_partner.head(3).sum() / sales_by_partner.sum()) * 100
        if top3_share >= 70:
            color = "red"
            comment = f"上位3社への売上集中度が{top3_share:.1f}%と、非常に高い水準（70%以上）にあります。主要取引先の業績や方針変更が自社の経営に直撃するリスクがあるため、顧客の分散化が求められます。"
        elif top3_share < 50:
            color = "blue"
            comment = f"上位3社への売上集中度は{top3_share:.1f}%と、健全な水準（50%未満）に抑えられています。特定の顧客に依存しすぎない、バランスの良いポートフォリオが形成されています。"
        else:
            color = "yellow"
            comment = f"上位3社への売上集中度は{top3_share:.1f}%と、中程度（50%以上70%未満）です。極端な集中ではありませんが、中長期的な安定のために新規チャネルの開拓を進めることが望ましいです。"
        results.append(["③ 売上構造", "上位3社売上集中度", f"{top3_share:.1f}%", comment, color])
    else:
        results.append(["③ 売上構造", "上位3社売上集中度", "なし", "取引先別の売上データがありません", "grey"])

    # --- 4. 仕入コスト ---
    # 4.1 上位3社仕入集中度
    cogs_by_partner = df_j[df_j['is_cogs']].groupby('partner')['cogs_amt'].sum().sort_values(ascending=False)
    if not cogs_by_partner.empty and cogs_by_partner.sum() > 0:
        top3_cogs_share = (cogs_by_partner.head(3).sum() / cogs_by_partner.sum()) * 100
        if top3_cogs_share >= 70:
            color = "red"
            comment = f"上位3社への仕入集中度が{top3_cogs_share:.1f}%と、非常に高い水準（70%以上）です。仕入先のトラブル時の供給停止リスクや、価格交渉力の低下の恐れがあるため、複数購買先の検討が推奨されます。"
        elif top3_cogs_share < 50:
            color = "blue"
            comment = f"上位3社への仕入集中度は{top3_cogs_share:.1f}%と、健全な水準（50%未満）です。適切な調達先の分散が行われており、供給リスクが抑制されています。"
        else:
            color = "yellow"
            comment = f"上位3社への仕入集中度は{top3_cogs_share:.1f}%と、中程度（50%以上70%未満）です。急な供給網トラブルに備え、第二・第三の調達候補を視野に入れておくと安心です。"
        results.append(["④ 仕入コスト", "上位3社仕入集中度", f"{top3_cogs_share:.1f}%", comment, color])
    else:
        results.append(["④ 仕入コスト", "上位3社仕入集中度", "なし", "取引先別の仕入データがありません", "grey"])

    # --- 5. その他 ---
    results.append(["⑤ その他", "売上入金リスト", "-", "リストに記載", "white"])
    results.append(["⑤ その他", "直入金売上リスト", "-", "リストに記載", "white"])
    results.append(["⑤ その他", "直払いリスト", "-", "リストに記載", "white"])
    results.append(["⑤ その他", "資金移動用途推定リスト", "-", "リストに記載", "white"])
    results.append(["⑤ その他", "長期未回収売掛リスト", "-", "リストに記載", "white"])

    # DataFrame化して返却
    return pd.DataFrame(results, columns=["category", "item", "result", "comment", "color"])
