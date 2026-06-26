import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.layout import Layout, ManualLayout
from typing import Tuple, Dict

# ==========================================================
# 「資金移動用途推定」シート専用の列幅設定（文字数換算）
# 後から列幅を変更したい場合は、以下の数値を編集してください。
# ==========================================================
BANK_LIST_COLUMN_WIDTHS = {
    "移動日": 15,
    "借方金額": 18,
    "貸方金額": 18,
    "想定用途": 20,
    "関連借方仕訳": 45,
    "関連貸方仕訳": 45,
    "一致区分": 15,
    "信頼度": 10
}

def format_journal_entry(account, sub_account, partner, amount) -> str:
    """
    関連仕訳の文字列をフォーマットするヘルパー関数。
    形式: 【[勘定科目]】[補助科目]（[金額]円）
    ※補助科目が空（NaNまたは空文字）の場合、代わりに摘要（partner）を使用します。
    ※括弧は全角「（」および「）」を使用します。
    """
    acc_str = str(account).strip() if pd.notna(account) else ""
    sub_str = str(sub_account).strip() if pd.notna(sub_account) else ""
    pat_str = str(partner).strip() if pd.notna(partner) else ""
    
    # 補助科目が無い場合は摘要を使用
    middle_str = sub_str if sub_str != "" else pat_str
    
    amt_str = f"{amount:.0f}" if pd.notna(amount) else "0"
    
    return f"【{acc_str}】{middle_str}（{amt_str}円）"

def cleanse_journal(df: pd.DataFrame) -> pd.DataFrame:
    """
    仕訳データのクレンジング処理。
    借方金額または貸方金額にマイナス値がある場合、それを正の数に変換し、
    debit/creditを反転させるクレンジングを行う。
    """
    df_clean = df.copy()
    
    # 借方金額がマイナスの場合の処理
    debit_minus = df_clean['debit_amount'] < 0
    if debit_minus.any():
        for idx in df_clean[debit_minus].index:
            row = df_clean.loc[idx]
            df_clean.loc[idx, ['debit_amount', 'credit_amount', 'debit_account', 'credit_account']] = \
                [0.0, abs(row['debit_amount']), row['credit_account'], row['debit_account']]
                
    # 貸方金額がマイナスの場合の処理
    credit_minus = df_clean['credit_amount'] < 0
    if credit_minus.any():
        for idx in df_clean[credit_minus].index:
            row = df_clean.loc[idx]
            df_clean.loc[idx, ['debit_amount', 'credit_amount', 'debit_account', 'credit_account']] = \
                [abs(row['credit_amount']), 0.0, row['credit_account'], row['debit_account']]
                
    return df_clean

def create_bank_excel(df_journal: pd.DataFrame, df_bs: pd.DataFrame) -> Tuple[bytes, Dict]:
    """
    銀行説明用リストのエクセルブック（全8シート）を作成する。
    また、売上計上思想指数の算出データを辞書形式で返す。
    """
    # 1. 仕訳クレンジング
    df_j = cleanse_journal(df_journal)
    df_j['date'] = pd.to_datetime(df_j['date'], errors='coerce')
    df_j = df_j.dropna(subset=['date']).sort_values('date')
    
    # --- 指標14: 売上計上思想指数 ---
    is_sales_debit = df_j['debit_account'].str.contains('売上', na=False) & ~df_j['debit_account'].str.contains('雑収入', na=False)
    is_sales_credit = df_j['credit_account'].str.contains('売上', na=False) & ~df_j['credit_account'].str.contains('雑収入', na=False)
    df_sales = df_j[is_sales_debit | is_sales_credit].copy()
    
    a_keywords = ['概算', '見込', '仮', '仮売上', '予想']
    b_keywords = ['修正', '取消', '訂正', '振替', '再計算']
    all_keywords = a_keywords + b_keywords
    
    def contains_keywords(val):
        if pd.isna(val):
            return False
        val_str = str(val)
        return any(kw in val_str for kw in all_keywords)
        
    df_sales['is_target'] = df_sales['partner'].apply(contains_keywords)
    df_sales_target = df_sales[df_sales['is_target']].copy()
    
    total_sales_count = len(df_sales)
    target_sales_count = len(df_sales_target)
    sales_index = (target_sales_count / total_sales_count * 100) if total_sales_count > 0 else 0.0
    
    sales_index_data = {
        "index": sales_index,
        "target_count": target_sales_count,
        "total_count": total_sales_count
    }
    
    # 1. 売上計上思想_該当仕訳
    df_sheet1 = df_sales_target.copy()
    # 金額は最大値を取得（クレンジングで正の数に変換済み）
    df_sheet1['amount'] = df_sheet1[['debit_amount', 'credit_amount']].max(axis=1)
    df_sheet1 = df_sheet1[['date', 'amount', 'partner', 'debit_account', 'credit_account', 'credit_partner']].rename(columns={'partner': '摘要', 'credit_partner': '貸方補助科目'})
    df_sheet1 = df_sheet1.sort_values('date')
    
    # 2. 売上計上思想_全売上仕訳
    df_sheet2 = df_sales.copy()
    df_sheet2['amount'] = df_sheet2[['debit_amount', 'credit_amount']].max(axis=1)
    df_sheet2 = df_sheet2[['date', 'amount', 'partner', 'debit_account', 'credit_account', 'credit_partner']].rename(columns={'partner': '摘要', 'credit_partner': '貸方補助科目'})
    df_sheet2 = df_sheet2.sort_values('date')
    
    # --- 指標15: 売上入金・直入金売上リスト ---
    debit_pat = '預金|現金|受取手形|電信|当座|普通'
    credit_pat = '売掛|未収|買入金銭債権'
    
    # 3. 売上入金
    df_nyukin = df_j[
        df_j['debit_account'].str.contains(debit_pat, na=False) &
        df_j['credit_account'].str.contains(credit_pat, na=False)
    ].copy()
    df_sheet3 = pd.DataFrame()
    if not df_nyukin.empty:
        df_sheet3['日付'] = df_nyukin['date']
        df_sheet3['金額'] = df_nyukin['debit_amount']
        df_sheet3['相手科目(借方/貸方)'] = df_nyukin['credit_account']
        df_sheet3['摘要'] = df_nyukin['partner']
        df_sheet3['貸方補助科目'] = df_nyukin['credit_partner']
        df_sheet3 = df_sheet3.sort_values('日付')
        
    # 4. 直入金売上
    df_choku = df_j[
        df_j['debit_account'].str.contains(debit_pat, na=False) &
        df_j['credit_account'].str.contains('売上', na=False)
    ].copy()
    df_sheet4 = pd.DataFrame()
    if not df_choku.empty:
        df_sheet4['日付'] = df_choku['date']
        df_sheet4['金額'] = df_choku['debit_amount']
        df_sheet4['相手科目(借方/貸方)'] = df_choku['credit_account']
        df_sheet4['摘要'] = df_choku['partner']
        df_sheet4['貸方補助科目'] = df_choku['credit_partner']
        df_sheet4 = df_sheet4.sort_values('日付')
        
    # --- 指標16: 直払いリスト ---
    exclude_pat = '買掛|未払|借入|利息|税|仮払|手数料'
    is_credit_yokin = df_j['credit_account'].str.contains('普通預金|当座預金', na=False)
    is_not_excluded = ~df_j['debit_account'].str.contains(exclude_pat, na=False)
    df_pay_base = df_j[is_credit_yokin & is_not_excluded].copy()
    
    def get_pay_category(row):
        acc = str(row['debit_account'])
        if '仕入' in acc:
            return '仕入支払'
        elif '外注' in acc:
            return '外注支払'
        elif any(k in acc for k in ['地代', '家賃', '賃借料']):
            return '固定費支払'
        elif any(k in acc for k in ['広告', '宣伝']):
            return '販促費支払'
        elif any(k in acc for k in ['消耗品', '手数料', '修繕', '運賃', '通信', '水道', '光熱', '電気', 'ガス']):
            return 'その他支払'
        return None
        
    df_pay_base['category'] = df_pay_base.apply(get_pay_category, axis=1)
    df_pay = df_pay_base[df_pay_base['category'].notna()].copy()
    
    df_sheet5 = pd.DataFrame()
    if not df_pay.empty:
        df_sheet5['日付'] = df_pay['date']
        df_sheet5['金額'] = df_pay['credit_amount']
        df_sheet5['借方科目'] = df_pay['debit_account']
        df_sheet5['摘要'] = df_pay['partner']
        df_sheet5['借方補助科目'] = df_pay['debit_partner']
        df_sheet5['カテゴリ'] = df_pay['category']
        df_sheet5 = df_sheet5.sort_values('日付')
        
    # --- 指標17: 預金体力推移 ---
    bs_available = (not df_bs.empty) and ("期末現預金合計" in df_bs.columns)
    df_sheet6 = pd.DataFrame()
    
    if bs_available:
        cash_balance_end = df_bs["期末現預金合計"].iloc[0]
        
        # 現預金科目の増減計算
        # 借方に現預金科目がある場合はプラス、貸方にある場合はマイナス
        cash_pat = '預金|現金|当座|普通'
        df_j['debit_is_cash'] = df_j['debit_account'].str.contains(cash_pat, na=False)
        df_j['credit_is_cash'] = df_j['credit_account'].str.contains(cash_pat, na=False)
        
        df_j['cash_diff'] = df_j.apply(
            lambda r: (r['debit_amount'] if r['debit_is_cash'] else 0) - (r['credit_amount'] if r['credit_is_cash'] else 0),
            axis=1
        )
        
        # 期首月の1日（開始仕訳・繰越仕訳）の現預金増減を除外（0にする）
        min_date = df_j['date'].min()
        if pd.notna(min_date):
            opening_month = min_date.month
            is_opening_day = (df_j['date'].dt.month == opening_month) & (df_j['date'].dt.day == 1)
            df_j.loc[is_opening_day, 'cash_diff'] = 0.0
        
        # 日次で集計
        df_daily = df_j.groupby('date')['cash_diff'].sum().reset_index()
        
        # 最終日から遡って残高を計算
        df_daily = df_daily.sort_values('date', ascending=False).reset_index(drop=True)
        daily_balances = []
        current_bal = cash_balance_end
        
        for idx, row in df_daily.iterrows():
            daily_balances.append({
                "date": row['date'],
                "balance": current_bal
            })
            current_bal -= row['cash_diff']
            
        df_bal = pd.DataFrame(daily_balances).sort_values('date').reset_index(drop=True)
        
        # 月次集計
        df_bal['year_month'] = df_bal['date'].dt.to_period('M')
        
        monthly_data = []
        for ym, group in df_bal.groupby('year_month'):
            # 月末残高
            end_bal = group.sort_values('date').iloc[-1]['balance']
            # 月内最低残高
            min_row = group.sort_values('balance').iloc[0]
            min_bal = min_row['balance']
            min_date = min_row['date']
            
            monthly_data.append({
                "対象年月": ym.strftime("%y%m"),
                "月末残高": int(end_bal),
                "月内最低残高": int(min_bal),
                "月内最低残高の記録日": min_date
            })
            
        df_sheet6 = pd.DataFrame(monthly_data).sort_values("対象年月")
        
    # --- 指標18: 資金移動用途推定リスト ---
    # 起点仕訳: 貸方預金かつ50万円以上
    is_origin = df_j['credit_account'].str.contains('普通預金|当座預金', na=False) & (df_j['credit_amount'] >= 500000)
    df_origins = df_j[is_origin].copy()
    
    # 支払・移動・税金仕訳の判定関数
    def get_payment_info(row):
        acc = str(row['debit_account']) if pd.notna(row['debit_account']) else ""
        partner = str(row['partner']) if pd.notna(row['partner']) else ""
        
        # 0. 口座間移動（資金移動）
        if '預金' in acc or any(k in acc for k in ['当座', '普通', '定期', '別段']):
            return '口座間移動（資金移動）', 0
            
        # 0.5. 現金引き出し
        if '現金' in acc:
            return '現金引き出し', 0.5
            
        # 1. 給与支払い
        if any(k in acc for k in ['給料', '役員報酬', '賞与']) or any(k in partner for k in ['給与', '賞与', '役員', '給料']):
            return '給与支払い', 1
            
        # 1.5. 税金・社会保険料
        if any(k in acc for k in ['租税公課', '預り金', '法定福利費']) or \
           any(k in partner for k in ['税', '社会保険', '年金', '健保', '国保', '住民税', '所得税', '厚生年金', '税務署', '市役所', '都税', '県税']):
            return '税金・社会保険料', 1.5
            
        # 2. 銀行返済
        if '借入' in acc or any(k in partner for k in ['返済', '元金', '利息', '融資', '公庫', '保証協会']):
            return '銀行返済', 2
            
        # 3. カード決済
        card_kws = ['カード', '引き落とし', '決済', 'JC', 'VISA', 'AMEX', 'JCB', 'ニコス', 'NICOS', 'セゾン', 'SAISON', '三井住友', '楽天', 'オリコ', 'ジャックス', 'UC', 'ライフ', 'エポス', 'ダイナース', 'UFJカード', '三菱UFJニコス']
        if any(k in acc for k in ['未払金', '未払費用']) and any(k in partner for k in card_kws):
            return 'カード決済', 3
            
        # 4. 月次支払い
        if any(k in acc for k in ['買掛', '未払', '外注']):
            return '月次支払い', 4
            
        # 5. 大口支払い
        if any(k in acc for k in ['仕入', '設備', '土地', '建物', '車両', '機械', '構築物', 'ソフトウェア', '商標', '特許', 'のれん']):
            return '大口支払い', 5
            
        return None, 99
        
    # 支払・移動系仕訳を抽出
    df_j['pay_type'], df_j['pay_priority'] = zip(*df_j.apply(get_payment_info, axis=1))
    df_payments = df_j[df_j['pay_type'].notna()].copy()
    
    card_kws_search = ['カード', '引き落とし', '決済', 'JC', 'VISA', 'AMEX', 'JCB', 'ニコス', 'NICOS', 'セゾン', 'SAISON', '三井住友', '楽天', 'オリコ', 'ジャックス', 'UC', 'ライフ', 'エポス', 'ダイナース', 'UFJカード', '三菱UFJニコス']
    
    estimated_list = []
    for _, origin in df_origins.iterrows():
        o_date = origin['date']
        o_amt = origin['credit_amount']
        t_no = origin['transaction_no']
        
        # --- 自明な複合・単一取引の一括除外フィルター ---
        is_self_evident = False
        
        # A. 取引Noがある場合、その取引No全体の貸借が一致しているかを調べる
        if pd.notna(t_no) and str(t_no).strip() != "":
            related_txs = df_j[df_j['transaction_no'] == t_no]
            debit_sum = related_txs['debit_amount'].sum()
            credit_sum = related_txs['credit_amount'].sum()
            
            # 貸借の合計金額がほぼ一致している場合（誤差1%以内）
            if debit_sum > 0 and abs(debit_sum - credit_sum) / debit_sum <= 0.01:
                # 取引内にカード決済または現金引き出しが含まれているかチェック
                has_card_or_cash = False
                for _, r in related_txs.iterrows():
                    p_type, _ = get_payment_info(r)
                    
                    # 補助科目や摘要にカードキーがある未払金の場合、カード決済に格上げする判定をここでも考慮
                    has_card_kw = any(k in str(r['partner']) or k in str(r['credit_partner']) or k in str(r['debit_partner']) for k in card_kws_search)
                    if has_card_kw and '未払' in str(r['debit_account']):
                        p_type = 'カード決済'
                        
                    if p_type in ['カード決済', '現金引き出し']:
                        has_card_or_cash = True
                        break
                
                # カード決済・現金引き出しを含まない場合、「自明な取引」として除外
                if not has_card_or_cash:
                    is_self_evident = True
                    
        # B. 取引Noがない単一行仕訳の場合（金額一致かつカード・現金以外を除外）
        elif origin['debit_amount'] == origin['credit_amount'] and origin['debit_amount'] > 0:
            p_type, _ = get_payment_info(origin)
            
            has_card_kw = any(k in str(origin['partner']) or k in str(origin['credit_partner']) or k in str(origin['debit_partner']) for k in card_kws_search)
            if has_card_kw and '未払' in str(origin['debit_account']):
                p_type = 'カード決済'
                
            if p_type not in ['カード決済', '現金引き出し']:
                is_self_evident = True
                
        if is_self_evident:
            continue
        
        # 初期状態の設定
        matched_debit_amount = 0.0
        debit_acc_str = "（なし）"
        debit_part_str = "（なし）"
        purpose = "用途不明（社長関連資金等）"
        related_debit = "（なし）"
        related_credit = format_journal_entry(origin['credit_account'], origin['credit_partner'], origin['partner'], origin['credit_amount'])
        
        match_type = "なし"
        confidence = "低"
        is_resolved = False
        
        # --- ステップ1: 同一行での単一仕訳判定 (最優先) ---
        if not is_resolved and origin['debit_amount'] == origin['credit_amount'] and origin['debit_amount'] > 0:
            deb_acc = origin['debit_account']
            deb_part = origin['debit_partner']
            deb_amt = origin['debit_amount']
            
            p_type, p_prio = get_payment_info(origin)
            purpose = p_type if p_type else "その他支払"
            
            # 補助科目や摘要にカードキーがある未払金の場合、カード決済に格上げ
            has_card_kw = any(k in str(origin['partner']) or k in str(origin['credit_partner']) or k in str(origin['debit_partner']) for k in card_kws_search)
            if has_card_kw and '未払' in str(deb_acc):
                purpose = "カード決済"
                
            # カード決済または現金引き出しの場合、時間軸探索へ進む
            if purpose == "カード決済":
                start_range = o_date - pd.Timedelta(days=60)
                end_range = o_date - pd.Timedelta(days=20)
                
                card_name = ""
                for kw in card_kws_search:
                    if kw in str(origin['partner']) or kw in str(origin['credit_partner']) or kw in str(origin['debit_partner']):
                        card_name = kw
                        break
                
                if card_name:
                    card_use_cond = df_j['credit_account'].str.contains('未払金|未払費用', na=False) & \
                                    (df_j['credit_partner'].astype(str).str.contains(card_name, na=False) | \
                                     df_j['partner'].astype(str).str.contains(card_name, na=False))
                else:
                    card_use_cond = df_j['credit_account'].str.contains('未払金|未払費用', na=False)
                    
                near_card_uses = df_j[card_use_cond & (df_j['date'] >= start_range) & (df_j['date'] <= end_range)].copy()
                
                if not near_card_uses.empty:
                    total_use_amt = near_card_uses['debit_amount'].sum()
                    debit_lines = []
                    for _, r in near_card_uses.iterrows():
                        debit_lines.append(format_journal_entry(r['debit_account'], r['debit_partner'], r['partner'], r['debit_amount']))
                    related_debit = "\n".join(debit_lines)
                    
                    matched_debit_amount = total_use_amt
                    debit_acc_str = ",".join(near_card_uses['debit_account'].dropna().unique())
                    debit_part_str = ",".join(near_card_uses['debit_partner'].dropna().unique())
                    match_type = "1対N"
                    
                    diff_pct = abs(total_use_amt - o_amt) / o_amt if o_amt != 0 else 999
                    if diff_pct <= 0.02:
                        confidence = "高"
                    elif diff_pct <= 0.10:
                        confidence = "中"
                    else:
                        confidence = "低"
                    is_resolved = True
                    
            elif purpose == "現金引き出し":
                start_range = o_date
                end_range = o_date + pd.Timedelta(days=10)
                
                cash_use_cond = df_j['credit_account'].str.contains('現金', na=False) & \
                                ~df_j['debit_account'].str.contains('預金|現金|当座|普通', na=False)
                                
                near_cash_uses = df_j[cash_use_cond & (df_j['date'] >= start_range) & (df_j['date'] <= end_range)].copy()
                
                if not near_cash_uses.empty:
                    total_use_amt = near_cash_uses['debit_amount'].sum()
                    debit_lines = []
                    for _, r in near_cash_uses.iterrows():
                        debit_lines.append(format_journal_entry(r['debit_account'], r['debit_partner'], r['partner'], r['debit_amount']))
                    related_debit = "\n".join(debit_lines)
                    
                    matched_debit_amount = total_use_amt
                    debit_acc_str = ",".join(near_cash_uses['debit_account'].dropna().unique())
                    debit_part_str = ",".join(near_cash_uses['debit_partner'].dropna().unique())
                    match_type = "1対N"
                    
                    diff_pct = abs(total_use_amt - o_amt) / o_amt if o_amt != 0 else 999
                    if diff_pct <= 0.02:
                        confidence = "高"
                    elif diff_pct <= 0.10:
                        confidence = "中"
                    else:
                        confidence = "低"
                    is_resolved = True
            
            # 未判定または時間軸探索でマッチしなかった場合の自己完結
            if not is_resolved:
                related_debit = format_journal_entry(origin['debit_account'], origin['debit_partner'], origin['partner'], origin['debit_amount'])
                matched_debit_amount = deb_amt
                debit_acc_str = deb_acc
                debit_part_str = deb_part if pd.notna(deb_part) else ""
                match_type = "単一仕訳直接"
                confidence = "高"
                is_resolved = True

        # --- ステップ2: 取引Noによる複合仕訳の解決 ---
        if not is_resolved and pd.notna(t_no) and str(t_no).strip() != "":
            related_txs = df_j[df_j['transaction_no'] == t_no].copy()
            debits = related_txs[related_txs['debit_amount'] > 0].copy()
            credits = related_txs[related_txs['credit_amount'] > 0].copy()
            
            if not debits.empty:
                debit_total = debits['debit_amount'].sum()
                credit_total = credits['credit_amount'].sum()
                
                # 複数預金口座からの引き出し（貸方が複数）があり、かつ借方全体の合計がこの貸方1行の金額と一致しない場合、
                # 借方仕訳の中から、この貸方金額と合計がほぼ一致する組み合わせ（サブセット）を探索する
                if len(credits) > 1 and abs(debit_total - o_amt) > o_amt * 0.05:
                    import itertools
                    matched_subset = None
                    best_diff = float('inf')
                    
                    # 借方仕訳の数が多すぎると組合せ爆発するため、安全のために最大15件に制限
                    if len(debits) <= 15:
                        debit_list = list(debits.iterrows())
                        for r in range(1, len(debit_list) + 1):
                            for subset in itertools.combinations(debit_list, r):
                                subset_sum = sum(row['debit_amount'] for _, row in subset)
                                diff = abs(subset_sum - o_amt)
                                # 誤差1%以内、かつほぼぴったり一致するものを探索
                                if diff / o_amt <= 0.01:
                                    if diff < best_diff:
                                        best_diff = diff
                                        matched_subset = subset
                                        
                    if matched_subset is not None:
                        # 一致する特定の組み合わせが見つかった場合、その組み合わせの借方のみを対象とする
                        matched_indices = [idx for idx, _ in matched_subset]
                        debits = debits.loc[matched_indices].copy()
                        debit_total = debits['debit_amount'].sum()
                        # この貸方口座に対応する特定の借方グループが特定できたため、
                        # 貸方件数を 1 とみなして安全フィルターを通過させる
                        credits = credits.loc[credits.index == origin.name].copy()
                
                # 値のチェック：複合仕訳内で借方合計と一致するか、あるいは貸方がこの1行のみか
                # 複数預金口座からの引き出しが混ざっている場合、全借方を安易に紐付けないようにする
                if len(credits) == 1 or abs(debit_total - o_amt) <= o_amt * 0.05:
                    # 借方科目から代表用途を特定
                    debits_types = []
                    for _, r in debits.iterrows():
                        p_type, p_prio = get_payment_info(r)
                        if p_type:
                            debits_types.append((p_type, p_prio))
                    
                    if debits_types:
                        debits_types.sort(key=lambda x: x[1])
                        purpose = debits_types[0][0]
                    else:
                        purpose = "その他支払"
                    
                    # カード決済の特記（借方に「未払」や「経費」があり、摘要にカード名がある場合）
                    has_card_keyword = debits['partner'].astype(str).str.contains('|'.join(card_kws_search), na=False).any()
                    if has_card_keyword and purpose in ["その他支払", "月次支払い", "大口支払い"]:
                        purpose = "カード決済"
                    
                    # 関連借方仕訳の構築
                    debit_lines = []
                    for _, r in debits.iterrows():
                        debit_lines.append(format_journal_entry(r['debit_account'], r['debit_partner'], r['partner'], r['debit_amount']))
                    related_debit = "\n".join(debit_lines)
                    
                    # 関連貸方仕訳の構築
                    credit_lines = []
                    for _, r in credits.iterrows():
                        credit_lines.append(format_journal_entry(r['credit_account'], r['credit_partner'], r['partner'], r['credit_amount']))
                    related_credit = "\n".join(credit_lines)
                    
                    matched_debit_amount = debit_total
                    debit_acc_str = ",".join(debits['debit_account'].dropna().unique())
                    debit_part_str = ",".join(debits['debit_partner'].dropna().unique())
                    match_type = "取引No一致"
                    
                    diff_pct = abs(debit_total - o_amt) / o_amt if o_amt != 0 else 999
                    if diff_pct <= 0.02:
                        confidence = "高"
                    elif diff_pct <= 0.10:
                        confidence = "中"
                    else:
                        confidence = "低"
                        
                    is_resolved = True
                
        # --- ステップ3: フォールバック (従来の近傍探索) ---
        if not is_resolved:
            start_range = o_date - pd.Timedelta(days=2)
            end_range = o_date + pd.Timedelta(days=7)
            
            near_pays = df_payments[(df_payments['date'] >= start_range) & (df_payments['date'] <= end_range)].copy()
            
            if not near_pays.empty:
                # 1. 1対1マッチング
                near_pays['diff_pct'] = (near_pays['debit_amount'] - o_amt).abs() / o_amt
                valid_1to1 = near_pays[near_pays['diff_pct'] <= 0.20].sort_values(['pay_priority', 'diff_pct'])
                
                if not valid_1to1.empty:
                    matched_pay = valid_1to1.iloc[0]
                    match_type = "1対1"
                    purpose = matched_pay['pay_type']
                    related_debit = format_journal_entry(matched_pay['debit_account'], matched_pay['debit_partner'], matched_pay['partner'], matched_pay['debit_amount'])
                    matched_debit_amount = matched_pay['debit_amount']
                    debit_acc_str = matched_pay['debit_account']
                    debit_part_str = matched_pay['debit_partner'] if pd.notna(matched_pay['debit_partner']) else ""
                    
                    diff = matched_pay['diff_pct']
                    if diff <= 0.02:
                        confidence = "高"
                    elif diff <= 0.10:
                        confidence = "中"
                    else:
                        confidence = "低"
                else:
                    # 2. 1対Nマッチング (想定用途グループ化)
                    grouped_pays = near_pays.groupby('pay_type').agg({
                        'debit_amount': 'sum',
                        'pay_priority': 'first',
                        'debit_account': lambda x: ",".join(map(str, x.astype(str).unique())),
                        'debit_partner': lambda x: ",".join(map(str, x.dropna().astype(str).unique()))
                    }).reset_index()
                    
                    grouped_pays['diff_pct'] = (grouped_pays['debit_amount'] - o_amt).abs() / o_amt
                    valid_1toN = grouped_pays[grouped_pays['diff_pct'] <= 0.20].sort_values(['pay_priority', 'diff_pct'])
                    
                    if not valid_1toN.empty:
                        matched_g = valid_1toN.iloc[0]
                        match_type = "1対N"
                        purpose = matched_g['pay_type']
                        
                        type_pays = near_pays[near_pays['pay_type'] == purpose]
                        debit_lines = []
                        for _, r in type_pays.iterrows():
                            debit_lines.append(format_journal_entry(r['debit_account'], r['debit_partner'], r['partner'], r['debit_amount']))
                            
                        related_debit = f"【複数口合算】{purpose}（{matched_g['debit_amount']:.0f}円）\n" + "\n".join(debit_lines)
                        matched_debit_amount = matched_g['debit_amount']
                        debit_acc_str = matched_g['debit_account']
                        debit_part_str = matched_g['debit_partner']
                        
                        diff = matched_g['diff_pct']
                        if diff <= 0.02:
                            confidence = "高"
                        elif diff <= 0.10:
                            confidence = "中"
                        else:
                            confidence = "低"
                            
        estimated_list.append({
            "移動日": o_date,
            "銀行名": origin['credit_partner'] if pd.notna(origin['credit_partner']) else "（空欄）",
            "預金形態": origin['credit_account'],
            "借方金額": matched_debit_amount,
            "貸方金額": o_amt,
            "借方勘定科目": debit_acc_str,
            "借方補助科目": debit_part_str if debit_part_str != "" else "（空欄）",
            "想定用途": purpose,
            "関連借方仕訳": related_debit,
            "関連貸方仕訳": related_credit,
            "一致区分": match_type,
            "信頼度": confidence
        })
        
    df_sheet7 = pd.DataFrame(estimated_list)
    if not df_sheet7.empty:
        # 古い日付が上の行になるように日付（移動日）の昇順でソートする
        df_sheet7 = df_sheet7.sort_values("移動日", ascending=True)
        # ご要望の8列のみにフィルタリングし、順番を合わせる
        df_sheet7 = df_sheet7[["移動日", "借方金額", "貸方金額", "想定用途", "関連借方仕訳", "関連貸方仕訳", "一致区分", "信頼度"]]
    else:
        df_sheet7 = pd.DataFrame(columns=["移動日", "借方金額", "貸方金額", "想定用途", "関連借方仕訳", "関連貸方仕訳", "一致区分", "信頼度"])
        
    # --- 指標19: 長期未回収売掛リスト ---
    # 期間と期首日の算出
    min_date = df_j['date'].min()
    start_date = pd.Timestamp(year=min_date.year, month=min_date.month, day=1)
    
    # 2年目の期首日（1年目のちょうど1年後）を判定するための年月
    year2_year = start_date.year + 1
    year2_month = start_date.month
    
    # 2年目の期首日にある「借方が売掛金系」の仕訳をすべて除外するフィルター
    # （開始仕訳等のキーワードに頼らず、該当日の借方発生を一律で除外する）
    ar_pattern_for_carryover = '売掛|未収|買入金銭債権'
    is_year2_opening_debit = (
        (df_j['date'].dt.year == year2_year) & 
        (df_j['date'].dt.month == year2_month) & 
        (df_j['date'].dt.day == 1) & 
        df_j['debit_account'].astype(str).str.contains(ar_pattern_for_carryover, na=False)
    )
    
    clean_df = df_j[~is_year2_opening_debit].copy()
    
    import re
    
    def do_cleanse(val):
        if pd.isna(val):
            return ''
        s = str(val).strip()
        if not s or s.lower() == 'nan':
            return ''
            
        # B: 法人格の削除（スペース関係なく除去）
        corp_patterns = [
            r"株式会社", r"有限会社", r"合資会社", r"合名会社", r"合同会社",
            r"\(株\)", r"（株）", r"\(有\)", r"（有）", r"\(合\)", r"（合）",
            r"㈱", r"㈲", r"㈴", r"㈵", r"法人",
            r"カブシキガイシャ", r"ユウゲンガイシャ", r"ゴウドウガイシャ",
            r"\(カ\)", r"（カ）", r"カ\)", r"\(カ", r"（カ", r"カ）",
            r"カ\.", r"\.カ",
            r"\(ユ\)", r"（ユ）", r"ユ\)", r"\(ユ\)", r"（ユ", r"ユ）",
            r"ユ\.", r"\.ユ",
            r"トクヒ\)", r"\(トクヒ", r"トクヒ"
        ]
        for pat in corp_patterns:
            s = re.sub(pat, "", s)
            
        # A: 「(」または「（」以降を削除
        s = re.split(r'[(（]', s)[0]
        
        # C, D: スペースで分割してトークンごとに判定
        tokens = re.split(r'[ 　]+', s)
        filtered_tokens = []
        for t in tokens:
            if not t:
                continue
            # C: 特定キーワードが存在するトークンを除外
            if any(k in t for k in ['銀行', '金庫', '信用組合', '信組', '農協', '営業部', '支店', '預金', '振込', 'フリコミ', 'ﾌﾘｺﾐ']):
                continue
            # D: 6桁以上の数字を含むトークンを除外（口座番号や管理番号対策）
            if re.search(r'\d{6,}', t):
                continue
            filtered_tokens.append(t)
            
        # 最後にスペースなしで結合
        s = "".join(filtered_tokens)
        s = re.sub(r"^[.\-_ー]+|[.\-_ー]+$", "", s)
        return s

    def get_clean_partner_for_ar(row):
        is_deb_ar = bool(re.search(r'売掛|未収|買入金銭債権', str(row.get('debit_account', ''))))
        is_cred_ar = bool(re.search(r'売掛|未収|買入金銭債権', str(row.get('credit_account', ''))))
        
        clean_dp = do_cleanse(row.get('debit_partner'))
        clean_cp = do_cleanse(row.get('credit_partner'))
        clean_p = do_cleanse(row.get('partner'))
        
        if is_deb_ar:
            # 発生（借方AR）の場合は、借方補助科目を最優先し、なければ摘要
            return clean_dp if clean_dp else clean_p
        elif is_cred_ar:
            # 回収（貸方AR）の場合は、貸方補助科目を最優先し、なければ摘要
            return clean_cp if clean_cp else clean_p
        else:
            # ARに関係ない仕訳の場合は、とりあえず何か入れておく
            return clean_dp if clean_dp else (clean_cp if clean_cp else clean_p)

    # 名寄せ用の取引先名
    clean_df['partner_clean'] = clean_df.apply(get_clean_partner_for_ar, axis=1)
    
    # 発生・回収のどちらにおいても、補助科目と摘要がどちらも空欄（追跡不可）のものは除外
    clean_df = clean_df[clean_df['partner_clean'] != '']
    
    
    # 【包含一致 名寄せロジック】
    # 短い名称が長い名称に含まれる場合、同一取引先とみなして名寄せする。
    raw_partners = [x for x in clean_df['partner_clean'].unique() if x != '（空欄）']
    raw_partners_sorted = sorted(raw_partners, key=len, reverse=True)
    alias_map = {}
    for i, long_p in enumerate(raw_partners_sorted):
        for short_p in raw_partners_sorted[i+1:]:
            if len(short_p) >= 3 and short_p in long_p:
                if len(long_p) > len(short_p):
                    alias_map[long_p] = short_p
                    break
    if alias_map:
        clean_df['partner_clean'] = clean_df['partner_clean'].replace(alias_map)
        
    # AR（売掛・未収）系科目の判定パターン
    ar_pattern = '売掛|未収|買入金銭債権'
    
    # 期首残高仕訳（真の期首日の借方売掛金/未収入金/買入金銭債権）
    is_opening = (clean_df['date'] == start_date) & clean_df['debit_account'].str.contains(ar_pattern, na=False)
    opening_df = clean_df[is_opening]
    opening_bal = opening_df.groupby('partner_clean')['debit_amount'].sum().to_dict()
    
    # 借方・貸方それぞれのAR判定
    is_debit_ar = clean_df['debit_account'].str.contains(ar_pattern, na=False)
    is_credit_ar = clean_df['credit_account'].str.contains(ar_pattern, na=False)
    
    # 【追加フィルター】同一行内で借方・貸方の両方がAR系の場合は「請求締め」等の内部振替とみなし除外
    is_internal_ar_transfer = is_debit_ar & is_credit_ar
    
    # 期中発生仕訳（借方AR、期首仕訳および内部振替を除く）
    df_gen = clean_df[is_debit_ar & ~is_opening & ~is_internal_ar_transfer].copy()
    
    # 期中回収仕訳（貸方AR、内部振替を除く）
    df_kai = clean_df[is_credit_ar & ~is_internal_ar_transfer].copy()
    
    base_date = clean_df['date'].max()
    uncollected_items = []
    debug_ar_list = []
    
    # 全取引先の一覧
    all_partners = set(clean_df['partner_clean'].unique())
    
    for p in all_partners:
        op = opening_bal.get(p, 0.0)
        deb_sum = df_gen[df_gen['partner_clean'] == p]['debit_amount'].sum()
        cred_sum = df_kai[df_kai['partner_clean'] == p]['credit_amount'].sum()
        
        # 逆算フォールバック：回収額が発生と期首を上回る場合は期首を補正
        op_adj = max(op, cred_sum - deb_sum)
        
        # 期内残存売掛金残高（期末残高）
        rem_bal = op_adj + deb_sum - cred_sum
        
        # デバッグ用レコードの追加
        debug_ar_list.append({
            "取引先(名寄せ後)": p,
            "期首残高": op,
            "期中発生(借方)": deb_sum,
            "期中回収(貸方)": cred_sum,
            "期首補正額": op_adj - op,
            "期末残高": rem_bal
        })
        
        if rem_bal <= 0:
            continue
            
        # 発生仕訳（借方）を古い順（昇順）に整理する。
        # まず、期首残高（補正後）を start_date (2023-09-01) の仮想的な発生レコードとしてリストの先頭に配置する。
        p_gens_list = []
        if op_adj > 0:
            # 代表的な科目名を特定
            p_gens_temp = clean_df[(clean_df['partner_clean'] == p) & is_debit_ar]
            acc_name = p_gens_temp['debit_account'].iloc[0] if not p_gens_temp.empty else "売掛金"
            p_gens_list.append({
                "date": start_date,
                "debit_amount": op_adj,
                "debit_account": acc_name
            })
            
        # 次に、期中発生（期首日以外の借方仕訳）を古い順（昇順）に追加する
        p_gen_mid = df_gen[df_gen['partner_clean'] == p].sort_values('date', ascending=True)
        for _, row in p_gen_mid.iterrows():
            p_gens_list.append({
                "date": row['date'],
                "debit_amount": row['debit_amount'],
                "debit_account": row['debit_account']
            })
            
        # 回収の総額 `cred_sum` を用いて、古い発生から順次消し込む
        remaining_cred = cred_sum
        uncollected_gens = []
        
        for g in p_gens_list:
            g_amt = g['debit_amount']
            if remaining_cred >= g_amt:
                # この発生は完全に回収された
                remaining_cred -= g_amt
            elif remaining_cred > 0:
                # この発生は部分的に回収された
                uncollected_amt = g_amt - remaining_cred
                remaining_cred = 0.0
                uncollected_gens.append({
                    "date": g['date'],
                    "amount": uncollected_amt,
                    "debit_account": g['debit_account']
                })
            else:
                # 回収原資が尽きているため、この発生は丸ごと未回収
                uncollected_gens.append({
                    "date": g['date'],
                    "amount": g_amt,
                    "debit_account": g['debit_account']
                })
                
        # 未回収明細のうち、31日以上滞留しているものをリストアップする
        for ug in uncollected_gens:
            gen_date = ug['date']
            amt = ug['amount']
            if amt <= 0.01:
                continue
                
            days = (base_date - gen_date).days
            if days >= 31:
                if days >= 91:
                    eval_str = "🔴長期滞留"
                elif days >= 61:
                    eval_str = "🔶要整理"
                else:
                    eval_str = "⚠️注意"
                    
                uncollected_items.append({
                    "発生日": gen_date,
                    "滞留日数": days,
                    "金額": amt,
                    "取引先": p,
                    "勘定科目": ug['debit_account'],
                    "評価": eval_str
                })
                
    df_sheet8 = pd.DataFrame(uncollected_items, columns=["発生日", "滞留日数", "金額", "取引先", "勘定科目", "評価"])
    if not df_sheet8.empty:
        df_sheet8 = df_sheet8.sort_values("滞留日数", ascending=False)
        
    df_debug_ar = pd.DataFrame(debug_ar_list)
    if not df_debug_ar.empty:
        df_debug_ar = df_debug_ar.sort_values("期末残高", ascending=False)
        
    # --- Excelの出力・装飾 (openpyxl) ---
    wb = openpyxl.Workbook()
    # デフォルトのSheetを削除するため、まずはシート作成を行ってから最後に削除する
    
    sheets_info = [
        ("売上計上思想_該当仕訳", df_sheet1, ["日付", "金額", "摘要", "借方科目", "貸方科目", "貸方補助科目"]),
        ("売上計上思想_全売上仕訳", df_sheet2, ["日付", "金額", "摘要", "借方科目", "貸方科目", "貸方補助科目"]),
        ("売上入金", df_sheet3, ["日付", "金額", "相手科目(借方/貸方)", "摘要", "貸方補助科目"]),
        ("直入金売上", df_sheet4, ["日付", "金額", "相手科目(借方/貸方)", "摘要", "貸方補助科目"]),
        ("直払いリスト", df_sheet5, ["日付", "金額", "借方科目", "摘要", "借方補助科目", "カテゴリ"]),
        ("預金体力推移", df_sheet6, ["対象年月", "月末残高", "月内最低残高", "月内最低残高の記録日"]),
        ("資金移動用途推定", df_sheet7, ["移動日", "借方金額", "貸方金額", "想定用途", "関連借方仕訳", "関連貸方仕訳", "一致区分", "信頼度"]),
        ("長期未回収売掛", df_sheet8, ["発生日", "滞留日数", "金額", "取引先", "勘定科目", "評価"])
    ]
    
    # スタイル定義
    FONT_NAME = "BIZ UDゴシック"
    font_regular = Font(name=FONT_NAME, size=10)
    font_bold = Font(name=FONT_NAME, size=10, bold=True)
    font_header = Font(name=FONT_NAME, size=11, color="FFFFFF", bold=True)
    
    fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    fill_red = PatternFill(start_color="FFE8E8", end_color="FFE8E8", fill_type="solid")
    fill_orange = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
    fill_yellow = PatternFill(start_color="FFFDE7", end_color="FFFDE7", fill_type="solid")
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    
    border_thin = Side(border_style="thin", color="D3D3D3")
    border_cell = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
    
    for sheet_name, df_data, cols in sheets_info:
        ws = wb.create_sheet(title=sheet_name)
        
        # 1. 預金体力推移でB/Sデータがない場合
        if sheet_name == "預金体力推移" and not bs_available:
            ws.cell(row=2, column=2, value="※貸借対照表（B/S）がアップロードされていないため、預金体力推移は算出できません。").font = font_bold
            ws.column_dimensions['B'].width = 80
            continue
            
        # ヘッダー書き込み
        ws.row_dimensions[1].height = 28
        for col_idx, col_name in enumerate(cols, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = border_cell
            
        # データ書き込み
        if not df_data.empty:
            for row_idx, row_data in enumerate(df_data.values, 2):
                row_height = 20
                
                # 「資金移動用途推定」シートはセル内改行数に応じて行の高さを広げる
                if sheet_name == "資金移動用途推定":
                    idx_debit = cols.index("関連借方仕訳") if "関連借方仕訳" in cols else 4
                    idx_credit = cols.index("関連貸方仕訳") if "関連貸方仕訳" in cols else 5
                    val_debit = str(row_data[idx_debit]) if pd.notna(row_data[idx_debit]) else ""
                    val_credit = str(row_data[idx_credit]) if pd.notna(row_data[idx_credit]) else ""
                    lines_debit = val_debit.count('\n') + 1
                    lines_credit = val_credit.count('\n') + 1
                    max_lines = max(lines_debit, lines_credit)
                    row_height = max(20, max_lines * 16) # 1行あたり16pt
                    
                ws.row_dimensions[row_idx].height = row_height
                
                for col_idx, val in enumerate(row_data, 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.font = font_regular
                    cell.border = border_cell
                    
                    # 型に応じたフォーマットと配置
                    if isinstance(val, pd.Timestamp):
                        cell.value = val.strftime('%Y-%m-%d')
                        cell.alignment = align_center
                    elif isinstance(val, (int, float, np.integer, np.floating)):
                        cell.value = val
                        col_name = cols[col_idx-1]
                        if "金額" in col_name or col_name == "月末残高" or col_name == "月内最低残高":
                            cell.number_format = "#,##0"
                            cell.alignment = align_right
                        elif col_name == "滞留日数":
                            cell.number_format = "#,##0"
                            cell.alignment = align_right
                        else:
                            cell.alignment = align_left
                    else:
                        cell.value = str(val) if pd.notna(val) else ""
                        cell.alignment = align_left
                        
                    # 資金移動用途推定シートの折り返し・縦位置上揃え設定
                    if sheet_name == "資金移動用途推定":
                        col_name = cols[col_idx-1]
                        if col_name in ["関連借方仕訳", "関連貸方仕訳"]:
                            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                        else:
                            curr_align = cell.alignment
                            cell.alignment = Alignment(
                                horizontal=curr_align.horizontal if curr_align else "left",
                                vertical="top"
                            )
                        
                # 長期未回収売掛シートでの行背景色（評価に基づく）
                if sheet_name == "長期未回収売掛":
                    eval_val = str(row_data[5]) # 評価列 (インデックス5: 6番目の要素)
                    fill_to_apply = None
                    if "🔴" in eval_val or "長期滞留" in eval_val:
                        fill_to_apply = fill_red
                    elif "🔶" in eval_val or "要整理" in eval_val:
                        fill_to_apply = fill_orange
                    elif "⚠️" in eval_val or "注意" in eval_val:
                        fill_to_apply = fill_yellow
                        
                    if fill_to_apply:
                        for cell in ws[row_idx]:
                            cell.fill = fill_to_apply
                            
        # 預金体力推移シートでの自動グラフ化（B/Sありの場合）
        if sheet_name == "預金体力推移" and bs_available and not df_sheet6.empty:
            chart = LineChart()
            chart.title = "預金体力推移（縦軸は「残高(円)」、横軸は「対象年月」）"
            chart.style = 10
            
            # グラフのサイズをやや広げて視認性と余白を確保
            chart.width = 20
            chart.height = 13
            
            # プロットエリア自体のレイアウトを設定して、上下左右に白い余白を確保
            # これにより、タイトル、軸タイトル、軸数値、凡例がグラフ線や枠線と被るのを防ぎます
            chart.plot_area.layout = Layout(
                manualLayout=ManualLayout(
                    x=0.15,      # 左余白 (全体を1.0とした割合、左から15%の位置から開始)
                    y=0.15,      # 上余白 (上から15%の位置から開始)
                    w=0.70,      # プロットエリアの幅 (右側に15%の凡例用余白を確保)
                    h=0.68,      # プロットエリアの高さ (下側に17%の横軸ラベル・タイトル用余白を確保)
                    xMode="edge",
                    yMode="edge"
                )
            )
            
            # 月末残高と月内最低残高 (B列とC列)
            data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=len(df_sheet6)+1)
            # 対象年月 (A列)
            cats = Reference(ws, min_col=1, min_row=2, max_row=len(df_sheet6)+1)
            
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            
            # 軸の明示的表示（削除フラグをFalseに設定し、数値フォーマットを適用）
            chart.y_axis.delete = False
            chart.x_axis.delete = False
            chart.y_axis.number_format = '#,##0'
            
            # 凡例の位置調整：右上かつプロットエリア外に配置してグラフと重ねない
            chart.legend.position = "tr"  # Top Right (右上)
            chart.legend.overlay = False  # グラフエリアと重ねない
            
            # 横軸ラベル（年月）が重なるのを防ぐために-45度斜め回転
            from openpyxl.chart.text import RichText
            from openpyxl.drawing.text import RichTextProperties
            chart.x_axis.txPr = RichText(
                bodyPr=RichTextProperties(
                    rot="-2700000",  # -45度 (角度 * -60,000)
                    anchor="ctr",
                    anchorCtr="1",
                    spcFirstLastPara="1",
                    vertOverflow="ellipsis",
                    wrap="square"
                )
            )
            
            # グラフの折れ線のデザイン変更 (青基調、直線的)
            colors = ["1B365D", "4169E1"]
            for i, color_hex in enumerate(colors):
                if i < len(chart.series):
                    s = chart.series[i]
                    s.graphicalProperties.line.solidFill = color_hex
                    s.graphicalProperties.line.width = 25000  # 2.5pt
                    s.smooth = False
            
            # グラフの配置
            ws.add_chart(chart, "F2")
            
        # 列幅の設定
        if sheet_name == "資金移動用途推定":
            # 「資金移動用途推定」シートは自動調整を行わず、ファイル最上部で指定された固定幅を直接適用する
            for col_idx, col_name in enumerate(cols, 1):
                col_letter = get_column_letter(col_idx)
                width = BANK_LIST_COLUMN_WIDTHS.get(col_name, 15)
                ws.column_dimensions[col_letter].width = width
        else:
            # その他のシートは従来どおり自動スケーリング（改行考慮・全列合計最大200制限）を適用する
            col_widths = {}
            total_width = 0
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    val_str = str(cell.value or '')
                    # 改行で分割し、各行の中での最大長を測定する
                    lines = val_str.split('\n')
                    for line in lines:
                        length = sum(2 if ord(c) > 127 else 1 for c in line)
                        if length > max_len:
                            max_len = length
                
                # 推奨幅（文字数 + バッファ）
                recommended_width = max(max_len + 4, 12)
                col_widths[col_letter] = recommended_width
                total_width += recommended_width
                
            # 全列合計の列幅を最大200に制限するスケーリング調整
            MAX_TOTAL_WIDTH = 200
            if total_width > MAX_TOTAL_WIDTH:
                scale_factor = MAX_TOTAL_WIDTH / total_width
                for col_letter, w in col_widths.items():
                    col_widths[col_letter] = max(w * scale_factor, 8) # 最小幅は8を維持
                    
            # 列幅の適用
            for col_letter, w in col_widths.items():
                ws.column_dimensions[col_letter].width = w
            
    # デフォルトのSheetを削除
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
        
    # bytes に変換
    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_bytes = excel_io.getvalue()
    
    return excel_bytes, sales_index_data
