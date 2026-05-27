import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.layout import Layout, ManualLayout
from typing import Tuple, Dict

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
    df_sheet1 = df_sheet1[['date', 'amount', 'partner', 'debit_account', 'credit_account']].rename(columns={'partner': '摘要'})
    df_sheet1['取引先'] = df_sheet1['摘要']
    df_sheet1 = df_sheet1.sort_values('date')
    
    # 2. 売上計上思想_全売上仕訳
    df_sheet2 = df_sales.copy()
    df_sheet2['amount'] = df_sheet2[['debit_amount', 'credit_amount']].max(axis=1)
    df_sheet2 = df_sheet2[['date', 'amount', 'partner', 'debit_account', 'credit_account']].rename(columns={'partner': '摘要'})
    df_sheet2['取引先'] = df_sheet2['摘要']
    df_sheet2 = df_sheet2.sort_values('date')
    
    # --- 指標15: 売上入金・直入金売上リスト ---
    debit_pat = '預金|現金|受取手形|電信|当座|普通'
    credit_pat = '売掛|未収'
    
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
        df_sheet3['補助科目(取引先名)'] = df_nyukin['partner']
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
        df_sheet4['補助科目(取引先名)'] = df_choku['partner']
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
        df_sheet5['取引先'] = df_pay['partner']
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
    
    # 支払仕訳の判定関数
    def get_payment_info(row):
        acc = str(row['debit_account'])
        partner = str(row['partner']) if pd.notna(row['partner']) else ""
        
        # 1. 給与支払い
        if any(k in acc for k in ['給料', '役員報酬']) or any(k in partner for k in ['給与', '賞与', '役員']):
            return '給与支払い', 1
        # 2. 銀行返済
        if '借入' in acc or any(k in partner for k in ['返済', '元金', '利息', '融資']):
            return '銀行返済', 2
        # 3. カード決済
        if any(k in partner for k in ['カード', '引き落とし', '決済', 'JC', 'VISA', 'AMEX']):
            return 'カード決済', 3
        # 4. 月次支払い
        if any(k in acc for k in ['買掛', '未払', '外注']):
            return '月次支払い', 4
        # 5. 大口支払い
        if any(k in acc for k in ['仕入', '設備', '外注']):
            return '大口支払い', 5
            
        return None, 99
        
    # 支払系仕訳を抽出
    df_j['pay_type'], df_j['pay_priority'] = zip(*df_j.apply(get_payment_info, axis=1))
    df_payments = df_j[df_j['pay_type'].notna()].copy()
    
    estimated_list = []
    for _, origin in df_origins.iterrows():
        o_date = origin['date']
        o_amt = origin['credit_amount']
        
        # 探索範囲: 起点日 -2日 〜 +7日
        start_range = o_date - pd.Timedelta(days=2)
        end_range = o_date + pd.Timedelta(days=7)
        
        # 近傍の支払仕訳
        near_pays = df_payments[(df_payments['date'] >= start_range) & (df_payments['date'] <= end_range)].copy()
        
        matched_pay = None
        match_type = "なし"
        confidence = "低"
        essence_purpose = "社長関連資金 → 現金引出 → 用途不明"
        related_details = "（なし）"
        
        if not near_pays.empty:
            # 1. 1対1マッチング (金額誤差 ±20% 以内)
            near_pays['diff_pct'] = (near_pays['debit_amount'] - o_amt).abs() / o_amt
            valid_1to1 = near_pays[near_pays['diff_pct'] <= 0.20].sort_values(['pay_priority', 'diff_pct'])
            
            if not valid_1to1.empty:
                matched_pay = valid_1to1.iloc[0]
                match_type = "1対1"
                essence_purpose = matched_pay['pay_type']
                related_details = f"【{matched_pay['debit_account']}】{matched_pay['partner']} ({matched_pay['debit_amount']:.0f}円)"
                
                diff = matched_pay['diff_pct']
                if diff <= 0.02:
                    confidence = "高"
                elif diff <= 0.10:
                    confidence = "中"
                else:
                    confidence = "低"
            else:
                # 2. 1対Nマッチング (想定用途ごとにグループ化し、合計金額が ±20% 以内)
                grouped_pays = near_pays.groupby('pay_type').agg({
                    'debit_amount': 'sum',
                    'pay_priority': 'first',
                    'debit_account': lambda x: ",".join(x.astype(str).unique()),
                    'partner': lambda x: ",".join(x.dropna().astype(str).unique())
                }).reset_index()
                
                grouped_pays['diff_pct'] = (grouped_pays['debit_amount'] - o_amt).abs() / o_amt
                valid_1toN = grouped_pays[grouped_pays['diff_pct'] <= 0.20].sort_values(['pay_priority', 'diff_pct'])
                
                if not valid_1toN.empty:
                    matched_g = valid_1toN.iloc[0]
                    match_type = "1対N"
                    essence_purpose = matched_g['pay_type']
                    related_details = f"【複数口合算】{matched_g['debit_account']} ({matched_g['debit_amount']:.0f}円)"
                    
                    diff = matched_g['diff_pct']
                    if diff <= 0.02:
                        confidence = "高"
                    elif diff <= 0.10:
                        confidence = "中"
                    else:
                        confidence = "低"
                        
        estimated_list.append({
            "移動日": o_date,
            "金額": o_amt,
            "想定用途": essence_purpose,
            "関連仕訳": related_details,
            "一致区分": match_type,
            "信頼度": confidence
        })
        
    df_sheet7 = pd.DataFrame(estimated_list)
    if not df_sheet7.empty:
        df_sheet7 = df_sheet7.sort_values("移動日")
        
    # --- 指標19: 長期未回収売掛リスト ---
    # 発生仕訳: 借方売掛/未収
    df_gen = df_j[df_j['debit_account'].str.contains('売掛|未収', na=False)].copy()
    # 回収仕訳: 貸方売掛/未収
    df_kai = df_j[df_j['credit_account'].str.contains('売掛|未収', na=False)].copy()
    
    base_date = df_j['date'].max()
    uncollected_items = []
    
    # 取引先ごとに名寄せ (空欄は除外)
    partners = set(df_gen['partner'].dropna()) | set(df_kai['partner'].dropna())
    
    for p in partners:
        p_gen = df_gen[df_gen['partner'] == p].sort_values('date')
        p_kai = df_kai[df_kai['partner'] == p].sort_values('date')
        
        total_kai = p_kai['credit_amount'].sum()
        temp_kai = total_kai
        
        for _, row in p_gen.iterrows():
            gen_amt = row['debit_amount']
            gen_date = row['date']
            
            if temp_kai >= gen_amt:
                temp_kai -= gen_amt
            else:
                uncollected_amt = gen_amt - temp_kai
                temp_kai = 0
                
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
                        "金額": uncollected_amt,
                        "滞留日数": days,
                        "評価": eval_str,
                        "補助科目": p
                    })
                    
    # 取引先が空欄の仕訳
    df_gen_null = df_gen[df_gen['partner'].isna() | (df_gen['partner'] == '')]
    for _, row in df_gen_null.iterrows():
        gen_amt = row['debit_amount']
        gen_date = row['date']
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
                "金額": gen_amt,
                "滞留日数": days,
                "評価": eval_str,
                "補助科目": "（空欄）"
            })
            
    df_sheet8 = pd.DataFrame(uncollected_items)
    if not df_sheet8.empty:
        df_sheet8 = df_sheet8.sort_values("滞留日数", ascending=False)
        
    # --- Excelの出力・装飾 (openpyxl) ---
    wb = openpyxl.Workbook()
    # デフォルトのSheetを削除するため、まずはシート作成を行ってから最後に削除する
    
    sheets_info = [
        ("売上計上思想_該当仕訳", df_sheet1, ["日付", "金額", "摘要", "借方科目", "貸方科目", "取引先"]),
        ("売上計上思想_全売上仕訳", df_sheet2, ["日付", "金額", "摘要", "借方科目", "貸方科目", "取引先"]),
        ("売上入金", df_sheet3, ["日付", "金額", "相手科目(借方/貸方)", "摘要", "補助科目(取引先名)"]),
        ("直入金売上", df_sheet4, ["日付", "金額", "相手科目(借方/貸方)", "摘要", "補助科目(取引先名)"]),
        ("直払いリスト", df_sheet5, ["日付", "金額", "借方科目", "摘要", "取引先", "カテゴリ"]),
        ("預金体力推移", df_sheet6, ["対象年月", "月末残高", "月内最低残高", "月内最低残高の記録日"]),
        ("資金移動用途推定", df_sheet7, ["移動日", "金額", "想定用途", "関連仕訳", "一致区分", "信頼度"]),
        ("長期未回収売掛", df_sheet8, ["発生日", "金額", "滞留日数", "評価", "補助科目"])
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
                ws.row_dimensions[row_idx].height = 20
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
                        # 「金額」列の判定（列名に 金額 を含むか、列定義から判断）
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
                        
                # 長期未回収売掛シートでの行背景色（評価に基づく）
                if sheet_name == "長期未回収売掛":
                    eval_val = str(row_data[3]) # 評価列
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
            
        # 列幅の自動調整
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                # 全角文字などを考慮して長さを計算
                length = sum(2 if ord(c) > 127 else 1 for c in val_str)
                if length > max_len:
                    max_len = length
            # グラフがある「預金体力推移」は幅が狭くならないよう考慮
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
    # デフォルトのSheetを削除
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
        
    # bytes に変換
    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_bytes = excel_io.getvalue()
    
    return excel_bytes, sales_index_data
