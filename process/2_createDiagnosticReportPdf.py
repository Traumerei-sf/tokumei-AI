import pandas as pd
from typing import Dict

from process.p2_1_exeMiyataLogic import exe_miyata_logic
from process.p2_2_Template_DiagnosticPDF import render_diagnostic_pdf

def create_diagnostic_report(df_journal: pd.DataFrame, df_bs: pd.DataFrame) -> Dict:
    """
    診断レポート作成のメインプロセス。
    分析(MiyataLogic)実行 -> PDF化を行い、結果を返す。
    """
    # 1. 分析実行
    analysis_results = exe_miyata_logic(df_journal, df_bs)
    
    # 2. 会計期間の取得
    df_j = df_journal.copy()
    df_j['date'] = pd.to_datetime(df_j['date'], errors='coerce')
    valid_dates = df_j['date'].dropna()
    if not valid_dates.empty:
        min_date = valid_dates.min()
        max_date = valid_dates.max()
        accounting_period = f"{min_date.strftime('%Y年%m月')} - {max_date.strftime('%Y年%m月')}"
    else:
        accounting_period = "不明"

    # 3. PDF生成 & プレビュー生成
    pdf_bytes, preview_md = render_diagnostic_pdf(analysis_results, accounting_period)
    
    return {
        "pdf_bytes": pdf_bytes,
        "preview_md": preview_md,
        "analysis_df": analysis_results
    }
