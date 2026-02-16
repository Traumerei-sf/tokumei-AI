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
    
    # 2. PDF生成 & プレビュー生成
    pdf_bytes, preview_md = render_diagnostic_pdf(analysis_results)
    
    return {
        "pdf_bytes": pdf_bytes,
        "preview_md": preview_md,
        "analysis_df": analysis_results
    }
