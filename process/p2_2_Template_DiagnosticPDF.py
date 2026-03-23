import pandas as pd
import io
import os
from fpdf import FPDF
from fpdf.fonts import FontFace
import datetime

# 「問題の本質」の定型文マッピング
ESSENCE_MAP = {
    "現金薄さ": "手元資金の不足は経営の心理的余裕を奪い、近視眼的な判断を招きます。常に3ヶ月分程度の現預金を確保する意識が必要です。",
    "買掛・未払残高": "支払いの先延ばしは取引先からの信用毀損に直結します。資金繰りの苦しさが、支払いサイクルに現れ始めていないか注視が必要です。",
    "仕訳入力遅延": "経理作業の遅れは、経営判断の遅れそのものです。『今』の状態が分からないまま舵取りをすることの危うさを認識すべきです。",
    "粗利率ブレ": "月ごとの粗利変動は、原価管理の甘さや、その場しのぎの値引き営業が行われている可能性を示唆しています。",
    "入金サイト延伸": "回収の遅れは、顧客に対する力関係の弱体化を意味します。サービス提供への対価を正当に、迅速に受け取る仕組みを見直すべきです。",
    "新規取引先数": "新規開拓の停滞は、既存事業の陳腐化へのカウントダウンです。常に新しい血を入れ続ける営業活力を維持できているかが問われます。",
    "新規継続率": "新規客が定着しないのは、商品力や初期対応の満足度に課題があるためです。釣り上げた魚を逃さない仕組みの構築が急務です。",
    "粗利率トレンド": "緩やかな粗利率の下落は、競合過多や生産性低下のサインです。価格競争に巻き込まれない独自の価値提供を再定義する時期です。",
    "上位3社売上集中度": "特定顧客への依存は、経営の生殺与奪の権を他者に委ねることと同じです。不測の事態に備え、収益の柱を分散させる戦略が不可欠です。",
    "上位3社仕入集中度": "仕入先の固定化は、コスト削減の機会損失や、供給停止リスクを孕みます。常に代替案を持ち、交渉力を維持する姿勢が求められます。"
}

def render_diagnostic_pdf(analysis_df: pd.DataFrame, accounting_period: str) -> tuple[bytes, str]:
    """
    分析結果のDataFrameから、PDFデータ（バイナリ）とプレビュー用テキスト（Markdown）を生成する。
    """
    # --- 1. PDF生成 (fpdf2) ---
    pdf = FPDF()
    pdf.add_page()
    
    # 日本語フォントの設定 (中略)
    # 1. assetsフォルダ内のカスタムフォントを優先
    # 2. Linux (Streamlit Cloud) の一般的なパス
    # 3. Windows の一般的なパス
    fonts_to_try = [
        os.path.join("assets", "font.ttf"), # 自分で用意する場合
        "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
        r"C:\Windows\Fonts\meiryo.ttc",
        r"C:\Windows\Fonts\YuGothR.ttc",
    ]
    
    font_family = "helvetica"
    for font_path in fonts_to_try:
        if os.path.exists(font_path):
            try:
                # fpdf2 は .ttc の場合、インデックス指定が必要な場合があるが、まずはシンプルに試行
                pdf.add_font("JP-Font", "", font_path)
                font_family = "JP-Font"
                break
            except Exception as e:
                print(f"Font loading failed ({font_path}): {e}")
                continue
    
    # フォントが見つからない場合のフォールバック警告（本当はここでエラーにするか代替手段）
    if font_family == "helvetica":
        print("WARNING: No Japanese font found. PDF may have character errors.")
    
    pdf.set_font(font_family, size=9)
    today_str = datetime.datetime.now().strftime("%Y年%m月%d日")
    pdf.cell(0, 5, f"作成日時：{today_str}", new_x="LMARGIN", new_y="NEXT", align="R")
    pdf.cell(0, 5, f"会計期間：{accounting_period}", new_x="LMARGIN", new_y="NEXT", align="R")
    pdf.ln(5)
    
    pdf.set_font(font_family, size=16)
    
    # ① タイトル
    pdf.cell(0, 15, "特命AI 診断レポート", new_x="LMARGIN", new_y="NEXT", align="C")
    
    # ② 分析結果の概要
    red_count = len(analysis_df[analysis_df["color"] == "red"])
    if red_count == 0:
        summary_msg = "経営の存続に関わる【赤信号】はありませんでした。"
    elif red_count == 1:
        summary_msg = "経営の存続に関わる【赤信号】が一つありました。"
    else:
        summary_msg = "経営の存続に関わる複数の【赤信号】がありました。"
    
    pdf.set_font_size(12)
    pdf.cell(0, 10, f"貴社の会計データを「特命AI財務分析ロジック」で分析した結果", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.set_text_color(255, 0, 0) if red_count > 0 else pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 10, summary_msg, new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(5)
    
    # ③ 画像の挿入 (中央揃え)
    img_path = os.path.join("assets", "特命AI_レポート画像_1.jpeg")
    if os.path.exists(img_path):
        # ページ中央に配置 (A4横幅は210mm)
        img_width = 40
        pdf.image(img_path, x=(210 - img_width) / 2, w=img_width)
        pdf.ln(5)

    # ④ 診断結果表
    pdf.set_font_size(9)
    
    # テーブルデータの準備
    table_data = []
    # ヘッダー
    table_data.append(["カテゴリ", "評価項目", "結果", "コメント", "問題の本質"])
    
    # 内容
    row_colors = []
    for _, row in analysis_df.iterrows():
        color_val = str(row.get("color", "grey"))
        if color_val == "red":
            fill_color = (255, 230, 230)
        elif color_val == "blue":
            fill_color = (230, 240, 255)
        else:
            fill_color = (255, 255, 255)
        row_colors.append(fill_color)
        
        essence = ESSENCE_MAP.get(str(row["item"]), "")
        table_data.append([
            str(row["category"]),
            str(row["item"]),
            str(row["result"]),
            str(row["comment"]),
            essence
        ])

    # テーブルの描画 (fpdf2のtable機能を使用)
    with pdf.table(
        width=190, 
        col_widths=(15, 25, 15, 35, 100),
        text_align=("LEFT", "LEFT", "CENTER", "LEFT", "LEFT"),
        borders_layout="ALL",
        line_height=6,
        headings_style=FontFace(emphasis="")
    ) as table:
        for i, data_row in enumerate(table_data):
            if i == 0:
                # ヘッダー行
                row = table.row()
                style = FontFace(fill_color=(240, 240, 240))
                for datum in data_row:
                    row.cell(datum, style=style)
            else:
                # データ行
                row = table.row()
                fill = row_colors[i-1]
                style = FontFace(fill_color=fill)
                for datum in data_row:
                    row.cell(datum, style=style)

    pdf.ln(10)
    
    # ⑤ 最後の一文
    pdf.set_font(font_family, size=11)
    footer_text = (
        "これらの数値が貴社の「肌感覚」と一致しているか、\n"
        "至急、「答え合わせ」の面談(30分)をお願いいたします。\n"
        "これは、貴社の未来を左右する重大な警告です。\n"
        "早急にご連絡ください。\n\n"
        "大阪キャピタル株式会社　宮田幸治"
    )
    pdf.multi_cell(0, 7, footer_text, align="L")

    # PDFバイナリ取得
    pdf_bytes = bytes(pdf.output())

    # --- 2. プレビュー用 Markdown ---
    preview_md = f"### {summary_msg}\n\n"
    preview_md += "| 評価項目 | 結果 | コメント |\n| :--- | :--- | :--- |\n"
    
    # 赤項目を優先して3項目を抽出
    red_rows = analysis_df[analysis_df["color"] == "red"]
    other_rows = analysis_df[analysis_df["color"] != "red"]
    preview_rows = pd.concat([red_rows, other_rows]).head(3)

    for _, row in preview_rows.iterrows():
        color_val = str(row.get("color", "grey"))
        emoji = "🔴" if color_val == "red" else "🔵" if color_val == "blue" else "⚪"
        preview_md += f"| {row['item']} | {emoji} {row['result']} | {row['comment']} |\n"
    
    return pdf_bytes, preview_md
