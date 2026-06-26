import pandas as pd
import io
import os
from fpdf import FPDF
from fpdf.fonts import FontFace
import datetime
import urllib.parse

# 「問題の本質」の定型文マッピング
ESSENCE_MAP = {
    "現金薄さ": "銀行は手元流動性（月商比）を厳しくチェックします。キャッシュ比率が危険水準（3%未満）にある企業は、債務償還能力や不測の事態への耐性がないと判断され、信用格付が大幅に低下します。",
    "買掛・未払残高": "買掛残高の連続増加は、取引先への支払延期＝資金繰り窮迫のシグナルと銀行に警戒されます。支払余力の低下は、融資判断における重大なマイナス査定要因となります。",
    "仕訳入力遅延": "試算表の提出遅延は、財務管理体制の脆弱性や情報隠蔽（粉飾の疑い）を銀行に強く懸念させます。タイムリーな記帳と開示は、銀行取引の基本条件です。",
    "粗利率ブレ": "月次の粗利率が激しく変動する場合、原価計算の不正確さや期末在庫調整による粉飾を疑われ、決算書や試算表全体の信頼性そのものが損なわれます。",
    "入金サイト延伸": "回収サイトの延伸は、取引先に対する交渉力の低下や、実体のない不良債権（架空売上）の滞留を疑わせ、返済原資の評価においてマイナス査定となります。",
    "新規取引先数": "新規開拓の停滞は、既存取引先への過度な依存を意味し、主要先の業績悪化が自社の連鎖倒産に直結するリスク（売上依存度リスク）として評価を下げます。",
    "新規継続率": "新規客の定着率が低いことは、自社サービス・製品の市場競争力喪失と看做され、中長期的な返済能力の持続性に懸念を持たれます。",
    "粗利率トレンド": "継続的な粗利率の低下は、価格競争への巻き込まれや収益性の悪化を意味し、元利金の返済原資（キャッシュフロー）の減少として銀行から厳しく評価されます。",
    "上位3社売上集中度": "特定先への売上集中は、主要顧客の取引停止や倒産が自社の存続に直撃するハイリスクな事業構造と評価され、与信枠が制限される原因になります。",
    "上位3社仕入集中度": "仕入先の固定化は、調達トラブル時の供給停止リスクや、仕入価格の主導権喪失を意味し、サプライチェーンの脆弱性として銀行に警戒されます。",
    "税金・社会保険料の納付確認": "税金や社会保険料の滞納・支払遅延は、銀行融資において最も重大な拒絶（即否決）事由です。国税優先原則による資産差し押さえや、融資資金の滞納解消への流用を強く警戒されます。",
    "売上計上思想指数": "概算や修正・取消仕訳の多発は、決算書の実態性に重大な疑念を抱かせます。粉飾決算（架空売上）のシグナルと看做され、融資継続が不可能になるリスクがあります。",
    "預金体力推移": "銀行は月末残高だけでなく、月内の資金ショート懸念を最も警戒します。月内最低残高の推移は、企業の真の返済余裕と資金繰りの健全性を示す最重要指標です。",
    "売上入金リスト": "売掛債権が当初の契約通り確実に現金化（回収）されているかは、運転資金の融資審査において最も重視される『返済の確実性』の客観的証拠となります。",
    "直入金売上リスト": "売掛金を経由しない直接入金は、簿外負債の決済や、代表者による資金の私的流用（使途外流用）などの不透明な資金流出として銀行から強く警戒されます。",
    "直払いリスト": "買掛金等を通さない直接現金支払いは、取引実態や価格妥当性の客観的証拠（エビデンス）が不足しているとみなされ、不透明な財務体質として評価を下げます。",
    "資金移動用途推定リスト": "法人口座からの大口の資金流出について、使途（給与、返済等）が不明確な場合、役員貸付金（使途外流用）とみなされ、融資規律の違反として厳格に審査されます。",
    "長期未回収売掛リスト": "30日を超える売掛金の滞留は、融資審査上の『資産査定』において無価値（不良債権）として処理され、実質的な自己資本比率を低下させる重大な要因となります。"
}

def load_essence_map() -> dict:
    """
    Googleスプレッドシートから「問題の本質」を動的に取得する。
    取得失敗時はデフォルトの ESSENCE_MAP を返す。
    """
    spreadsheet_id = "1UySgxZ6uPpxwSY9t994k-jsmR6WY-t8AAIrgHcxn0N4"
    worksheet_name = "問題の本質"
    
    essence = ESSENCE_MAP.copy()
    try:
        encoded_worksheet = urllib.parse.quote(worksheet_name)
        csv_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/gviz/tq?tqx=out:csv&sheet={encoded_worksheet}"
        df_sheet = pd.read_csv(csv_url, header=None)
        
        # 取得するのはA2～B25程度（ヘッダー行を除くため、index 1から）
        max_rows = min(df_sheet.shape[0], 26)  # B25くらいまでなら最大26行
        for i in range(1, max_rows):
            if df_sheet.shape[1] >= 2:
                key = df_sheet.iloc[i, 0]
                val = df_sheet.iloc[i, 1]
                if pd.notna(key) and pd.notna(val):
                    key_str = str(key).strip()
                    val_str = str(val).strip()
                    if key_str != "" and val_str != "":
                        essence[key_str] = val_str
        print("Successfully loaded Essence Map from Google Sheet.")
    except Exception as e:
        print(f"Warning: Error loading Essence Map from Google Sheet ({e}). Using fallback map.")
    return essence

def render_diagnostic_pdf(analysis_df: pd.DataFrame, accounting_period: str) -> tuple[bytes, str]:
    """
    分析結果のDataFrameから、PDFデータ（バイナリ）とプレビュー用テキスト（Markdown）を生成する。
    """
    current_essence_map = load_essence_map()
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
        elif color_val == "yellow":
            fill_color = (255, 255, 204)  # 薄黄色
        elif color_val == "blue":
            fill_color = (230, 240, 255)
        else:
            fill_color = (255, 255, 255)
        row_colors.append(fill_color)
        
        essence = current_essence_map.get(str(row["item"]), "")
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
        col_widths=(15, 25, 15, 50, 85),
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



    # PDFバイナリ取得
    pdf_bytes = bytes(pdf.output())

    # --- 2. プレビュー用 Markdown ---
    preview_md = f"### {summary_msg}\n\n"
    preview_md += "| 評価項目 | 結果 | コメント |\n| :--- | :--- | :--- |\n"
    
    # 赤は最優先、黄色があれば次に優先、それ以降は元々の順番通りに抽出
    red_rows = analysis_df[analysis_df["color"] == "red"]
    yellow_rows = analysis_df[analysis_df["color"] == "yellow"]
    other_rows = analysis_df[~analysis_df["color"].isin(["red", "yellow"])]
    preview_rows = pd.concat([red_rows, yellow_rows, other_rows]).head(3)

    for _, row in preview_rows.iterrows():
        color_val = str(row.get("color", "grey"))
        emoji = "🔴" if color_val == "red" else "🟡" if color_val == "yellow" else "🔵" if color_val == "blue" else "⚪"
        preview_md += f"| {row['item']} | {emoji} {row['result']} | {row['comment']} |\n"
    
    return pdf_bytes, preview_md
