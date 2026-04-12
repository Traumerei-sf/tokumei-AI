import pandas as pd
import io
import json
import re
from typing import Optional, Dict, List, Tuple
from process.u_accessGemini import exe_gemini_structure_forJournal, exe_gemini_structure_forBS

# --- 定数定義 ---
STANDARD_JOURNAL_COLUMNS = [
    "date", "debit_account", "debit_amount", "credit_account", "credit_amount", "partner", "created_at"
]

def load_file_to_df(file: io.BytesIO) -> pd.DataFrame:
    """
    ファイルを読み込み DataFrame に変換する (1番目のシート)
    """
    file.seek(0)
    filename = getattr(file, "name", "").lower()
    
    if filename.endswith(".xlsx"):
        return pd.read_excel(file, sheet_name=0)
    else:
        # CSV の場合はエンコーディングを試行
        for enc in ['utf-8-sig', 'cp932', 'utf-8', 'shift_jis']:
            try:
                file.seek(0)
                return pd.read_csv(file, encoding=enc)
            except:
                continue
    raise ValueError("ファイル形式が csv または xlsx ではありません。")

def _flatten_journal(df: pd.DataFrame) -> pd.DataFrame:
    """
    借方・貸方に分かれた仕訳データを、1レコード1科目のフラットな形式に変換する。
    (診断ロジックの内部で使用される可能性があるため残す)
    """
    if "debit_account" in df.columns and "credit_account" in df.columns:
        common_cols = [c for c in df.columns if c not in ["debit_account", "debit_amount", "credit_account", "credit_amount"]]
        # 借方
        debit_df = df[common_cols].copy()
        debit_df["account"] = df["debit_account"]
        debit_df["amount"] = df["debit_amount"]
        debit_df["side"] = "debit"
        # 貸方
        credit_df = df[common_cols].copy()
        credit_df["account"] = df["credit_account"]
        credit_df["amount"] = df["credit_amount"]
        credit_df["side"] = "credit"
        flat_df = pd.concat([debit_df, credit_df], ignore_index=True)
        return flat_df[flat_df["account"].notna()]
    return df

def process_journal_single(file: io.BytesIO, file_num: int = 1) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """
    1枚の仕訳帳ファイルを処理する。
    返り値: (DataFrame, ErrorMessage) ※レポート作成用にWide形式で返す。
    """
    try:
        # 1. 読み込み
        df_raw = load_file_to_df(file)
        if df_raw.empty:
            return None, "ファイルが空です。"

        # 2. Gemini へのプロンプト作成
        sample_data = df_raw.head(15).astype(str).values.tolist()
        columns_info = list(df_raw.columns.astype(str))
        
        prompt = f"""
あなたはプロの会計士でありデータアナリストです。
提供された会計データ（仕訳帳）のサンプルから、各項目が「何列目」にあるかと、「実データが何行目から始まるか」を特定してください。

【項目定義】
- date: 取引日（日付）
- debit_account: 借方勘定科目
- debit_amount: 借方金額
- credit_account: 貸方勘定科目
- credit_amount: 貸方金額
- partner: 取引先名称や摘要（相手先）。借方・貸方などの区別は不要で、その取引行の代表的な1列だけを特定してください。
- created_at: 作成日時/入力日時（作成日時, 入力日付時間, 入力日付, 登録日時, 入力日, 入力日時, 仕分日, 仕分日時, Created At 等）

【サンプルデータ】
ヘッダー候補: {columns_info}
データサンプル（最初の15行）:
{json.dumps(sample_data, ensure_ascii=False, indent=2)}

【ルール】
- 列番号は0からカウントしてください。
- 該当する項目が見当たらない場合は null を返してください。
- date, debit_account, debit_amount, credit_account, credit_amount は必須項目です。
- data_start_row は、ヘッダーを除く実際のデータ（1件目の取引）が始まる「元のデータの行番号」を指定してください。
"""
        
        # 3. Gemini 呼び出し
        print("--- DEBUG: Gemini Prompt (created_at inclusion) ---")
        response_json = exe_gemini_structure_forJournal(prompt)
        print(f"--- DEBUG: Gemini Raw Response ---\n{response_json}")

        try:
            json_str = response_json.strip()
            if json_str.startswith("```"):
                json_str = re.sub(r'^```(?:json)?\n?|\n?```$', '', json_str, flags=re.MULTILINE)
            mapping_data = json.loads(json_str)
        except Exception as e:
            return None, f"JSON解析エラー: {str(e)}"
        
        mapping = mapping_data.get("column_mapping", {})
        start_row = mapping_data.get("data_start_row", 0)

        # 4. バリデーション
        required_keys = ["date", "debit_account", "debit_amount", "credit_account", "credit_amount"]
        missing_keys = [k for k in required_keys if mapping.get(k) is None]
        if missing_keys:
            return None, f"必須項目不足: {', '.join(missing_keys)}"

        # 5. データ抽出 (Wide形式)
        extracted_data = {}
        for std_name in STANDARD_JOURNAL_COLUMNS:
            col_idx = mapping.get(std_name)
            if col_idx is not None and col_idx < len(df_raw.columns):
                extracted_data[std_name] = df_raw.iloc[start_row:, col_idx].reset_index(drop=True)
            else:
                extracted_data[std_name] = pd.NA
        df_wide = pd.DataFrame(extracted_data)
        
        # 6. クリーニング
        def robust_parse_date(val):
            if pd.isna(val) or str(val).strip() == "": return pd.NaT
            s = str(val).strip()
            s = s.translate(str.maketrans('０１２３４５６７８９．', '0123456789/'))
            era_map = {"令和": 2018, "平成": 1988, "昭和": 1925}
            for era, base_year in era_map.items():
                if era in s:
                    match = re.search(rf"{era}(\d+)年(\d+)月(\d+)日", s)
                    if match:
                        y, m, d = match.groups()
                        s = f"{int(y) + base_year}/{m}/{d}"
                        break
            try: return pd.to_datetime(s, errors='coerce')
            except: return pd.NaT

        def clean_amount(val):
            if pd.isna(val) or str(val).strip() == "": return 0.0
            s = str(val).replace(',', '').replace('¥', '').replace('円', '').replace('△', '-').strip()
            s = s.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
            cleaned = re.sub(r'[^\d.-]', '', s)
            try: return float(cleaned) if cleaned else 0.0
            except: return 0.0

        # 型変換
        df_wide["date"] = df_wide["date"].apply(robust_parse_date)
        df_wide["created_at"] = df_wide["created_at"].apply(robust_parse_date)
        df_wide["debit_amount"] = df_wide["debit_amount"].apply(clean_amount)
        df_wide["credit_amount"] = df_wide["credit_amount"].apply(clean_amount)

        def clean_partner(val):
            if pd.isna(val) or str(val).strip() == "": return pd.NA
            s = str(val)
            
            # 1. メタデータ的な塊を削除（依頼人名:以降の連続する文字列など）
            s = re.sub(r"\(?依頼人名:\S*", "", s)
            s = re.sub(r"振込予定日:\S*", "", s)
            s = re.sub(r"管理番号:\S*", "", s)
            
            # 追加: X月分、令和X年X月分などの日付メタデータを削除
            s = re.sub(r"(?:令和|平成|昭和)\d+年(?:度)?\d+月(?:分|度)?", "", s)
            s = re.sub(r"年\d+月(?:分|度)?", "", s)
            s = re.sub(r"\d+月(?:分|度)?", "", s)

            # 追加: 業務関連の定型ワードを削除
            biz_words = [
                "業務委託費用", "業務委託費", "業務委託料", "業務委託",
                "コンサルティング報酬", "コンサルティング費用", "コンサルティング手数料", "コンサルティング料", "コンサルティング", "コンサルタント料",
                "アドバイザリー費用", "インスタグラム運用代行", "運用代行",
                "初期費用", "月額費用", "共同事業分", "年契約", "経営パートナー契約",
                "稼働料金", "コスト削減", "採用戦略", "経営", "返金？", "作成", "代行"
            ]
            for w in biz_words:
                s = s.replace(w, "")
                
            # 2. 定型ノイズワードを削除（完全一致部分のみ）
            noise_words = [
                "他行振込手数料", "振込手数料", "手数料", "他行振込", "組戻手数料", "現金売上", "現金仕入",
                "消費税", "普通預金", "当座預金", "普通", "当座", "法人立ち上げ"
            ]
            for w in noise_words:
                s = s.replace(w, "")
                
            # 3. 銀行・支店・長い数字（口座番号）などを削除
            s = re.sub(r"\S*銀行", "", s)
            s = re.sub(r"\S*信用金庫", "", s)
            s = re.sub(r"\S*信用組合", "", s)
            s = re.sub(r"\S*支店", "", s)
            s = re.sub(r"\d{4,}", "", s) # 4桁以上の数字
            
            # 4. 全てのスペースを削除（名寄せのため）
            s = s.replace(" ", "").replace("　", "")
            
            # 5. 先頭や末尾に取り残された孤立したカッコなどの除去
            s = re.sub(r"^[(\[【]+|[)\]】]+$", "", s)
            
            s = s.strip()
            return s if s else pd.NA

        if "partner" in df_wide.columns:
            df_wide["partner"] = df_wide["partner"].apply(clean_partner)

        # 日付前方埋め
        df_wide["date"] = df_wide["date"].ffill()
        
        # デバッグ表示
        print("--- DEBUG: Extracted Samples (Wide Format) ---")
        for col in df_wide.columns:
            print(f"DEBUG: {col}: {df_wide[col].head(5).tolist()}")

        # 有効行フィルタ
        df_wide = df_wide[df_wide["date"].notna() & (df_wide["debit_account"].notna() | df_wide["credit_account"].notna())]
        
        if df_wide.empty:
            return None, "有効データなし"

        print(f"--- DEBUG: ファイル{file_num}枚目から、{len(df_wide)}件の取引データを抽出しました ---")

        # 8. 期間確認
        df_wide = df_wide.sort_values("date")
        min_date = df_wide["date"].min()
        max_date = df_wide["date"].max()
        months = (max_date.year - min_date.year) * 12 + (max_date.month - min_date.month) + 1
        if months > 36:
            return None, f"このファイル単体で期間が長すぎます（{months}ヶ月）。"

        return df_wide, None

    except Exception as e:
        return None, f"SAD内部エラー: {str(e)}"

def process_bs_single(file: io.BytesIO) -> Tuple[Optional[Dict], Optional[str]]:
    """
    1枚の貸借対照表ファイルを処理する。
    """
    try:
        print("--- DEBUG: Starting B/S Processing ---")
        df_raw = load_file_to_df(file)
        if df_raw.empty:
            return None, "ファイルが空です。"

        # 全情報をCSV化して渡す
        csv_text = df_raw.to_csv(index=False)

        prompt = f"""
あなたはプロの財務アナリストです。
提供された貸借対照表（B/S）の全データから、以下の2つの情報を読み取って抽出してください。

1. この貸借対照表における、期末の年月（YYYY/MM 形式の文字列で。例: 2024/09）
2. この貸借対照表における、現預金の合計金額（数値で）

貸借対照表データ：
{csv_text}
"""
        response_json = exe_gemini_structure_forBS(prompt)
        
        try:
            json_str = response_json.strip()
            if json_str.startswith("```"):
                json_str = re.sub(r'^```(?:json)?\n?|\n?```$', '', json_str, flags=re.MULTILINE)
            result = json.loads(json_str)
        except Exception as e:
            return None, f"JSON解析エラー: {str(e)}"

        ym = result.get("year_month")
        cash = result.get("cash_amount")

        if ym is None or cash is None:
            return None, "期末の年月、または現預金の合計金額を取得できませんでした。"
        
        # ymが正しい形式か確認
        if not re.match(r"^\d{4}/\d{2}$", str(ym)):
            # geminiが推測したフォーマットが崩れている場合
            pass

        return {"year_month": str(ym), "cash_amount": cash}, None

    except Exception as e:
        return None, f"SAD(BS)内部エラー: {str(e)}"

def check_accounting_files(file_journal1: io.BytesIO, file_bs: Optional[io.BytesIO] = None, 
                           file_journal2: Optional[io.BytesIO] = None) -> List[Dict]:
    return [{"message": "OK", "color": "green", "success": True}]

def standardize_logic(file_journal1: io.BytesIO, file_journal2: Optional[io.BytesIO] = None, 
                      file_bs: Optional[io.BytesIO] = None) -> Dict[str, pd.DataFrame]:
    # 互換性維持
    df1, _ = process_journal_single(file_journal1)
    return {"journal": df1, "bs": pd.DataFrame()}

if __name__ == "__main__":
    print("Standardization logic module loaded.")