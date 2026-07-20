import pandas as pd
import io
import json
import re
from typing import Optional, Dict, List, Tuple
from process.u_accessGemini import exe_gemini_structure_forJournal, exe_gemini_structure_forBS
from process.partner_resolution import normalize_description, resolve_partner_columns

# --- 定数定義 ---
STANDARD_JOURNAL_COLUMNS = [
    "date", "debit_account", "debit_amount", "credit_account", "credit_amount", "description", "partner", "created_at", "transaction_no", "debit_partner", "credit_partner"
]

TRANSACTION_NO_NULL_STRINGS = {"", "nan", "none", "null", "<na>"}


def normalize_transaction_numbers(series: pd.Series, file_num: int) -> pd.Series:
    """有効な取引Noだけを文字列化し、ファイル単位の接頭辞を付ける。"""
    normalized = series.astype("string").str.strip()
    is_missing = normalized.isna() | normalized.str.lower().isin(TRANSACTION_NO_NULL_STRINGS)
    normalized = normalized.mask(is_missing, pd.NA)
    valid = normalized.notna()
    normalized.loc[valid] = f"{file_num}_" + normalized.loc[valid]
    return normalized

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
- partner: 摘要、備考、取引内容が書かれた列の列番号。補助科目・取引先列ではなく、仕訳の説明文を優先してください。
- debit_partner: 借方の補助科目、または借方の取引先名が記載されている列（例:「借方 補助科目」「補助科目(借方)」「補助科目」など。存在しない場合は null）
- credit_partner: 貸方の補助科目、または貸方の取引先名が記載されている列（例:「貸方 補助科目」「補助科目(貸方)」「補助科目」など。存在しない場合は null）
  ※「補助科目」列が借方・貸方で分かれておらず、1つしか無い場合は、debit_partner と credit_partner の両方にその同じ列番号を指定してください。
- transaction_no: 同一の取引（複合仕訳など）をバインドするための「伝票No.」「仕訳No」「伝票No」「No」等の識別子が記載されている列。
  ※【極めて重要】ヘッダーに「伝票No.」や「伝票番号」などの明確なラベルが存在する列を最優先で選んでください。
  ※【極めて重要】0列目などに入っている「単なる行インデックス（1, 2, 3... や 17111, 17112... のような単なる行番号）」は取引No（伝票No）ではありません。ヘッダーが空欄で単なる行の連番になっている列は、絶対に transaction_no にマッピングせず、null を指定してください。
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
        
        debit_part_idx = mapping.get("debit_partner")
        credit_part_idx = mapping.get("credit_partner")
        partner_idx = mapping.get("partner")
        
        debit_acc_idx = mapping.get("debit_account")
        credit_acc_idx = mapping.get("credit_account")
        
        merged_partner = None
        base_partner = pd.Series(pd.NA, index=range(len(df_raw) - start_row))
        if (debit_part_idx is not None and debit_part_idx < len(df_raw.columns)) or \
           (credit_part_idx is not None and credit_part_idx < len(df_raw.columns)) or \
           (partner_idx is not None and partner_idx < len(df_raw.columns)):
            
            debit_series = df_raw.iloc[start_row:, debit_part_idx].reset_index(drop=True) if (debit_part_idx is not None and debit_part_idx < len(df_raw.columns)) else pd.Series(pd.NA, index=range(len(df_raw) - start_row))
            credit_series = df_raw.iloc[start_row:, credit_part_idx].reset_index(drop=True) if (credit_part_idx is not None and credit_part_idx < len(df_raw.columns)) else pd.Series(pd.NA, index=range(len(df_raw) - start_row))
            base_partner = df_raw.iloc[start_row:, partner_idx].reset_index(drop=True) if (partner_idx is not None and partner_idx < len(df_raw.columns)) else pd.Series(pd.NA, index=range(len(df_raw) - start_row))
            
            debit_acc_series = df_raw.iloc[start_row:, debit_acc_idx].reset_index(drop=True) if (debit_acc_idx is not None and debit_acc_idx < len(df_raw.columns)) else pd.Series("", index=range(len(df_raw) - start_row))
            credit_acc_series = df_raw.iloc[start_row:, credit_acc_idx].reset_index(drop=True) if (credit_acc_idx is not None and credit_acc_idx < len(df_raw.columns)) else pd.Series("", index=range(len(df_raw) - start_row))
            
            # 口座・現金科目の判定パターン
            yokin_pat = "預金|現金|当座|普通|手形|電信"
            
            # 借方/貸方が預金科目の場合は、その補助科目を無効化（NaNにする）
            # なぜなら口座名は取引先名ではないからである
            is_debit_yokin = debit_acc_series.astype(str).str.contains(yokin_pat, na=False)
            is_credit_yokin = credit_acc_series.astype(str).str.contains(yokin_pat, na=False)
            
            debit_series_cleaned_yokin = debit_series.copy()
            debit_series_cleaned_yokin[is_debit_yokin] = pd.NA
            
            credit_series_cleaned_yokin = credit_series.copy()
            credit_series_cleaned_yokin[is_credit_yokin] = pd.NA
            
            # NaNや空欄を適切に処理してマージ (debit_partner -> credit_partner -> partner の優先順位)
            debit_series_clean = debit_series_cleaned_yokin.replace(r'^\s*$', pd.NA, regex=True)
            credit_series_clean = credit_series_cleaned_yokin.replace(r'^\s*$', pd.NA, regex=True)
            base_partner_clean = base_partner.replace(r'^\s*$', pd.NA, regex=True)
            
            merged_partner = debit_series_clean.fillna(credit_series_clean).fillna(base_partner_clean)
        
        for std_name in STANDARD_JOURNAL_COLUMNS:
            if std_name == "description":
                extracted_data["description"] = base_partner.reset_index(drop=True)
            elif std_name == "partner" and merged_partner is not None:
                extracted_data["partner"] = merged_partner
            else:
                col_idx = mapping.get(std_name)
                if col_idx is not None and col_idx < len(df_raw.columns):
                    extracted_data[std_name] = df_raw.iloc[start_row:, col_idx].reset_index(drop=True)
                else:
                    extracted_data[std_name] = pd.NA
        df_wide = pd.DataFrame(extracted_data)
        # 名寄せの根拠監査と法人判定に使うため、正規化前の値を保持する。
        for raw_column in ("description", "partner", "debit_partner", "credit_partner"):
            if raw_column in df_wide.columns:
                df_wide[f"{raw_column}_raw"] = df_wide[raw_column].copy()
        
        # 6. クリーニング
        # 6. クリーニング
        def vectorized_parse_date(series):
            if series is None or series.empty: return pd.to_datetime(series)
            s = series.astype(str).str.strip().replace("nan", "")
            s = s.str.translate(str.maketrans('０１２３４５６７８９．', '0123456789/'))
            s = s.str.replace(r'令和(\d+)年(\d+)月(\d+)日', lambda m: f"{int(m.group(1))+2018}/{m.group(2)}/{m.group(3)}", regex=True)
            s = s.str.replace(r'平成(\d+)年(\d+)月(\d+)日', lambda m: f"{int(m.group(1))+1988}/{m.group(2)}/{m.group(3)}", regex=True)
            s = s.str.replace(r'昭和(\d+)年(\d+)月(\d+)日', lambda m: f"{int(m.group(1))+1925}/{m.group(2)}/{m.group(3)}", regex=True)
            return pd.to_datetime(s, errors='coerce')

        def vectorized_clean_amount(series):
            if series is None or series.empty: return pd.to_numeric(series)
            # `\¥` はLinuxの正規表現エンジンで不正なエスケープになるため、
            # 通貨・桁区切り文字はOS非依存のリテラル置換で除去する。
            s = series.astype(str)
            for token in (',', '¥', '\\', '円'):
                s = s.str.replace(token, '', regex=False)
            s = s.str.replace('△', '-', regex=False)
            s = s.str.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
            s = s.str.replace(r'[^\d.-]', '', regex=True)
            return pd.to_numeric(s, errors='coerce').fillna(0.0)

        # 型変換
        if "date" in df_wide.columns:
            df_wide["date"] = vectorized_parse_date(df_wide["date"])
        if "created_at" in df_wide.columns:
            df_wide["created_at"] = vectorized_parse_date(df_wide["created_at"])
        if "debit_amount" in df_wide.columns:
            df_wide["debit_amount"] = vectorized_clean_amount(df_wide["debit_amount"])
        if "credit_amount" in df_wide.columns:
            df_wide["credit_amount"] = vectorized_clean_amount(df_wide["credit_amount"])

        # 法人格の正規表現パターン（一つに結合して高速化）
        corp_pattern = (
            r"株式会社|有限会社|合資会社|合名会社|合同会社|"
            r"\(株\)|（株）|\(有\)|（有）|\(合\)|（合）|"
            r"㈱|㈲|㈴|㈵|法人|"
            r"カブシキガイシャ|ユウゲンガイシャ|ゴウドウガイシャ|"
            r"\(カ\)|（カ）|カ\)|\(カ|（カ|カ）|カ\.|\\.カ|"
            r"\(ユ\)|（ユ）|ユ\)|\(ユ\)|（ユ|ユ）|ユ\.|\\.ユ|"
            r"トクヒ\)|\(トクヒ|トクヒ"
        )
        prefix_pattern = r"^(?:振込|フリコミ|ﾌﾘｺﾐ|振込口|組戻|クミモドシ|クミモド|トウニユウ|トウニュウ|トウニユウグチ|ネット|ネツト)"
        
        def vectorized_clean_partner(series):
            if series is None or series.empty: return series
            # 0. 文字列化し、欠損値などを空文字に
            s = series.fillna("").astype(str)
            # Unicode正規化はapplyが必要だが組み込み関数なので比較的高速
            import unicodedata
            s = s.apply(lambda x: unicodedata.normalize('NFKC', x) if x else "")
            
            # ベクトル化された文字列置換 (一括処理)
            s = s.str.replace(prefix_pattern, "", regex=True)
            s = s.str.replace(corp_pattern, "", regex=True)
            s = s.str.replace(r"[\n\r\t]+", "", regex=True)
            s = s.str.replace(r"^[(\[【.\-_ー]+|[)\]】.\-_ー]+$", "", regex=True)
            s = s.str.strip().replace("", pd.NA)
            return s

        if "partner" in df_wide.columns:
            df_wide["partner"] = vectorized_clean_partner(df_wide["partner"])
        if "debit_partner" in df_wide.columns:
            df_wide["debit_partner"] = vectorized_clean_partner(df_wide["debit_partner"])
        if "credit_partner" in df_wide.columns:
            df_wide["credit_partner"] = vectorized_clean_partner(df_wide["credit_partner"])
            
        if "description" in df_wide.columns:
            df_wide["description"] = df_wide["description"].apply(normalize_description)

        # 取引Noのクレンジング (文字列として統一し、スペース等の不要な文字を除去、欠損値は NA)
        # ※異なるファイル（年度）間で取引Noが重複するのを防ぐため、ファイル番号をプレフィックスとして付与します。
        if "transaction_no" in df_wide.columns:
            df_wide["transaction_no"] = normalize_transaction_numbers(df_wide["transaction_no"], file_num)

        # 日付前方埋め
        df_wide["date"] = df_wide["date"].ffill()
        
        # デバッグ表示
        # print("--- DEBUG: Extracted Samples (Wide Format) ---")
        # for col in df_wide.columns:
        #     print(f"DEBUG: {col}: {df_wide[col].head(5).tolist()}")

        # 有効行フィルタ
        df_wide = df_wide[df_wide["date"].notna() & (df_wide["debit_account"].notna() | df_wide["credit_account"].notna())]
        
        if df_wide.empty:
            return None, "有効データなし"

        df_wide = resolve_partner_columns(df_wide)

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
