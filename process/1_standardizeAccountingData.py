# 目的と全体像：全会計データを標準化し、後続の「診断レポート作成」と「営業先リスト作成」に渡せるようにする
# 引数：ユーザーからアップロードされた各種会計データ。
# 処理：①アップロードされた会計データが正しいか判断する。②会計データの形式を標準化する。③メイン関数に返す
# 返り値：標準化された会計データ（データ形式未定。csv系か？）

# 関数構成予定：
#①会計データ正しいか判断：前提として、仕訳帳（必須）、総勘定元帳（任意）、貸借対照表（任意）、損益計算書（任意）が引数として与えられる。ただし、必須の仕訳帳は必ず引数として与えられているものとする
# 各会計データに関して、エクスポート元の会計ソフトによって項目名が若干異なる可能性がある。
# ひとまずここでは、各会計データに関して、必須で存在すると思われる項目が存在するかどうかを判断する。できれば各会計データに対して3-4つ程度の項目でデータが正しいかの✅をしてほしい
#①の処理について、もし正しくない場合は、のちの処理をスキップしてメイン画面に「会計データが正しくありません」などのエラー文を表示する。エラーの可能性はもう一つあって、仕訳帳における会計データが12-24か月以内でない場合もエラーを返すようにする。この時の文言は先ほどとは違う。
#①の処理で問題ない場合は、「会計データは正常です。診断を開始します」とメイン画面に表示したうえで、内部的には②の処理に進む

#②会計データの形式を標準化する：
# ここが非常に難しいところである。
# まず、それぞれの会計データに対して、宮田ロジックに渡すためのデータ構造（csvまたはjson的な構造）を定義するべきである。
# その後、各会計データをそのデータ構造に変換する。
# このとき例えば仕訳帳の中で「取引日」や「借方金額」といった項目があるが、これが会計ソフトによって微妙に異なる可能性がある。
# これを踏まえ、各項目に対して、表記揺れを吸収できるような処理を実装する。
# 

import pandas as pd
import io
from typing import Optional, Dict, List

# --- 定数定義 ---
STANDARD_JOURNAL_COLUMNS = [
    "date", "debit_account", "debit_amount", "debit_partner",
    "credit_account", "credit_amount", "credit_partner", "created_at"
]

# 表記ゆれ吸収用マッパー
HEADER_MAPPING = {
    "date": ["日付", "年月日", "取引日", "発生日", "取引日付", "日付 yyyy/MM/dd（必須）", "伝票日付"],
    "debit_account": ["借方科目", "借方科目名", "借方科目名称", "借方勘定科目"],
    "debit_amount": ["借方金額", "借方金額(円)", "借方金額（必須）", "金額"],
    "debit_partner": ["借方取引先", "摘要", "取引摘要", "行摘要"],
    "credit_account": ["貸方科目", "貸方科目名", "貸方科目名称", "貸方勘定科目"],
    "credit_amount": ["貸方金額", "貸方金額(円)", "貸方金額（必須）", "金額"],
    "credit_partner": ["貸方取引先", "摘要", "取引摘要", "行摘要"],
    "created_at": ["作成日時", "入力日付時間", "入力日付", "登録日時", "入力日", "入力日時", "仕分日", "仕分日時", "Created At"],
}

def validate_accounting_data(df: pd.DataFrame) -> bool:
    """
    会計データが最低限必要な情報（日付、科目、金額）を持っているか判定する。
    """
    required = ["date", "debit_account", "debit_amount", "credit_account", "credit_amount"]
    return all(col in df.columns for col in required)

def _map_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    マスタに基づいてヘッダーを標準名に変換する。
    重複するマッピングが発生した場合は、最初に一致したものを優先する。
    """
    rename_dict = {}
    used_std_names = set()
    
    # 完全に一致するものを優先してマッピング
    cols = list(df.columns)
    for col in cols:
        col_str = str(col).strip()
        # pandasの重複カラム名「.1」「.2」等を除去して比較
        import re
        base_col = re.sub(r'\.\d+$', '', col_str)
        
        for std_name, aliases in HEADER_MAPPING.items():
            if std_name in used_std_names:
                continue
            if col_str in aliases or base_col in aliases:
                rename_dict[col] = std_name
                used_std_names.add(std_name)
                break
                
    # 部分一致やより柔軟な判定（金額や科目など、必須項目がまだ見つからない場合）
    for col in cols:
        if col in rename_dict: continue
        col_str = str(col).strip()
        
        for std_name, aliases in HEADER_MAPPING.items():
            if std_name in used_std_names:
                continue
            # 非常に限定的な部分一致（誤爆を防ぐため）
            if std_name == "debit_amount" and ("借方" in col_str and "金額" in col_str):
                rename_dict[col] = std_name
                used_std_names.add(std_name)
                break
            if std_name == "credit_amount" and ("貸方" in col_str and "金額" in col_str):
                rename_dict[col] = std_name
                used_std_names.add(std_name)
                break
                
    return df.rename(columns=rename_dict)

def _find_header_row(file_journal: io.BytesIO, encoding: str) -> int:
    """
    「日付」「借方」「貸方」などのキーワードを含む行を探し、ヘッダー行のインデックスを返す。
    見つからない場合は0を返す。
    """
    file_journal.seek(0)
    try:
        # 最初の10行程度を確認
        df_top = pd.read_csv(file_journal, nrows=10, header=None, encoding=encoding)
        for i, row in df_top.iterrows():
            row_str = "".join(row.astype(str).values)
            if ("日付" in row_str or "年月" in row_str or "取引日" in row_str) and \
               ("借方" in row_str or "科目" in row_str) and \
               ("金額" in row_str):
                return i
    except:
        pass
    return 0

def _flatten_journal(df: pd.DataFrame) -> pd.DataFrame:
    """
    借方・貸方に分かれた仕訳データを、1レコード1科目のフラットな形式に変換する。
    """
    # 貸借分離型（借方科目/貸方科目の列がある場合）か判定
    if "debit_account" in df.columns and "credit_account" in df.columns:
        # 借方データを抽出
        debit_df = df.copy()
        debit_df["account"] = debit_df["debit_account"]
        debit_df["amount"] = debit_df["debit_amount"]
        debit_df["side"] = "debit"
        
        # 貸方データを抽出
        credit_df = df.copy()
        credit_df["account"] = credit_df["credit_account"]
        credit_df["amount"] = credit_df["credit_amount"]
        credit_df["side"] = "credit"
        
        # 結合
        flat_df = pd.concat([debit_df, credit_df], ignore_index=True)
    else:
        # すでにフラットな場合や不明な場合はそのまま（sideなしならデフォルトdebitとする等）
        flat_df = df.copy()
        if "side" not in flat_df.columns:
            flat_df["side"] = "unknown"
            
    return flat_df[flat_df["account"].notna()]

def check_accounting_files(file_journal: io.BytesIO, file_bs: Optional[io.BytesIO] = None) -> List[Dict]:
    """
    会計データのバリデーションを行う。
    結果のリスト（メッセージ、色、続行可否）を返す。
    """
    results = []
    
    # --- 1. 仕訳帳の確認 ---
    try:
        # A: 1-2行目のキーワードチェック
        file_journal.seek(0)
        df_head = None
        for enc in ['utf-8-sig', 'cp932', 'utf-8', 'shift_jis']:
            try:
                file_journal.seek(0)
                df_head = pd.read_csv(file_journal, nrows=2, header=None, encoding=enc)
                break
            except:
                continue
        
        if df_head is None:
            results.append({
                "message": "仕訳帳の読み込みに失敗しました（文字コードエラー）",
                "color": "red",
                "success": False
            })
            return results
        
        # 全セルをリスト化
        head_text = df_head.astype(str).values.flatten().tolist()
        
        # 借方・貸方チェック（既存のロジックを維持：部分一致）
        head_combined = "".join(head_text)
        has_debit = "借方" in head_combined
        has_credit = "貸方" in head_combined
        
        # 取引日チェック（新規追加：完全一致）
        date_aliases = HEADER_MAPPING["date"]
        has_date = any(alias in head_text for alias in date_aliases)
        
        if not (has_debit and has_credit and has_date):
            missing = []
            if not (has_debit and has_credit): missing.append("「借方」「貸方」")
            if not has_date: missing.append("「日付」関連（取引日、年月日等）")
            
            results.append({
                "message": f"仕訳帳のヘッダー（1〜2行目）に必要な項目（{', '.join(missing)}）が含まれていません",
                "color": "red",
                "success": False
            })
            return results

        # B: 期間チェック (12-24ヶ月)
        file_journal.seek(0)
        df_journal = None
        for enc in ['utf-8-sig', 'cp932', 'utf-8', 'shift_jis']:
            try:
                file_journal.seek(0)
                df_journal = pd.read_csv(file_journal, encoding=enc)
                break
            except:
                continue

        if df_journal is None:
            results.append({
                "message": "仕訳帳の読み込みに失敗しました（文字コードエラー）",
                "color": "red",
                "success": False
            })
            return results

        # 日付列を検索（HEADER_MAPPING["date"]のいずれかと完全一致）
        date_col = None
        date_aliases = HEADER_MAPPING["date"]
        for col in df_journal.columns:
            if str(col) in date_aliases:
                date_col = col
                break
        
        if date_col is None:
            results.append({
                "message": "仕訳帳に有効な日付列（日付、取引日、年月日等）が見つかりません",
                "color": "red",
                "success": False
            })
            return results

        df_journal[date_col] = pd.to_datetime(df_journal[date_col], errors='coerce')
        df_journal = df_journal[df_journal[date_col].notna()]
        
        if df_journal.empty:
            results.append({
                "message": "仕訳帳の「取引日」列に有効な日付データが見つかりませんでした",
                "color": "red",
                "success": False
            })
            return results
            
        min_date = df_journal[date_col].min()
        max_date = df_journal[date_col].max()
        
        # 月数の計算
        months = (max_date.year - min_date.year) * 12 + (max_date.month - min_date.month)
        
        # ユーザー要件: 12か月以上24か月以内
        if 11 <= months <= 23:
            results.append({
                "message": "仕訳帳は正常です。処理を続けます",
                "color": "green",
                "success": True
            })
        else:
            results.append({
                "message": f"仕訳帳のデータ期間が12ヶ月〜24ヶ月の範囲外です（現在の期間: {months}ヶ月）",
                "color": "red",
                "success": False
            })
            return results

    except Exception as e:
        results.append({
            "message": f"仕訳帳のチェック中に想定外のエラーが発生しました: {str(e)}",
            "color": "red",
            "success": False
        })
        return results

    # --- 2. 貸借対照表の確認 ---
    if file_bs is None:
        results.append({
            "message": "貸借対照表は無しで分析を開始します",
            "color": "black",
            "success": True
        })
    else:
        try:
            file_bs.seek(0)
            df_bs = None
            for enc in ['utf-8-sig', 'cp932', 'utf-8', 'shift_jis']:
                try:
                    file_bs.seek(0)
                    df_bs = pd.read_csv(file_bs, header=None, encoding=enc)
                    break
                except:
                    continue

            if df_bs is None:
                results.append({
                    "message": "貸借対照表の読み込みに失敗しました（文字コードエラー）",
                    "color": "red",
                    "success": False
                })
                return results
            
            # 全データの中から「現金」または「預金」を探す
            bs_combined = "".join(df_bs.astype(str).values.flatten().tolist())
            
            if "現金" in bs_combined or "預金" in bs_combined:
                results.append({
                    "message": "貸借対照表は正常です。分析を開始します",
                    "color": "green",
                    "success": True
                })
            else:
                results.append({
                    "message": "貸借対照表には少なくとも、「現金」「普通預金」「当座預金」「定期預金」のいずれかが含まれている必要があります",
                    "color": "red",
                    "success": False
                })
                return results
        except Exception as e:
            results.append({
                "message": f"貸借対照表のチェック中にエラーが発生しました: {str(e)}",
                "color": "red",
                "success": False
            })
            return results

    return results

def _extract_bs_data(file_bs: io.BytesIO) -> pd.DataFrame:
    """
    貸借対照表から期末残高を2次元検索で抽出する。
    """
    try:
        file_bs.seek(0)
        df_bs = None
        for enc in ['utf-8-sig', 'cp932', 'utf-8', 'shift_jis']:
            try:
                file_bs.seek(0)
                df_bs = pd.read_csv(file_bs, header=None, encoding=enc)
                break
            except:
                continue

        if df_bs is None:
            return pd.DataFrame()

        # 1-3行目から「期末」が含まれる列を特定（一番右を優先）
        target_col_idx = -1
        for r in range(min(3, len(df_bs))):
            row_data = df_bs.iloc[r]
            for c_idx, val in enumerate(row_data):
                if "期末" in str(val):
                    target_col_idx = max(target_col_idx, c_idx)
        
        if target_col_idx == -1:
            return pd.DataFrame()

        # 1-3列目から勘定科目を特定
        accounts = {
            "期末現金": ["現金"],
            "期末普通預金": ["普通預金"],
            "期末当座預金": ["当座預金"],
            "期末定期預金": ["定期預金"]
        }
        
        results = {}
        for key, keywords in accounts.items():
            value = 0
            found = False
            for r_idx in range(len(df_bs)):
                # 1-3列目をチェック
                for c_idx in range(min(3, len(df_bs.columns))):
                    cell_val = str(df_bs.iloc[r_idx, c_idx])
                    if any(kw in cell_val for kw in keywords):
                        # 期末列の値を取得
                        try:
                            raw_val = str(df_bs.iloc[r_idx, target_col_idx]).replace(',', '')
                            import re
                            # 数字以外の文字を除去（ただしマイナスは残す可能性を考慮しつつ基本は数値抽出）
                            num_str = re.sub(r'[^\d.-]', '', raw_val)
                            value = float(num_str) if num_str else 0
                        except:
                            value = 0
                        found = True
                        break
                if found: break
            results[key] = value

        # 合計の計算
        results["期末現預金合計"] = sum([results["期末現金"], results["期末普通預金"], results["期末当座預金"], results["期末定期預金"]])
        
        return pd.DataFrame([results])

    except Exception as e:
        print(f"BS extraction error: {e}")
        return pd.DataFrame()

def standardize_logic(file_journal: io.BytesIO, file_ledger: Optional[io.BytesIO] = None, 
                      file_bs: Optional[io.BytesIO] = None, file_pl: Optional[io.BytesIO] = None) -> Dict[str, pd.DataFrame]:
    """
    アップロードされたファイルを読み込み、標準化されたDataFrameの辞書を返す。
    """
    # 1. ヘッダー行の自動判定と読み込み
    df_j = None
    for enc in ['utf-8-sig', 'cp932', 'utf-8', 'shift_jis']:
        try:
            skip_rows = _find_header_row(file_journal, enc)
            file_journal.seek(0)
            df_j = pd.read_csv(file_journal, encoding=enc, skiprows=skip_rows)
            
            # 正常に読み込めたか、最低限の項目チェック
            mapped_temp = _map_headers(df_j)
            if "date" in mapped_temp.columns and ("debit_amount" in mapped_temp.columns or "credit_amount" in mapped_temp.columns):
                break # 成功
            else:
                # 判定に失敗した場合はskiprows=0で再試行してみる
                file_journal.seek(0)
                df_j = pd.read_csv(file_journal, encoding=enc)
                mapped_temp = _map_headers(df_j)
                if "date" in mapped_temp.columns:
                    break
        except Exception as e:
            print(f"Read error with {enc}: {e}")
            continue

    if df_j is None or df_j.empty:
        raise ValueError("仕訳帳の読み込みに失敗しました。ファイルが空か、対応していない形式です。")
    
    # 2. ヘッダーの名寄せ
    df_j = _map_headers(df_j)
    
    # 3. 必要項目の抽出
    # マッピング結果から存在するカラムのみ、さらに指定の名称に統一
    final_j_cols = {}
    for std_name in STANDARD_JOURNAL_COLUMNS:
        if std_name in df_j.columns:
            final_j_cols[std_name] = df_j[std_name]
        else:
            final_j_cols[std_name] = pd.NA
            
    df_journal_std = pd.DataFrame(final_j_cols)

    # 4. 貸借対照表の処理
    df_bs_std = pd.DataFrame()
    if file_bs:
        df_bs_std = _extract_bs_data(file_bs)
    
    return {
        "journal": df_journal_std,
        "bs": df_bs_std
    }

if __name__ == "__main__":
    # テスト用
    print("Standardization logic module loaded.")