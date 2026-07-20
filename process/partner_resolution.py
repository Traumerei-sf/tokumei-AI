import re
import unicodedata

import pandas as pd


CASH_ACCOUNT_PATTERN = r"預金|現金|当座|普通|定期|別段|手形|電信"
SALES_ACCOUNT_PATTERN = r"売上"
PURCHASE_ACCOUNT_PATTERN = r"仕入|売上原価|外注"
AR_ACCOUNT_PATTERN = r"売掛|未収|買入金銭債権"
AP_ACCOUNT_PATTERN = r"買掛|未払"
GENERIC_PARTNER_NAMES = {"諸口", "摘要", "取引先", "不明", "なし", "空欄"}
CORPORATE_MARKER_PATTERN = (
    r"株式会社|有限会社|合資会社|合名会社|合同会社|医療法人|社会福祉法人|"
    r"一般社団法人|公益社団法人|一般財団法人|公益財団法人|NPO法人|"
    r"\(株\)|（株）|\(有\)|（有）|㈱|㈲|"
    r"(?:\(カ\)|（カ）|カ\)|\(カ|（カ|カ）|\(ユ\)|（ユ）|ユ\)|\(ユ|（ユ|ユ）)|"
    r"カブシキガイシャ|ユウゲンガイシャ|ゴウドウガイシャ|"
    r"(?:CO\.?[, ]*LTD|CORPORATION|INC\.?)"
)
BANK_DESCRIPTION_PATTERN = (
    r"銀行|信用金庫|信用組合|信金|信組|農協|JA|"
    r"ミツビシ|ミツイスミトモ|エスビーアイ|ユウチョ|ラクテン|"
    r"ヨコハマ|エヒメシンキン|GMO|振込|フリコミ|ﾌﾘｺﾐ"
)
RESOLVED_PARTNER_COLUMNS = tuple(
    f"{name}_partner{suffix}"
    for name in ("sales", "purchase", "ar", "payment")
    for suffix in ("", "_source", "_kind")
)


def normalize_description(value):
    """摘要は意味を削らず、Unicodeと空白だけを整える。"""
    if pd.isna(value) or not str(value).strip():
        return pd.NA
    return unicodedata.normalize("NFKC", str(value)).strip()


def normalize_partner_name(value):
    """分析キーとして使う取引先名を保守的に正規化する。"""
    value = normalize_description(value)
    if pd.isna(value):
        return pd.NA
    text = str(value)
    text = re.sub(r"^(?:振込|フリコミ|振込口|ネット)", "", text)
    corporate_patterns = (
        r"株式会社|有限会社|合資会社|合名会社|合同会社|法人|"
        r"\(株\)|（株）|\(有\)|（有）|\(合\)|（合）|㈱|㈲|"
        r"\(カ\)|（カ）|カ\)|\(カ|（カ|カ）|\(ユ\)|（ユ）|ユ\)|\(ユ|（ユ|ユ）|"
        r"カブシキガイシャ|ユウゲンガイシャ|ゴウドウガイシャ"
    )
    text = re.sub(corporate_patterns, "", text)
    text = re.sub(r"[\s　]+", "", text)
    text = re.sub(r"^[.\-_ー]+|[.\-_ー]+$", "", text)
    if text.lower() in GENERIC_PARTNER_NAMES:
        return pd.NA
    return text if text else pd.NA


def _text(value) -> str:
    return "" if pd.isna(value) else str(value)


def _contains(value, pattern: str) -> bool:
    return bool(re.search(pattern, _text(value)))


def _valid_candidate(value, account, reject_cash=False):
    if pd.isna(value) or not str(value).strip():
        return None
    if reject_cash and _contains(account, CASH_ACCOUNT_PATTERN):
        return None
    return normalize_partner_name(value)


def _first_candidate(candidates, fallback_source=""):
    for value, source in candidates:
        if value is not None and not pd.isna(value) and str(value).strip():
            return value, source
    return pd.NA, fallback_source


def _unique_group_candidate(group: pd.DataFrame, conditions):
    candidates = []
    for mask_col, column, source in conditions:
        for value in group.loc[group[mask_col], column].dropna():
            normalized = normalize_partner_name(value)
            if pd.notna(normalized) and normalized not in candidates:
                candidates.append(normalized)
    if len(candidates) == 1:
        return candidates[0], source
    return pd.NA, ""


def _party_kind(raw_value) -> str:
    """法人表記が明示されたものだけ法人確定とし、それ以外は不明扱いにする。"""
    raw = normalize_description(raw_value)
    if pd.notna(raw) and re.search(CORPORATE_MARKER_PATTERN, str(raw), flags=re.IGNORECASE):
        return "corporate"
    return "unknown"


def _known_partner_catalog(result: pd.DataFrame) -> dict[str, str]:
    """補助科目に明示された債権債務先を、銀行摘要照合用の既知先として集める。"""
    catalog = {}
    conditions = (
        (result["debit_account"].astype(str).str.contains(AR_ACCOUNT_PATTERN, na=False), "debit_partner"),
        (result["credit_account"].astype(str).str.contains(AR_ACCOUNT_PATTERN, na=False), "credit_partner"),
        (result["debit_account"].astype(str).str.contains(AP_ACCOUNT_PATTERN, na=False), "debit_partner"),
        (result["credit_account"].astype(str).str.contains(AP_ACCOUNT_PATTERN, na=False), "credit_partner"),
    )
    corporate_names = set()
    for column in ("description_raw", "description", "partner_raw"):
        if column not in result.columns:
            continue
        for raw in result[column].dropna():
            if _party_kind(raw) == "corporate":
                normalized = normalize_partner_name(raw)
                if pd.notna(normalized):
                    corporate_names.add(str(normalized))
    for mask, column in conditions:
        raw_column = f"{column}_raw" if f"{column}_raw" in result.columns else column
        selected = result.loc[mask & result[column].notna()]
        for raw in selected[raw_column]:
            normalized = normalize_partner_name(raw)
            if pd.isna(normalized) or len(str(normalized)) < 3:
                continue
            kind = "corporate" if str(normalized) in corporate_names else _party_kind(raw)
            previous = catalog.get(str(normalized))
            if previous != "corporate" or kind == "corporate":
                catalog[str(normalized)] = kind
    # 銀行の定型項目（銀行・支店・口座番号）が揃う摘要は、それ自体を安全な既知先にできる。
    ar_or_ap = (
        result["debit_account"].astype(str).str.contains(f"{AR_ACCOUNT_PATTERN}|{AP_ACCOUNT_PATTERN}", na=False)
        | result["credit_account"].astype(str).str.contains(f"{AR_ACCOUNT_PATTERN}|{AP_ACCOUNT_PATTERN}", na=False)
    )
    for raw in result.loc[ar_or_ap, "description"].dropna().drop_duplicates():
        normalized = normalize_partner_name(raw)
        candidate, source, kind = _extract_structured_bank_counterparty(raw, normalized)
        if source == "銀行摘要→定型項目除去":
            catalog[str(candidate)] = kind
    return catalog


def _bank_partner_suffix_index(catalog: dict[str, str]) -> dict[str, list[tuple[str, str]]]:
    index = {}
    for known, kind in catalog.items():
        index.setdefault(known[-3:], []).append((known, kind))
    for values in index.values():
        values.sort(key=lambda item: len(item[0]), reverse=True)
    return index


def _extract_known_partner_from_bank_description(value, suffix_index):
    """銀行情報付き摘要の末尾を既知先と完全照合する。推測だけでは置換しない。"""
    raw = normalize_description(value)
    normalized = normalize_partner_name(raw)
    if pd.isna(normalized):
        return normalized, "摘要", "unknown"
    if not suffix_index:
        return _extract_structured_bank_counterparty(raw, normalized)
    text = str(normalized)
    matches = []
    for known, kind in suffix_index.get(text[-3:], ()):
        if text == known:
            matches.append((known, kind, ""))
        elif text.endswith(known):
            matches.append((known, kind, text[:-len(known)]))
    matches = [item for item in matches if not item[2] or re.search(BANK_DESCRIPTION_PATTERN, item[2], flags=re.IGNORECASE)]
    if not matches:
        return _extract_structured_bank_counterparty(raw, normalized)
    longest = max(len(item[0]) for item in matches)
    winners = [item for item in matches if len(item[0]) == longest]
    if len(winners) != 1:
        return normalized, "摘要", "unknown"
    known, kind, prefix = winners[0]
    if not prefix:
        return known, "摘要（既知取引先完全一致）", kind
    resolved_kind = kind if kind == "corporate" else "individual"
    return known, "銀行摘要→既知取引先完全一致", resolved_kind


def _extract_structured_bank_counterparty(raw, normalized):
    """口座番号等を含む定型銀行摘要だけ、銀行情報を除いて相手名を抽出する。"""
    if pd.isna(raw):
        return normalized, "摘要", "unknown"
    text = unicodedata.normalize("NFKC", str(raw)).strip()
    head = re.split(r"\((?:依頼人名|振込予定|管理番号)|（(?:依頼人名|振込予定|管理番号)", text, maxsplit=1)[0]
    tokens = [token for token in re.split(r"[\s　]+", head) if token]
    is_transfer_description = bool(tokens and re.fullmatch(r"振込|フリコミ|ﾌﾘｺﾐ", tokens[0]))
    candidates = []
    removed_bank_context = False
    has_account_context = False
    for token in tokens:
        if re.search(BANK_DESCRIPTION_PATTERN, token, flags=re.IGNORECASE):
            removed_bank_context = True
            continue
        if re.search(r"営業部|支店|出張所|普通預金|当座預金|貯蓄預金", token):
            removed_bank_context = True
            has_account_context = True
            continue
        if re.fullmatch(r"\d{5,}", token):
            removed_bank_context = True
            has_account_context = True
            continue
        candidates.append(token)
    if not removed_bank_context or not candidates:
        return normalized, "摘要", "unknown"
    is_corporate = _party_kind(head) == "corporate"
    has_representative_label = bool(re.search(r"代表|ダイヒョウ|ダイヒヨウ", head))
    # 個人の銀行摘要は「銀行略称 姓 名」になりやすい。構造語を除去した後に
    # 3トークン以上残る場合だけ、末尾の姓・名を採用する。
    if len(candidates) >= 3 and not is_corporate and not has_representative_label:
        candidates = candidates[-2:]
        has_account_context = has_account_context or is_transfer_description
    if not has_account_context:
        return normalized, "摘要", "unknown"
    candidate_raw = "".join(candidates)
    candidate = normalize_partner_name(candidate_raw)
    if pd.isna(candidate) or len(str(candidate)) < 3:
        return normalized, "摘要", "unknown"
    kind = "corporate" if is_corporate else "individual"
    return candidate, "銀行摘要→定型項目除去", kind


def resolve_partner_columns(df: pd.DataFrame, force: bool = False) -> pd.DataFrame:
    """生の摘要・補助科目・勘定科目・取引Noから用途別取引先を生成する。"""
    # 2年度のDataFrameをconcatした場合も、重複indexで別仕訳を誤更新しないよう振り直す。
    result = df.copy().reset_index(drop=True)
    if not force and all(column in result.columns for column in RESOLVED_PARTNER_COLUMNS):
        return result
    for column in ("partner", "debit_partner", "credit_partner", "transaction_no"):
        if column not in result.columns:
            result[column] = pd.NA
    if "description" not in result.columns:
        result["description"] = result["partner"]
    result["description"] = result["description"].apply(normalize_description)
    known_catalog = _known_partner_catalog(result)
    bank_suffix_index = _bank_partner_suffix_index(known_catalog)

    resolved_rows = []
    # Seriesを行ごとに生成するiterrowsは大規模元帳で重いため、辞書レコードを使う。
    for row in result.to_dict("records"):
        debit_account = row.get("debit_account")
        credit_account = row.get("credit_account")
        debit_partner = _valid_candidate(row.get("debit_partner"), debit_account, reject_cash=True)
        credit_partner = _valid_candidate(row.get("credit_partner"), credit_account, reject_cash=True)
        description_partner = normalize_partner_name(row.get("description"))
        description_source = "摘要"
        description_kind = _party_kind(row.get("description"))
        is_bank_entry = _contains(debit_account, CASH_ACCOUNT_PATTERN) or _contains(credit_account, CASH_ACCOUNT_PATTERN)
        if is_bank_entry:
            description_partner, description_source, description_kind = _extract_known_partner_from_bank_description(
                row.get("description"), bank_suffix_index
            )
        legacy_partner = normalize_partner_name(row.get("partner"))

        if _contains(debit_account, AR_ACCOUNT_PATTERN):
            ar = _first_candidate(((debit_partner, "借方AR補助科目"), (description_partner, description_source), (legacy_partner, "互換partner")))
        elif _contains(credit_account, AR_ACCOUNT_PATTERN):
            ar = _first_candidate(((credit_partner, "貸方AR補助科目"), (description_partner, description_source), (legacy_partner, "互換partner")))
        else:
            ar = (pd.NA, "")

        if _contains(credit_account, SALES_ACCOUNT_PATTERN):
            sales = _first_candidate(((debit_partner, "売上相手側補助科目"), (credit_partner, "売上側補助科目"), (description_partner, description_source), (legacy_partner, "互換partner")))
        elif _contains(debit_account, SALES_ACCOUNT_PATTERN):
            sales = _first_candidate(((credit_partner, "売上相手側補助科目"), (debit_partner, "売上側補助科目"), (description_partner, description_source), (legacy_partner, "互換partner")))
        else:
            sales = (pd.NA, "")

        if _contains(debit_account, PURCHASE_ACCOUNT_PATTERN):
            purchase = _first_candidate(((credit_partner, "仕入相手側補助科目"), (debit_partner, "仕入側補助科目"), (description_partner, description_source), (legacy_partner, "互換partner")))
        elif _contains(credit_account, PURCHASE_ACCOUNT_PATTERN):
            purchase = _first_candidate(((debit_partner, "仕入相手側補助科目"), (credit_partner, "仕入側補助科目"), (description_partner, description_source), (legacy_partner, "互換partner")))
        else:
            purchase = (pd.NA, "")

        if _contains(credit_account, CASH_ACCOUNT_PATTERN):
            payment = _first_candidate(((debit_partner, "支払先側補助科目"), (description_partner, description_source), (legacy_partner, "互換partner")))
        elif _contains(debit_account, CASH_ACCOUNT_PATTERN):
            payment = _first_candidate(((credit_partner, "入金元側補助科目"), (description_partner, description_source), (legacy_partner, "互換partner")))
        else:
            payment = _first_candidate(((debit_partner, "借方補助科目"), (credit_partner, "貸方補助科目"), (description_partner, description_source), (legacy_partner, "互換partner")))

        resolved_rows.append((sales, purchase, ar, payment, description_kind))

    for index, name in enumerate(("sales", "purchase", "ar", "payment")):
        result[f"{name}_partner"] = [row[index][0] for row in resolved_rows]
        result[f"{name}_partner_source"] = [row[index][1] for row in resolved_rows]
        result[f"{name}_partner_kind"] = [
            row[4] if row[index][1].startswith(("摘要", "銀行摘要")) else "unknown"
            for row in resolved_rows
        ]

    valid_tx = result["transaction_no"].notna() & ~result["transaction_no"].astype(str).str.lower().isin(("", "nan", "none", "null", "<na>"))
    
    # 事前計算 (高速化のため)
    result["_is_debit_ar"] = result["debit_account"].astype(str).str.contains(AR_ACCOUNT_PATTERN, na=False)
    result["_is_credit_ar"] = result["credit_account"].astype(str).str.contains(AR_ACCOUNT_PATTERN, na=False)
    result["_is_credit_ap"] = result["credit_account"].astype(str).str.contains(AP_ACCOUNT_PATTERN, na=False)
    result["_is_debit_sales"] = result["debit_account"].astype(str).str.contains(SALES_ACCOUNT_PATTERN, na=False)
    result["_is_credit_sales"] = result["credit_account"].astype(str).str.contains(SALES_ACCOUNT_PATTERN, na=False)
    result["_is_debit_purchase"] = result["debit_account"].astype(str).str.contains(PURCHASE_ACCOUNT_PATTERN, na=False)
    result["_is_credit_purchase"] = result["credit_account"].astype(str).str.contains(PURCHASE_ACCOUNT_PATTERN, na=False)

    # 同一取引No補完が必要なのは複数行仕訳だけ。単一行取引のgroupbyを避ける。
    compound_tx = valid_tx & result["transaction_no"].duplicated(keep=False)
    for _, group in result[compound_tx].groupby("transaction_no", sort=False):
        sales_group = _unique_group_candidate(group, (("_is_debit_ar", "debit_partner", "同一取引NoのAR借方補助科目"),))
        purchase_group = _unique_group_candidate(group, (("_is_credit_ap", "credit_partner", "同一取引Noの買掛・未払補助科目"),))
        ar_group = _unique_group_candidate(group, (
            ("_is_debit_ar", "debit_partner", "同一取引NoのAR補助科目"),
            ("_is_credit_ar", "credit_partner", "同一取引NoのAR補助科目"),
        ))
        for column, candidate in (("sales", sales_group), ("purchase", purchase_group), ("ar", ar_group)):
            value, source = candidate
            if pd.isna(value):
                continue
            if column == "sales":
                relevant = group["_is_debit_sales"] | group["_is_credit_sales"]
                low_confidence = group["sales_partner_source"] != "売上相手側補助科目"
                target_rows = group.index[relevant & low_confidence]
            elif column == "purchase":
                relevant = group["_is_debit_purchase"] | group["_is_credit_purchase"]
                low_confidence = group["purchase_partner_source"] != "仕入相手側補助科目"
                target_rows = group.index[relevant & low_confidence]
            else:
                target_rows = group.index[group["ar_partner"].isna()]
            result.loc[target_rows, f"{column}_partner"] = value
            result.loc[target_rows, f"{column}_partner_source"] = source
            result.loc[target_rows, f"{column}_partner_kind"] = "unknown"

    # 事前計算したフラグ列の削除
    result.drop(columns=["_is_debit_ar", "_is_credit_ar", "_is_credit_ap", "_is_debit_sales", "_is_credit_sales", "_is_debit_purchase", "_is_credit_purchase"], inplace=True)

    return result
