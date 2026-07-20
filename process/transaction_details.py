import pandas as pd

from process.partner_resolution import (
    AP_ACCOUNT_PATTERN,
    AR_ACCOUNT_PATTERN,
    CASH_ACCOUNT_PATTERN,
    PURCHASE_ACCOUNT_PATTERN,
    SALES_ACCOUNT_PATTERN,
    normalize_partner_name,
    resolve_partner_columns,
)


UNKNOWN_PARTNER = "取引先不明"
NON_PARTNER_LABELS = {UNKNOWN_PARTNER, "現金売上", "掛売上", "売上", "売掛金回収"}
EXPENSE_NAME_PATTERN = r"費$|手当$|仕入高$|売上高$|製造経費$|外注加工費$|消耗費$|通信費$"


def consolidate_partner_aliases(values: pd.Series, protected_names=None) -> pd.Series:
    """会社種別等を除去し、包含関係にある名称を共通名へ寄せる。"""
    normalized = values.apply(normalize_partner_name)
    protected = {
        str(name) for name in (protected_names or set())
        if pd.notna(name) and str(name).strip()
    }
    names = sorted(
        {str(value) for value in normalized.dropna() if str(value).strip()},
        key=lambda value: (len(value), value),
    )
    aliases = {}
    for short_name in names:
        if short_name in protected:
            continue
        if len(short_name) < 3:
            continue
        for long_name in names:
            if long_name in protected:
                continue
            if short_name == long_name:
                continue
            if short_name in long_name:
                aliases[long_name] = short_name
    return normalized.replace(aliases)


def calculate_top3_share(details: pd.DataFrame) -> float | None:
    """不明額は分母だけに含め、判明先の上位3社を分子にする。"""
    if details.empty:
        return None
    amounts = details.groupby("partner")["amount"].sum().sort_values(ascending=False)
    denominator = amounts.sum()
    if denominator <= 0:
        return None
    numerator = amounts.drop(index=UNKNOWN_PARTNER, errors="ignore").head(3).sum()
    return float(numerator / denominator * 100)


def _contains(series: pd.Series, pattern: str) -> pd.Series:
    return series.fillna("").astype(str).str.contains(pattern, regex=True, na=False)


def _amount(value) -> float:
    return float(pd.to_numeric(value, errors="coerce") or 0.0)


def _tx_key(row_index, value):
    if pd.isna(value) or not str(value).strip() or str(value).lower() in {"nan", "none", "null", "<na>"}:
        return f"__row_{row_index}"
    return str(value)


def _append_allocations(records, group, candidates, total_amount, source):
    candidates = candidates.copy()
    candidates["_candidate_amount"] = pd.to_numeric(candidates["_candidate_amount"], errors="coerce").fillna(0.0)
    candidates = candidates[candidates["_candidate_amount"] > 0]
    if candidates.empty or total_amount <= 0:
        return False

    candidate_total = candidates["_candidate_amount"].sum()
    for _, candidate in candidates.iterrows():
        partner = candidate.get("_candidate_partner")
        amount = total_amount * candidate["_candidate_amount"] / candidate_total
        records.append({
            "date": candidate.get("date", group["date"].min()),
            "transaction_no": candidate.get("transaction_no"),
            "partner": partner if pd.notna(partner) else UNKNOWN_PARTNER,
            "amount": amount,
            "source": source,
        })
    return True


def build_sales_details(df: pd.DataFrame) -> pd.DataFrame:
    """売上高を、売掛金内訳優先・直入金次点で取引先別明細へ展開する。"""
    journal = resolve_partner_columns(df)
    journal["_tx_key"] = [
        _tx_key(index, value) for index, value in zip(journal.index, journal["transaction_no"])
    ]
    records = []

    for _, group in journal.groupby("_tx_key", sort=False):
        credit_sales = group[_contains(group["credit_account"], SALES_ACCOUNT_PATTERN)].copy()
        total_sales = pd.to_numeric(credit_sales.get("credit_amount"), errors="coerce").fillna(0.0).sum()
        if total_sales <= 0:
            continue  # 返品・売上取消は現方針どおり差し引かない

        ar_rows = group[_contains(group["debit_account"], AR_ACCOUNT_PATTERN)].copy()
        cash_rows = group[
            _contains(group["debit_account"], CASH_ACCOUNT_PATTERN)
            & _contains(group["credit_account"], SALES_ACCOUNT_PATTERN)
        ].copy()
        if not ar_rows.empty:
            ar_rows["_candidate_partner"] = ar_rows["ar_partner"]
            ar_rows["_candidate_amount"] = ar_rows["debit_amount"]
            if not cash_rows.empty:
                cash_rows["_candidate_partner"] = cash_rows["sales_partner"]
                cash_rows["_candidate_amount"] = cash_rows["debit_amount"]
                ar_rows = pd.concat([ar_rows, cash_rows], ignore_index=True)
                source = "売掛金・直入金内訳"
            else:
                source = "売掛金内訳"
            if _append_allocations(records, group, ar_rows, total_sales, source):
                continue

        if not cash_rows.empty:
            cash_rows["_candidate_partner"] = cash_rows["sales_partner"]
            cash_rows["_candidate_amount"] = cash_rows["debit_amount"]
            if _append_allocations(records, group, cash_rows, total_sales, "直入金売上"):
                continue

        direct = credit_sales.copy()
        direct["_candidate_partner"] = direct["sales_partner"]
        direct["_candidate_amount"] = direct["credit_amount"]
        if not _append_allocations(records, group, direct, total_sales, "売上仕訳"):
            records.append({
                "date": group["date"].min(), "transaction_no": group["transaction_no"].iloc[0],
                "partner": UNKNOWN_PARTNER, "amount": total_sales, "source": "取引先復元不能",
            })

    result = pd.DataFrame(records, columns=["date", "transaction_no", "partner", "amount", "source"])
    if not result.empty:
        result["partner"] = consolidate_partner_aliases(result["partner"]).fillna(UNKNOWN_PARTNER)
    return result


def build_purchase_details(df: pd.DataFrame) -> pd.DataFrame:
    """仕入・売上原価・外注費を取引先別明細へ展開する。"""
    journal = resolve_partner_columns(df)
    journal["_tx_key"] = [
        _tx_key(index, value) for index, value in zip(journal.index, journal["transaction_no"])
    ]
    records = []
    for _, group in journal.groupby("_tx_key", sort=False):
        purchase_rows = group[_contains(group["debit_account"], PURCHASE_ACCOUNT_PATTERN)].copy()
        total_purchase = pd.to_numeric(purchase_rows.get("debit_amount"), errors="coerce").fillna(0.0).sum()
        if total_purchase <= 0:
            continue

        # 同一行で仕入先を解決できる通常仕訳は、その金額を直接採用する。
        purchase_rows["_candidate_partner"] = purchase_rows["purchase_partner"]
        purchase_rows["_candidate_amount"] = purchase_rows["debit_amount"]
        if purchase_rows["_candidate_partner"].notna().all():
            _append_allocations(records, group, purchase_rows, total_purchase, "仕入仕訳")
            continue

        ap_rows = group[_contains(group["credit_account"], AP_ACCOUNT_PATTERN)].copy()
        ap_rows["_candidate_partner"] = ap_rows["credit_partner"]
        ap_rows["_candidate_amount"] = ap_rows["credit_amount"]
        if _append_allocations(records, group, ap_rows, total_purchase, "買掛・未払内訳"):
            continue

        _append_allocations(records, group, purchase_rows, total_purchase, "仕入仕訳")

    result = pd.DataFrame(records, columns=["date", "transaction_no", "partner", "amount", "source"])
    if not result.empty:
        result["partner"] = consolidate_partner_aliases(result["partner"]).fillna(UNKNOWN_PARTNER)
    return result


def build_direct_sales_details(df: pd.DataFrame) -> pd.DataFrame:
    """現預金／売上高の同一仕訳だけを、共通名寄せ済みで返す。"""
    journal = resolve_partner_columns(df)
    mask = _contains(journal["debit_account"], CASH_ACCOUNT_PATTERN) & _contains(
        journal["credit_account"], SALES_ACCOUNT_PATTERN
    )
    direct = journal[mask].copy()
    if direct.empty:
        return pd.DataFrame(columns=["date", "transaction_no", "partner", "amount", "source", "description", "credit_account"])
    direct["partner"] = consolidate_partner_aliases(direct["sales_partner"]).fillna(UNKNOWN_PARTNER)
    direct["amount"] = pd.to_numeric(direct["debit_amount"], errors="coerce").fillna(0.0)
    direct["source"] = direct["sales_partner_source"]
    return direct[["date", "transaction_no", "partner", "amount", "source", "description", "credit_account"]]


def build_customer_relationship_events(df: pd.DataFrame) -> pd.DataFrame:
    """売上計上額とは分離し、顧客との取引関係が確認できたイベントを返す。"""
    journal = resolve_partner_columns(df)
    sales = build_sales_details(journal).copy()
    if not sales.empty:
        sales["relationship_source"] = "売上計上"

    recovery_mask = _contains(journal["debit_account"], CASH_ACCOUNT_PATTERN) & _contains(
        journal["credit_account"], AR_ACCOUNT_PATTERN
    )
    recoveries = journal[recovery_mask].copy()
    if not recoveries.empty:
        recoveries = recoveries.assign(
            partner=recoveries["ar_partner"],
            amount=pd.to_numeric(recoveries["credit_amount"], errors="coerce").fillna(0.0),
            source=recoveries["ar_partner_source"],
            relationship_source="売掛金回収",
        )[["date", "transaction_no", "partner", "amount", "source", "relationship_source"]]

    columns = ["date", "transaction_no", "partner", "amount", "source", "relationship_source"]
    events = pd.concat([
        sales.reindex(columns=columns), recoveries.reindex(columns=columns)
    ], ignore_index=True)
    if events.empty:
        return pd.DataFrame(columns=columns)
    events["partner"] = consolidate_partner_aliases(events["partner"])
    events = events[events["partner"].notna() & ~events["partner"].isin(NON_PARTNER_LABELS)].copy()
    events["date"] = pd.to_datetime(events["date"], errors="coerce")
    events = events.dropna(subset=["date"])
    events["_event_tx"] = events["transaction_no"].astype("string")
    missing_tx = events["_event_tx"].isna() | events["_event_tx"].str.strip().str.lower().isin(
        ["", "nan", "none", "null", "<na>"]
    )
    events.loc[missing_tx, "_event_tx"] = "NO_TX_" + events.loc[missing_tx, "date"].astype(str)
    events = events.sort_values("date").drop_duplicates(["partner", "date", "_event_tx", "relationship_source"])
    return events[columns]


def resolve_payment_partner_name(row):
    """費用補助科目を支払先と誤認する場合は摘要へフォールバックする。"""
    candidate = normalize_partner_name(row.get("payment_partner"))
    debit_account = normalize_partner_name(row.get("debit_account"))
    if pd.notna(candidate):
        is_account_name = bool(pd.Series([str(candidate)]).str.contains(EXPENSE_NAME_PATTERN, regex=True).iloc[0])
        if not is_account_name and candidate != debit_account:
            return candidate
    description = normalize_partner_name(row.get("description"))
    return description if pd.notna(description) else UNKNOWN_PARTNER
