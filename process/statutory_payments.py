import pandas as pd


CATEGORIES = {
    "源泉所得税": {"keywords": ("所得税", "源泉"), "accounts": ("預り金", "法定福利費")},
    "住民税": {"keywords": ("住民税", "特別徴収", "市県民税"), "accounts": ("預り金", "法定福利費")},
    "社会保険料": {"keywords": ("社会保険", "健康保険", "厚生年金"), "accounts": ("預り金", "法定福利費")},
}


def _contains_any(value, keywords) -> bool:
    return any(keyword in str(value) for keyword in keywords) if pd.notna(value) else False


def _has_bank_credit(row, journal) -> bool:
    bank_words = ("預金", "現金", "当座", "普通", "手形", "電信")
    if _contains_any(row.get("credit_account"), bank_words):
        return True
    tx = row.get("transaction_no")
    if pd.isna(tx) or not str(tx).strip() or str(tx).lower() in {"nan", "none", "null", "<na>"}:
        return False
    related = journal[journal["transaction_no"] == tx]
    return related["credit_account"].apply(lambda value: _contains_any(value, bank_words)).any()


def _matches_category(row, config, side: str) -> bool:
    account = row.get(f"{side}_account")
    partner = row.get(f"{side}_partner")
    if not _contains_any(account, config["accounts"]):
        return False
    # 補助科目が明示されている場合はそれを優先する。摘要には別項目名が併記されることがあり、
    # 例: 借方「所得税」／貸方「雇用保険」、摘要「所得税 雇用保険」。
    if pd.notna(partner) and str(partner).strip():
        return _contains_any(partner, config["keywords"])
    return _contains_any(row.get("description"), config["keywords"])


def _transaction_key(index, value):
    if pd.isna(value) or not str(value).strip() or str(value).lower() in {"nan", "none", "null", "<na>"}:
        return f"__row_{index}"
    return str(value)


def _settlement_date(group: pd.DataFrame, journal: pd.DataFrame):
    """直接銀行決済、または未払金等へ振替後の銀行決済日を返す。"""
    bank_words = ("預金", "現金", "当座", "普通", "手形", "電信")
    if group["credit_account"].apply(lambda value: _contains_any(value, bank_words)).any():
        return group["date"].max()

    clearing_mask = group["credit_account"].fillna("").astype(str).str.contains("未払金|未払費用", regex=True)
    if not clearing_mask.any():
        return None
    clearing_total = pd.to_numeric(group.loc[clearing_mask, "credit_amount"], errors="coerce").fillna(0.0).sum()
    clearing_accounts = group.loc[clearing_mask, "credit_account"].dropna().astype(str).unique()
    start = group["date"].max()
    end = start + pd.Timedelta(days=10)
    later = journal[(journal["date"] >= start) & (journal["date"] <= end)]
    for account in clearing_accounts:
        candidates = later[
            later["debit_account"].fillna("").astype(str).str.contains(account, regex=False)
            & later["credit_account"].apply(lambda value: _contains_any(value, bank_words))
        ]
        for _, candidate in candidates.iterrows():
            amount = pd.to_numeric(candidate.get("debit_amount"), errors="coerce")
            if pd.notna(amount) and abs(float(amount) - clearing_total) <= 1.0:
                return candidate["date"]
    return None


def _due_date(category: str, occurrence_date: pd.Timestamp, withholding_special: bool) -> pd.Timestamp:
    if category == "源泉所得税" and withholding_special:
        if occurrence_date.month <= 6:
            return pd.Timestamp(occurrence_date.year, 7, 20)
        return pd.Timestamp(occurrence_date.year + 1, 1, 31)
    if category == "社会保険料":
        return occurrence_date.to_period("M").start_time + pd.DateOffset(months=2, days=4)
    return occurrence_date.to_period("M").start_time + pd.DateOffset(months=1, days=14)


def build_statutory_payment_ledger(journal: pd.DataFrame):
    """税・社会保険を項目別にFIFO消込し、未納額と期限超過を返す。"""
    if journal.empty:
        return pd.DataFrame(), False
    data = journal.copy()
    data["date"] = pd.to_datetime(data["date"], errors="coerce")
    data = data.dropna(subset=["date"]).sort_values("date")

    data["_tx_key"] = [_transaction_key(index, value) for index, value in zip(data.index, data["transaction_no"])]
    grouped_transactions = list(data.groupby("_tx_key", sort=False))

    payment_counts = 0
    for _, group in grouped_transactions:
        config = CATEGORIES["源泉所得税"]
        debit = pd.to_numeric(group.apply(lambda row: row.get("debit_amount", 0) if _matches_category(row, config, "debit") else 0, axis=1), errors="coerce").fillna(0.0).sum()
        credit = pd.to_numeric(group.apply(lambda row: row.get("credit_amount", 0) if _matches_category(row, config, "credit") else 0, axis=1), errors="coerce").fillna(0.0).sum()
        if debit > credit and _settlement_date(group, data) is not None:
            payment_counts += 1
    month_span = max(1, (data["date"].max().year - data["date"].min().year) * 12
                     + data["date"].max().month - data["date"].min().month + 1)
    withholding_special = month_span >= 6 and payment_counts / month_span <= 2 / 12

    ledger_rows = []
    for category, config in CATEGORIES.items():
        occurrences = []
        payments = []
        for _, group in grouped_transactions:
            debit = pd.to_numeric(
                group.apply(lambda row: row.get("debit_amount", 0) if _matches_category(row, config, "debit") else 0, axis=1),
                errors="coerce",
            ).fillna(0.0).sum()
            credit = pd.to_numeric(
                group.apply(lambda row: row.get("credit_amount", 0) if _matches_category(row, config, "credit") else 0, axis=1),
                errors="coerce",
            ).fillna(0.0).sum()
            net = float(credit - debit)
            event_date = group["date"].min()
            if net > 0:
                occurrences.append({"date": event_date, "original": net, "remaining": net})
            elif net < 0:
                paid_date = _settlement_date(group, data)
                if paid_date is not None:
                    payments.append({"date": paid_date, "remaining": -net})

        # 項目内でのみ、発生日以前の債務へ古い順に充当する。
        for payment in payments:
            for occurrence in occurrences:
                if occurrence["date"] > payment["date"] or occurrence["remaining"] <= 0:
                    continue
                applied = min(occurrence["remaining"], payment["remaining"])
                occurrence["remaining"] -= applied
                payment["remaining"] -= applied
                if payment["remaining"] <= 0:
                    break

        for occurrence in occurrences:
            due = _due_date(category, occurrence["date"], withholding_special)
            ledger_rows.append({
                "category": category, "occurrence_date": occurrence["date"], "due_date": due,
                "original_amount": occurrence["original"], "remaining_amount": occurrence["remaining"],
                "is_overdue": occurrence["remaining"] > 0 and data["date"].max() >= due,
            })

    return pd.DataFrame(ledger_rows), withholding_special


def evaluate_statutory_payments(journal: pd.DataFrame):
    ledger, withholding_special = build_statutory_payment_ledger(journal)
    if ledger.empty:
        return None

    overdue = ledger[ledger["is_overdue"]]
    overdue_months = overdue["due_date"].dt.to_period("M").nunique() if not overdue.empty else 0
    category_amounts = {
        category: float(group["remaining_amount"].sum())
        for category, group in overdue.groupby("category")
    }
    detail = "、".join(f"{category} {amount:,.0f}円" for category, amount in category_amounts.items())
    special_note = "（源泉所得税は納期特例と推定）" if withholding_special else ""

    if overdue_months >= 5:
        return "規律再設計期", "red", f"項目別・金額別に納付額を照合した結果、期限を超過した未消込残高として、{detail}が残っています。{special_note}", ledger
    if overdue_months >= 1:
        return "規律一部変動", "yellow", f"項目別・金額別に納付額を照合した結果、期限を超過した未消込残高として、{detail}が残っています。{special_note}", ledger
    return "資金規律安定", "blue", f"税金および社会保険料は、項目別の発生額に対して期限内に納付されています。{special_note}", ledger
