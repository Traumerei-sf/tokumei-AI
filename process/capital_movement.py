"""口座間の資金移動を起点に、期限付きの時系列充当で後続用途を推定する。"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable

import pandas as pd


BANK_PATTERN = re.compile(r"預金|当座|普通|定期|別段")
SUSPENSE_PATTERN = re.compile(r"^\s*(諸口|複合|振替)\s*$")
NULL_TRANSACTION_VALUES = {"", "nan", "none", "null", "<na>"}
EPSILON = 0.5
OUTPUT_COLUMNS = [
    "資金移動日",
    "移動金額",
    "用途充当額",
    "金額差（未充当額）",
    "金額差率（未充当率）",
    "想定用途",
    "資金移動仕訳",
    "用途推定仕訳",
    "一致区分",
    "信頼度",
]


@dataclass(frozen=True)
class PaymentCandidate:
    candidate_id: str
    group_key: str
    transaction_no: str
    order: int
    date: pd.Timestamp
    account: str
    sub_account: str
    amount: float
    purpose: str
    priority: int
    related_debit: str
    related_credit: str
    is_direct_pair: bool


@dataclass(frozen=True)
class TransferOrigin:
    origin_id: str
    group_key: str
    transaction_no: str
    order: int
    date: pd.Timestamp
    source_account: str
    source_sub_account: str
    destination_account: str
    destination_sub_account: str
    amount: float
    transfer_entry: str


@dataclass(frozen=True)
class Allocation:
    origin_id: str
    candidate_id: str
    amount: float
    days: int
    phase: str


def _text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _amount(value) -> float:
    try:
        if pd.isna(value):
            return 0.0
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _is_bank(account) -> bool:
    return bool(BANK_PATTERN.search(_text(account)))


def _valid_transaction_no(value) -> bool:
    return _text(value).lower() not in NULL_TRANSACTION_VALUES


def _group_key(index, transaction_no) -> str:
    if _valid_transaction_no(transaction_no):
        return f"tx:{_text(transaction_no)}"
    return f"row:{index}"


def _display_transaction_no(value) -> str:
    value = _text(value)
    return re.sub(r"^\d+_", "", value) if value else "番号なし"


def _account_label(account, sub_account) -> str:
    account = _text(account)
    sub_account = _text(sub_account)
    return f"{account} {sub_account}".strip()


def _format_entry(account, sub_account, description, amount) -> str:
    middle = _text(sub_account) or _text(description)
    return f"【{_text(account)}】【{middle}】（{_amount(amount):,.0f}円）"


def _purpose_info(rows: pd.DataFrame) -> tuple[str | None, int]:
    """諸口を用途にせず、対象となる借方明細と摘要から用途を決める。"""
    accounts = [
        _text(value)
        for value in rows.get("debit_account", pd.Series(dtype=object))
        if _text(value) and not _is_bank(value) and not SUSPENSE_PATTERN.match(_text(value))
    ]
    descriptions = " ".join(
        _text(value)
        for column in ("description", "payment_partner", "partner")
        if column in rows.columns
        for value in rows[column]
    )
    account_text = " ".join(accounts)
    # 一行で借方・銀行貸方が対応する場合は、補助科目もその支払固有の用途根拠にできる。
    if len(rows) == 1 and "debit_partner" in rows.columns:
        account_text = f"{account_text} {_text(rows.iloc[0].get('debit_partner'))}".strip()
    all_text = f"{account_text} {descriptions}"
    # 税・社会保険の実科目が借方に立つ支払は、摘要中の「給与」等より優先する。
    statutory_accounts = [
        "租税公課", "預り金", "法定福利費", "法人税", "消費税", "事業税", "住民税"
    ]
    payroll_accounts = ["給料", "給与", "役員報酬", "賞与", "退職金"]
    if (
        any(keyword in account_text for keyword in statutory_accounts)
        and not any(keyword in account_text for keyword in payroll_accounts)
    ):
        return "税金・社会保険料", 8
    rules = [
        ("給与支払い", 1, ["給料", "給与", "役員報酬", "賞与", "退職金"]),
        ("銀行返済", 2, ["借入", "返済", "元金", "融資", "公庫", "保証協会"]),
        ("カード決済", 3, ["カード", "VISA", "AMEX", "JCB", "ニコス", "NICOS", "セゾン", "SAISON"]),
        ("月末支払い", 4, ["買掛", "未払", "外注", "月末EB", "月末振込"]),
        ("大口支払い", 5, ["仕入", "設備", "土地", "建物", "車両", "機械", "構築物", "ソフトウェア", "商標", "特許", "のれん", "支払手形"]),
        ("社長関連資金", 6, ["社長", "役員貸付", "役員借入"]),
        ("現金引出", 7, ["現金"]),
        ("税金・社会保険料", 8, ["税務署", "住民税", "所得税", "法人税", "消費税", "事業税", "社会保険", "年金"]),
    ]
    for purpose, priority, keywords in rules:
        if any(keyword in all_text for keyword in keywords):
            return purpose, priority
    if accounts:
        return "その他支払", 9
    return None, 99


def _related_lines(rows: pd.DataFrame, side: str, meaningful_debits_only: bool = False) -> str:
    account_col = f"{side}_account"
    partner_col = f"{side}_partner"
    amount_col = f"{side}_amount"
    lines = []
    for _, row in rows.iterrows():
        amount = _amount(row.get(amount_col))
        account = _text(row.get(account_col))
        if amount <= 0 or not account:
            continue
        if meaningful_debits_only and side == "debit":
            if _is_bank(account) or SUSPENSE_PATTERN.match(account):
                continue
        lines.append(_format_entry(account, row.get(partner_col), row.get("description"), amount))
    return "\n".join(lines) if lines else "（なし）"


def _prepare_groups(journal: pd.DataFrame) -> list[tuple[str, pd.DataFrame]]:
    work = journal.copy().reset_index(drop=False).rename(columns={"index": "_source_index"})
    work["_source_order"] = range(len(work))
    transaction_numbers = work.get("transaction_no", pd.Series(pd.NA, index=work.index))
    work["_movement_group"] = [
        _group_key(index, value)
        for index, value in zip(work["_source_order"], transaction_numbers)
    ]
    return [(key, rows.copy()) for key, rows in work.groupby("_movement_group", sort=False)]


def _is_transfer_row(row: pd.Series) -> bool:
    return (
        _is_bank(row.get("debit_account"))
        and _is_bank(row.get("credit_account"))
        and _amount(row.get("debit_amount")) > 0
        and _amount(row.get("credit_amount")) > 0
    )


def _is_direct_payment_pair(row: pd.Series) -> bool:
    debit_account = _text(row.get("debit_account"))
    debit_amount = _amount(row.get("debit_amount"))
    credit_amount = _amount(row.get("credit_amount"))
    return (
        _is_bank(row.get("credit_account"))
        and not _is_bank(debit_account)
        and not SUSPENSE_PATTERN.match(debit_account)
        and debit_amount > 0
        and credit_amount > 0
        and abs(debit_amount - credit_amount) <= EPSILON
    )


def _build_candidates(groups: Iterable[tuple[str, pd.DataFrame]]) -> list[PaymentCandidate]:
    """明細対応できる支払は行単位、それ以外の複合仕訳は取引No・口座単位にする。"""
    candidates: list[PaymentCandidate] = []
    for group_key, rows in groups:
        # 口座間移動の取引Noは起点専用とし、手数料を含めて用途候補にはしない。
        if rows.apply(_is_transfer_row, axis=1).any():
            continue
        dates = rows.get("date", pd.Series(dtype="datetime64[ns]")).dropna()
        if dates.empty:
            continue
        group_date = dates.min()
        transaction_no = _display_transaction_no(rows.get("transaction_no", pd.Series(pd.NA)).iloc[0])
        bank_rows = rows[
            rows["credit_account"].map(_is_bank)
            & (rows["credit_amount"].map(_amount) > 0)
        ].copy()
        if bank_rows.empty:
            continue

        direct_rows = bank_rows[bank_rows.apply(_is_direct_payment_pair, axis=1)].copy()
        direct_orders = set(direct_rows["_source_order"].tolist())
        for _, row in direct_rows.iterrows():
            one_row = row.to_frame().T
            purpose, priority = _purpose_info(one_row)
            if purpose is None:
                continue
            order = int(row["_source_order"])
            candidates.append(PaymentCandidate(
                candidate_id=f"{group_key}|line:{order}",
                group_key=group_key,
                transaction_no=transaction_no,
                order=order,
                date=group_date,
                account=_text(row.get("credit_account")),
                sub_account=_text(row.get("credit_partner")),
                amount=_amount(row.get("credit_amount")),
                purpose=purpose,
                priority=priority,
                related_debit=_related_lines(one_row, "debit", meaningful_debits_only=True),
                related_credit=_related_lines(one_row, "credit"),
                is_direct_pair=True,
            ))

        remaining_bank_rows = bank_rows[~bank_rows["_source_order"].isin(direct_orders)].copy()
        if remaining_bank_rows.empty:
            continue
        purpose_rows = rows[~rows["_source_order"].isin(direct_orders)].copy()
        purpose, priority = _purpose_info(purpose_rows)
        if purpose is None:
            continue
        for (account, sub_account), account_rows in remaining_bank_rows.groupby(
            ["credit_account", "credit_partner"], dropna=False, sort=False
        ):
            amount = account_rows["credit_amount"].map(_amount).sum()
            if amount <= 0:
                continue
            order = int(account_rows["_source_order"].min())
            candidates.append(PaymentCandidate(
                candidate_id=f"{group_key}|account:{_text(account)}|{_text(sub_account)}|{order}",
                group_key=group_key,
                transaction_no=transaction_no,
                order=order,
                date=group_date,
                account=_text(account),
                sub_account=_text(sub_account),
                amount=amount,
                purpose=purpose,
                priority=priority,
                related_debit=_related_lines(purpose_rows, "debit", meaningful_debits_only=True),
                related_credit=_related_lines(account_rows, "credit"),
                is_direct_pair=False,
            ))
    return sorted(candidates, key=lambda item: (item.date, item.order, item.candidate_id))


def _build_transfer_origins(
    groups: Iterable[tuple[str, pd.DataFrame]], minimum_amount: float
) -> list[TransferOrigin]:
    origins: list[TransferOrigin] = []
    for group_key, rows in groups:
        dates = rows.get("date", pd.Series(dtype="datetime64[ns]")).dropna()
        if dates.empty:
            continue
        transfer_rows = rows[rows.apply(_is_transfer_row, axis=1)].copy()
        if transfer_rows.empty:
            continue
        transaction_no = _display_transaction_no(rows.get("transaction_no", pd.Series(pd.NA)).iloc[0])
        group_columns = ["credit_account", "credit_partner", "debit_account", "debit_partner"]
        for keys, paired_rows in transfer_rows.groupby(group_columns, dropna=False, sort=False):
            source_account, source_sub, destination_account, destination_sub = map(_text, keys)
            amount = paired_rows["credit_amount"].map(_amount).sum()
            if amount < minimum_amount:
                continue
            order = int(paired_rows["_source_order"].min())
            entry_lines = []
            for _, row in paired_rows.iterrows():
                entry_lines.append(
                    f"{pd.Timestamp(row['date']):%Y-%m-%d} No.{transaction_no} "
                    f"{_account_label(row['credit_account'], row.get('credit_partner'))} → "
                    f"{_account_label(row['debit_account'], row.get('debit_partner'))}　"
                    f"{_amount(row.get('credit_amount')):,.0f}円"
                )
            origins.append(TransferOrigin(
                origin_id=(
                    f"{group_key}|{source_account}|{source_sub}|"
                    f"{destination_account}|{destination_sub}"
                ),
                group_key=group_key,
                transaction_no=transaction_no,
                order=order,
                date=dates.min(),
                source_account=source_account,
                source_sub_account=source_sub,
                destination_account=destination_account,
                destination_sub_account=destination_sub,
                amount=amount,
                transfer_entry="\n".join(entry_lines),
            ))
    return sorted(origins, key=lambda item: (item.date, item.order, item.origin_id))


def _same_account(candidate: PaymentCandidate, origin: TransferOrigin) -> bool:
    if candidate.account != origin.destination_account:
        return False
    if candidate.sub_account and origin.destination_sub_account:
        return candidate.sub_account == origin.destination_sub_account
    return True


def _allocate(
    origins: list[TransferOrigin], candidates: list[PaymentCandidate]
) -> tuple[list[Allocation], dict[str, float], dict[str, float]]:
    """当日～+7日をFIFOで処理し、残額だけを-2日～-1日へ充当する。"""
    origin_remaining = {origin.origin_id: origin.amount for origin in origins}
    candidate_remaining = {candidate.candidate_id: candidate.amount for candidate in candidates}
    allocations: list[Allocation] = []

    # 支払発生日ごとに、その時点で有効な最古の資金移動から充当する。
    for candidate in candidates:
        eligible = [
            origin
            for origin in origins
            if origin_remaining[origin.origin_id] > EPSILON
            and origin.group_key != candidate.group_key
            and _same_account(candidate, origin)
            and origin.date <= candidate.date <= origin.date + pd.Timedelta(days=7)
        ]
        for origin in eligible:
            if candidate_remaining[candidate.candidate_id] <= EPSILON:
                break
            allocated = min(
                origin_remaining[origin.origin_id],
                candidate_remaining[candidate.candidate_id],
            )
            if allocated <= EPSILON:
                continue
            days = int((candidate.date - origin.date).days)
            allocations.append(Allocation(
                origin.origin_id, candidate.candidate_id, allocated, days, "future"
            ))
            origin_remaining[origin.origin_id] -= allocated
            candidate_remaining[candidate.candidate_id] -= allocated

    # 将来側で余った移動資金のみ、直前2日間の未使用支払へ補充として充当する。
    for origin in origins:
        if origin_remaining[origin.origin_id] <= EPSILON:
            continue
        eligible = [
            candidate
            for candidate in candidates
            if candidate_remaining[candidate.candidate_id] > EPSILON
            and origin.group_key != candidate.group_key
            and _same_account(candidate, origin)
            and origin.date - pd.Timedelta(days=2) <= candidate.date < origin.date
        ]
        for candidate in eligible:
            if origin_remaining[origin.origin_id] <= EPSILON:
                break
            allocated = min(
                origin_remaining[origin.origin_id],
                candidate_remaining[candidate.candidate_id],
            )
            if allocated <= EPSILON:
                continue
            days = int((candidate.date - origin.date).days)
            allocations.append(Allocation(
                origin.origin_id, candidate.candidate_id, allocated, days, "lookback"
            ))
            origin_remaining[origin.origin_id] -= allocated
            candidate_remaining[candidate.candidate_id] -= allocated

    return allocations, origin_remaining, candidate_remaining


def _connected_components(
    origins: list[TransferOrigin], allocations: list[Allocation]
) -> list[tuple[set[str], set[str], list[Allocation]]]:
    origin_to_candidates: dict[str, set[str]] = {origin.origin_id: set() for origin in origins}
    candidate_to_origins: dict[str, set[str]] = {}
    for allocation in allocations:
        origin_to_candidates[allocation.origin_id].add(allocation.candidate_id)
        candidate_to_origins.setdefault(allocation.candidate_id, set()).add(allocation.origin_id)

    components = []
    seen_origins: set[str] = set()
    for origin in origins:
        if origin.origin_id in seen_origins:
            continue
        component_origins: set[str] = set()
        component_candidates: set[str] = set()
        pending_origins = [origin.origin_id]
        while pending_origins:
            origin_id = pending_origins.pop()
            if origin_id in component_origins:
                continue
            component_origins.add(origin_id)
            for candidate_id in origin_to_candidates.get(origin_id, set()):
                if candidate_id not in component_candidates:
                    component_candidates.add(candidate_id)
                    pending_origins.extend(candidate_to_origins.get(candidate_id, set()))
        seen_origins.update(component_origins)
        component_allocations = [
            allocation
            for allocation in allocations
            if allocation.origin_id in component_origins
            and allocation.candidate_id in component_candidates
        ]
        components.append((component_origins, component_candidates, component_allocations))
    return components


def _match_type(origin_count: int, candidate_count: int) -> str:
    if candidate_count == 0:
        return "なし"
    if origin_count == 1 and candidate_count == 1:
        return "1対1"
    if origin_count == 1:
        return "1対N"
    if candidate_count == 1:
        return "N対1"
    return "N対N"


def _purpose_label(selected: list[PaymentCandidate]) -> str:
    purposes = []
    for candidate in sorted(selected, key=lambda item: (item.priority, item.date, item.order)):
        if candidate.purpose not in purposes:
            purposes.append(candidate.purpose)
    return "＋".join(purposes) if purposes else "口座間移動（後続用途不明）"


def _payment_entry(candidate: PaymentCandidate, allocated_amount: float) -> str:
    if abs(candidate.amount - allocated_amount) <= EPSILON:
        allocation_text = f"{candidate.amount:,.0f}円（全額充当）"
    else:
        allocation_text = (
            f"支払総額{candidate.amount:,.0f}円のうち"
            f"{allocated_amount:,.0f}円を充当"
        )
    return (
        f"{candidate.date:%Y-%m-%d} No.{candidate.transaction_no} "
        f"{candidate.purpose}　{allocation_text}\n{candidate.related_debit}"
    )


def _confidence(unallocated_rate: float) -> str:
    """対応関係や用途にかかわらず、未充当率だけで信頼度を判定する。"""
    if unallocated_rate <= 0.10:
        return "高"
    if unallocated_rate <= 0.20:
        return "中"
    return "低"


def build_capital_movement_list(
    journal: pd.DataFrame,
    minimum_amount: float = 500_000,
) -> pd.DataFrame:
    """口座間移動と支払を-2日～+7日の範囲で時系列充当する。"""
    if journal.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    groups = _prepare_groups(journal)
    origins = _build_transfer_origins(groups, minimum_amount)
    candidates = _build_candidates(groups)
    if not origins:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    allocations, _, _ = _allocate(origins, candidates)
    origin_by_id = {origin.origin_id: origin for origin in origins}
    candidate_by_id = {candidate.candidate_id: candidate for candidate in candidates}
    output = []

    for origin_ids, candidate_ids, component_allocations in _connected_components(origins, allocations):
        selected_origins = sorted(
            (origin_by_id[item] for item in origin_ids),
            key=lambda item: (item.date, item.order, item.origin_id),
        )
        selected_candidates = sorted(
            (candidate_by_id[item] for item in candidate_ids),
            key=lambda item: (item.date, item.order, item.candidate_id),
        )
        origin_amount = sum(origin.amount for origin in selected_origins)
        allocated_amount = sum(item.amount for item in component_allocations)
        unallocated_amount = max(0.0, origin_amount - allocated_amount)
        unallocated_rate = unallocated_amount / origin_amount if origin_amount else 0.0
        match_type = _match_type(len(selected_origins), len(selected_candidates))
        allocated_by_candidate = {
            candidate.candidate_id: sum(
                item.amount
                for item in component_allocations
                if item.candidate_id == candidate.candidate_id
            )
            for candidate in selected_candidates
        }
        if len(selected_origins) == 1:
            movement_date: object = selected_origins[0].date
        else:
            movement_date = "\n".join(f"{origin.date:%Y-%m-%d}" for origin in selected_origins)
        transfer_entry = "\n".join(origin.transfer_entry for origin in selected_origins)
        if selected_candidates:
            payment_entry = "\n\n".join(
                _payment_entry(candidate, allocated_by_candidate[candidate.candidate_id])
                for candidate in selected_candidates
            )
        else:
            payment_entry = "（-2日～+7日の範囲に充当可能な支払なし）"

        output.append({
            "_sort_date": min(origin.date for origin in selected_origins),
            "資金移動日": movement_date,
            "移動金額": origin_amount,
            "用途充当額": allocated_amount,
            "金額差（未充当額）": unallocated_amount,
            "金額差率（未充当率）": unallocated_rate,
            "想定用途": _purpose_label(selected_candidates),
            "資金移動仕訳": transfer_entry,
            "用途推定仕訳": payment_entry,
            "一致区分": match_type,
            "信頼度": _confidence(unallocated_rate),
        })

    return (
        pd.DataFrame(output)
        .sort_values("_sort_date")
        .drop(columns="_sort_date")
        .reindex(columns=OUTPUT_COLUMNS)
        .reset_index(drop=True)
    )
