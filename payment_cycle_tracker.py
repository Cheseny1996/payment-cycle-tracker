"""
Payment Cycle Tracker
----------------------
Tracks accounts-receivable health for a subscription/services business:
Days Sales Outstanding (DSO), invoice aging buckets, collection efficiency,
and monthly DSO trend.

Usage:
    python payment_cycle_tracker.py

Outputs:
    payment_cycle_report.png   -- aging + DSO trend charts
    invoices_sample.csv        -- underlying invoice-level data
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta

RNG_SEED = 42
N_INVOICES = 420
TODAY = datetime(2026, 7, 1)


def generate_invoices(n=N_INVOICES, seed=RNG_SEED) -> pd.DataFrame:
    """Simulate 12 months of B2B invoices with realistic payment behavior."""
    rng = np.random.default_rng(seed)
    customers = [f"Customer {i:03d}" for i in range(1, 61)]

    issue_dates = [
        TODAY - timedelta(days=int(rng.uniform(1, 365))) for _ in range(n)
    ]
    terms = rng.choice([30, 45, 60], size=n, p=[0.6, 0.3, 0.1])
    amounts = rng.gamma(shape=3.0, scale=1800, size=n).round(2)

    # Payment delay: most pay close to terms, a tail pays very late,
    # a small slice is still outstanding.
    delay = rng.normal(loc=0, scale=12, size=n)
    late_tail = rng.random(n) < 0.12
    delay[late_tail] += rng.uniform(20, 60, size=late_tail.sum())
    still_open = rng.random(n) < 0.08

    rows = []
    for i in range(n):
        issue = issue_dates[i]
        due = issue + timedelta(days=int(terms[i]))
        if still_open[i] and (TODAY - issue).days < 120:
            paid = None
        else:
            paid_offset = max(1, terms[i] + delay[i])
            paid = issue + timedelta(days=float(paid_offset))
            if paid > TODAY:
                paid = None
        rows.append({
            "invoice_id": f"INV-{1000 + i}",
            "customer": customers[rng.integers(0, len(customers))],
            "issue_date": issue,
            "due_date": due,
            "terms_days": int(terms[i]),
            "amount": float(amounts[i]),
            "paid_date": paid,
        })

    df = pd.DataFrame(rows)
    return df


def compute_metrics(df: pd.DataFrame, as_of=TODAY) -> pd.DataFrame:
    df = df.copy()
    df["status"] = np.where(df["paid_date"].isna(), "Open", "Paid")
    df["days_to_pay"] = (df["paid_date"] - df["issue_date"]).dt.days
    df["days_outstanding"] = np.where(
        df["status"] == "Open",
        (as_of - df["issue_date"]).dt.days,
        df["days_to_pay"],
    )
    df["days_past_due"] = np.where(
        df["status"] == "Open",
        (as_of - df["due_date"]).dt.days,
        (df["paid_date"] - df["due_date"]).dt.days,
    )

    def bucket(days):
        if days <= 0:
            return "Current"
        elif days <= 30:
            return "1-30 days"
        elif days <= 60:
            return "31-60 days"
        elif days <= 90:
            return "61-90 days"
        return "90+ days"

    df["aging_bucket"] = df["days_past_due"].apply(bucket)
    return df


def dso_trend(df: pd.DataFrame, as_of=TODAY) -> pd.DataFrame:
    """Rolling monthly DSO using (AR balance / credit sales) * days-in-period."""
    df = df.copy()
    df["issue_month"] = df["issue_date"].dt.to_period("M")
    months = sorted(df["issue_month"].unique())[-9:]

    records = []
    for m in months:
        month_end = m.end_time
        sales_in_month = df[df["issue_month"] == m]["amount"].sum()
        ar_balance = df[
            (df["issue_date"] <= month_end)
            & ((df["paid_date"].isna()) | (df["paid_date"] > month_end))
        ]["amount"].sum()
        dso = (ar_balance / sales_in_month * 30) if sales_in_month else np.nan
        records.append({"month": str(m), "dso": round(dso, 1) if pd.notna(dso) else None})
    return pd.DataFrame(records)


def summary_stats(df: pd.DataFrame) -> dict:
    paid = df[df["status"] == "Paid"]
    open_ = df[df["status"] == "Open"]
    return {
        "total_invoices": len(df),
        "total_billed": round(df["amount"].sum(), 2),
        "open_ar_balance": round(open_["amount"].sum(), 2),
        "avg_days_to_pay": round(paid["days_to_pay"].mean(), 1),
        "pct_paid_on_time": round((paid["days_past_due"] <= 0).mean() * 100, 1),
        "pct_ar_90_plus": round(
            open_[open_["aging_bucket"] == "90+ days"]["amount"].sum()
            / max(open_["amount"].sum(), 1)
            * 100,
            1,
        ),
    }


def make_charts(df: pd.DataFrame, trend: pd.DataFrame, out_path="payment_cycle_report.png"):
    fig, axes = plt.subplots(1, 2, figsize=(13, 5))

    # Aging bucket chart (open AR only)
    open_df = df[df["status"] == "Open"]
    order = ["Current", "1-30 days", "31-60 days", "61-90 days", "90+ days"]
    bucket_totals = open_df.groupby("aging_bucket")["amount"].sum().reindex(order).fillna(0)
    colors = ["#2E7D32", "#9E9D24", "#F9A825", "#EF6C00", "#C62828"]
    axes[0].bar(bucket_totals.index, bucket_totals.values, color=colors)
    axes[0].set_title("Open AR by Aging Bucket")
    axes[0].set_ylabel("Amount Outstanding ($)")
    axes[0].tick_params(axis="x", rotation=30)

    # DSO trend
    axes[1].plot(trend["month"], trend["dso"], marker="o", color="#1565C0")
    axes[1].axhline(y=45, color="gray", linestyle="--", linewidth=1, label="Target DSO (45d)")
    axes[1].set_title("DSO Trend (Last 9 Months)")
    axes[1].set_ylabel("Days Sales Outstanding")
    axes[1].tick_params(axis="x", rotation=45)
    axes[1].legend()

    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    print(f"Saved chart to {out_path}")


if __name__ == "__main__":
    invoices = generate_invoices()
    invoices_metrics = compute_metrics(invoices)
    trend = dso_trend(invoices)
    stats = summary_stats(invoices_metrics)

    invoices_metrics.to_csv("invoices_sample.csv", index=False)
    make_charts(invoices_metrics, trend)

    print("\n--- Payment Cycle Summary ---")
    for k, v in stats.items():
        print(f"{k}: {v}")
