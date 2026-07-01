# Stripe Revenue Report

A payout-based revenue accounting tool for Stripe. Given a month, it walks every bank payout, attributes each balance transaction to the product that earned it, distributes Stripe fees proportionally, and produces a reconciled revenue table.

The core question it answers:

> **How much did each product actually pay out to my bank account this month, after all fees, refunds, and disputes?**

---

## Why payout-based?

Invoice and charge objects record intent, not settlement. A charge can be refunded; an invoice can be voided; a dispute can claw money back weeks later. Only a **payout** represents money that actually landed in your bank account.

This script uses Stripe's balance transactions — grouped by payout — as the single source of truth, so every cent is accounted for.

---

## How it works

1. **Fetch payouts** in the selected month's date range.
2. **Fetch balance transactions** for each payout and immediately assert the sum equals `payout.amount`.
3. **Attribute every transaction** to a product bucket:
   | Transaction type | Resolution path |
   |---|---|
   | `charge` | Charge → Invoice *or* Checkout Session → Product |
   | `refund` | Refund → PaymentIntent → Invoice *or* Checkout Session → Product |
   | `dispute` | Dispute → Charge → Invoice *or* Checkout Session → Product |
   | `dispute_reversal` | Same as `dispute` (money returned when dispute is won) |
   | `fee` | Recorded as a separate `fee` bucket |
   | `payout_minimum_balance_*` | Tracked separately; asserted to net zero |
4. **Assert balance at every level**: per-transaction, per-payout, and globally.
5. **Distribute Stripe fees** proportionally across products by revenue share.
6. **Output** a terminal table, a CSV, and per-product Excel files.

---

## Accuracy guarantees

The script hard-asserts correctness at four levels. Any discrepancy raises an error with a precise gap value before writing any output.

| Level | What is checked |
|---|---|
| Per transaction | `net == amount − fee` (Stripe's own invariant) |
| Per transaction | Every non-zero-net type has explicit handling; unknown types raise |
| Per payout | Attributed transaction nets sum exactly to `payout.amount` |
| Global | `acc.total_cents() == sum(payout.amount for all payouts)` |
| Table | Revenue($) column sums to total payout; fee is fully distributed |

---

## Usage

```bash
pip install -r requirements.txt
export STRIPE_SECRET_KEY=sk_live_...

# Report for a specific month (defaults to current year/month)
python report.py --month 6
python report.py --year 2025 --month 12

# Show cache hits during execution
python report.py --month 6 --debug

# Invalidate a cached Stripe object by ID
python report.py --delete-cache ch_abc123
```

---

## Output

For a run against month `YYYY-M`, the script writes:

| File | Contents |
|---|---|
| Terminal | Reconciliation log + revenue table |
| `reports/YYYY-MM/{Product}.xlsx` | One workbook per product with transaction-level detail |

Each per-product sheet contains one row per charge, refund, and dispute/dispute_reversal attributed to
that product: transaction type, customer email, customer name, amount, Stripe fee, net, payment intent ID,
receipt URL, and balance transaction ID. Refunds and disputes are recorded with negative amounts, so summing
a sheet's `net` column reproduces that product's `Revenue ($)` in the terminal table exactly.

---

## Revenue table columns

| Column | Description |
|---|---|
| Product | Stripe product name |
| Revenue ($) | Raw net (after Stripe fee) from balance transactions |
| Adjusted Fee ($) | This product's share of the total Stripe fee, allocated by revenue |
| Adjusted Revenue ($) | Revenue + Adjusted Fee (what the product actually contributed) |
| Email | Partner email from product metadata (`metadata.email`) |
| Rate (%) | Revenue-share rate from product metadata (`metadata.rate`) |
| Net Payout ($) | Adjusted Revenue × (1 − Rate) — amount owed to the partner |

The **Total row** Revenue and Adjusted Revenue both equal the month's total payout amount.

---

## Caching

Stripe objects (Product, Invoice, PaymentIntent, Charge, Refund, Dispute) are cached in a local SQLite file (`stripe_cache.sqlite`) after the first fetch. Re-runs are fast and consume no Stripe API quota for already-seen objects.

The cache is transparent to financial logic — results are identical with or without it.

To invalidate a specific entry:

```bash
python report.py --delete-cache <stripe-object-id>
```

---

## Assumptions

- Each invoice has exactly one line item.
- Each checkout session has exactly one line item.
- `payout_minimum_balance_*` transactions net to zero for the month.
- Partner email and revenue-share rate are stored in Stripe product metadata as `email` and `rate` (e.g. `rate=0.15` for 15%).

---

## License

MIT
