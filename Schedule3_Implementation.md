# Dynamic Inventory & Stock Summary Measures

Use these measures to build a Stock Summary report that matches Tally.

---

## 1. Opening Measures (From Masters)
```dax
Op Stock Qty = SUM(StockItem[OpeningBalance])

Op Stock Value = SUM(StockItem[OpeningValue])
```

---

## 2. Transaction Measures (For selected Period)
```dax
Total Receipts Qty = CALCULATE(SUM(StockVoucher[BilledQty]), StockVoucher[BilledQty] > 0)

Total Issues Qty = ABS(CALCULATE(SUM(StockVoucher[BilledQty]), StockVoucher[BilledQty] < 0))

Total Receipts Value = CALCULATE(SUM(StockVoucher[Amount]), StockVoucher[Amount] > 0)

Total Issues Value = ABS(CALCULATE(SUM(StockVoucher[Amount]), StockVoucher[Amount] < 0))
```

---

## 3. Closing Measures (As on max Date)
```dax
Cl Stock Qty = 
VAR SelectedDate = MAX(DateTable[Date])
VAR NetTransQty = CALCULATE(SUM(StockVoucher[BilledQty]), ALL(DateTable), DateTable[Date] <= SelectedDate)
RETURN [Op Stock Qty] + NetTransQty

Cl Stock Value = 
VAR SelectedDate = MAX(DateTable[Date])
VAR NetTransValue = CALCULATE(SUM(StockVoucher[Amount]), ALL(DateTable), DateTable[Date] <= SelectedDate)
RETURN [Op Stock Value] + NetTransValue
```

---

## 4. Verification Table (Visual Suggestion)
Create a **Table Visual** with:
- `StockItem[Name]`
- `[Op Stock Qty]`
- `[Total Receipts Qty]`
- `[Total Issues Qty]`
- `[Cl Stock Qty]`
- `[Cl Stock Value]`

This should exactly match Tally's **Stock Summary** (F12: Show Opening, Inwards, Outward).
