# JKCL AP — Allocation Tool (new logic)

Production, user-independent tool that replaces the manual DTP (Steps 1–14). Any user
uploads the Exela DCN-level detail dump and gets a publish-ready, allocated workbook —
no hidden "JKCL Dashboard" file, no per-user formula dragging, no supervisor-mapping file.

## Run it

**Web app (recommended — for anyone on the team):**
```
pip install -r requirements.txt
streamlit run app.py
```
Upload the Exela detailed report → download the workbook.

**Command line (for schedulers / power users):**
```
python run_allocation.py "XBP_P2P_JKCL_Invoices_Status_Detailed_Report.xlsx"
```

## What it does (the new logic, encoded)

1. **Reads** the Exela DCN-level detail (header auto-detected; data from the real 82-col schema).
2. **Classifies** each invoice:
   - Utility / Non-Utility — from Vendor Group (51 mapped) with Account-Group fallback.
   - MSME flag — from MSME Number/Category (3-logic vendor-master lookup is a future hook).
   - Auto vs Manual — from Processing Type / Current Queue.
   - Fly Ash / Gypsum → flagged *auto-eligible, verify plant* (mirrors your master note).
3. **Routes to a team** by precedence: **Freight** (freight vendor groups) → **HO-White**
   (White division) → **HO-Grey** (owned vendor categories) → **Plant** (remainder).
   Auto-processing items → **API** users.
4. **Assigns a person**:
   - HO-Grey → by category ownership (LA/BO, CFA, IT, Contractor, CTS, Rent, Branding…).
   - Freight → by Grey/White scope, load-balanced.
   - Plant / HO-White → Parked→Parker, to-post→Poster, then even load-balance within role.
5. **Builds the publish workbook**:
   - `Allocation` — every invoice with Team / Role / Assigned To / Assigned Email + flags.
   - `Overall` — supervisor-wise pivot (dynamic COUNTIFS, reconciles to total).
   - `AP` — team & PO/Non-PO & utility summary for the email body.
   - `Review` — only the rows a human should check before publishing.

## Editing the logic (no code)

All rules live in **`allocation_config.json`**:
- `vendor_group_utility`, `account_group_utility` — utility classification.
- `hogrey_vendorgroup_to_category` — which vendor groups belong to which HO-Grey owner.
- `users` — roster, roles, HO-Grey category ownership, Freight scope.
- `poster_if_status_in` / `poster_if_queue_in` — the Parker/Poster split rule.
- `paints_route_to` — where the small "Paints" division folds (default HO-White).

Change the JSON, re-run — the engine and UI pick it up. No formulas to drag, no file chain.

## Open items flagged in the Review tab (by design, not bugs)

- **Blank vendor group / division** — Indexing-queue items not yet coded in Exela.
- **Verify auto-process plant** — Fly Ash/Gypsum: the dump carries delivery *city*, not the
  legal-entity plant in your master's auto list, so these are flagged not silently auto-routed.
- **Unmapped vendor group** — genuinely new groups your master doesn't list yet
  (e.g. *Advertisement & Brand*, *Trading Goods*).

## Next iteration (optional)
- Merge the SAP **FBV0 (Non-PO)** + **MIR6 (PO Held/Parked)** export to set Parker/Poster
  precisely (a DCN in the parked list → Poster), replacing the status heuristic.
- Wire the vendor-master file for true MSME flagging (logic-1/2/3).
