"""
CLI runner — no UI needed.
Usage:  python run_allocation.py <exela_detail_dump.xlsx> [output.xlsx]
"""
import sys, os, datetime as dt
import allocation_engine as E
import build_output as B

def main():
    if len(sys.argv) < 2:
        print("Usage: python run_allocation.py <dump.xlsx> [output.xlsx]")
        sys.exit(1)
    dump = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else \
        f"Pendency allocation new logic - {dt.date.today():%d_%b'%y}.xlsx"
    res = E.run(dump)
    B.build(res, out)
    flagged = (res["Needs Review"] != "").sum()
    print(f"OK  {len(res):,} invoices allocated -> {out}")
    print("    Teams:", dict(res["Team"].value_counts()))
    print(f"    Parker {(res['Role']=='Parker').sum():,} | Poster {(res['Role']=='Poster').sum():,}")
    if flagged:
        print(f"    {flagged} rows flagged in the Review tab.")

if __name__ == "__main__":
    main()
