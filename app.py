import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Company X | Slow Stock Review", page_icon="📦", layout="wide")
st.title("📦 Company X — Slow Stock Review")
st.caption("Upload the monthly workbook. The tool validates data, applies the case rules, and creates an Excel action list.")

REQUIRED = {
    "Item_Master": ["SKU","Item Name","Category","Master Pack Qty","Unit Cost (Rs)","Alternate Item (name)","Alternate Flag","Status"],
    "Stock_Aging": ["SKU","Qty 0-30 days","Qty 31-60 days","Qty 61-90 days","Qty 91-180 days","Qty 180+ days"],
    "Sales_Lines": ["Month","CustomerID","CustomerName","Industry","SKU","Qty","Value (Rs)"],
    "Customers": ["CustomerID","CustomerName","Industry","Area"],
}
STOCK_COLS = REQUIRED["Stock_Aging"][1:]

def clean_sku(s):
    return s.astype("string").str.strip().str.upper()

def run_review(upload):
    xls = pd.ExcelFile(upload)
    tables = {}
    problems = []
    for sheet, cols in REQUIRED.items():
        if sheet not in xls.sheet_names:
            raise ValueError(f"Missing required sheet: {sheet}")
        d = pd.read_excel(upload, sheet_name=sheet)
        missing = [c for c in cols if c not in d.columns]
        if missing:
            raise ValueError(f"{sheet} is missing columns: {', '.join(missing)}")
        d = d[cols].copy()
        tables[sheet] = d

    items, stock, sales, customers = (tables[k] for k in ["Item_Master","Stock_Aging","Sales_Lines","Customers"])
    for name, d in [("Item_Master",items),("Stock_Aging",stock),("Sales_Lines",sales)]:
        before = d["SKU"].astype("string")
        cleaned = clean_sku(d["SKU"])
        changed = int((before.fillna("") != cleaned.fillna("")).sum())
        if changed:
            problems.append(f"{name}: trimmed spaces and standardized case in {changed} SKU value(s).")
        d["SKU"] = cleaned
    items["Alternate Item (name)"] = items["Alternate Item (name)"].astype("string").str.strip()
    items["Status"] = items["Status"].astype("string").str.strip().str.title()
    items["Alternate Flag"] = items["Alternate Flag"].astype("string").str.strip().str.upper()
    customers["CustomerID"] = customers["CustomerID"].astype("string").str.strip()
    sales["CustomerID"] = sales["CustomerID"].astype("string").str.strip()

    if items["SKU"].duplicated().any():
        problems.append(f"Item_Master has {int(items['SKU'].duplicated(keep=False).sum())} rows with duplicate SKU values; aggregated/reporting may be ambiguous.")
    if stock["SKU"].duplicated().any():
        problems.append(f"Stock_Aging has {int(stock['SKU'].duplicated(keep=False).sum())} rows with duplicate SKUs; stock buckets are summed.")
    exact_dup = int(sales.duplicated().sum())
    if exact_dup:
        problems.append(f"Sales_Lines contains {exact_dup} exact duplicate row(s). Kept them because they may be legitimate repeated transaction lines; confirm against source transaction IDs.")
    if customers["Industry"].isna().any():
        problems.append(f"Customers has {int(customers['Industry'].isna().sum())} customer(s) with missing industry. These customers are excluded from the similar-business bucket unless another rule applies.")
    if (sales[["Qty","Value (Rs)"]].isna().any(axis=1)).any():
        problems.append("Sales_Lines has missing Qty or Value; these rows should be corrected. Missing numeric values are treated as zero for calculations.")
        sales[["Qty","Value (Rs)"]] = sales[["Qty","Value (Rs)"]].fillna(0)
    for col in STOCK_COLS:
        stock[col] = pd.to_numeric(stock[col], errors="coerce").fillna(0)
    for col in ["Master Pack Qty","Unit Cost (Rs)"]:
        items[col] = pd.to_numeric(items[col], errors="coerce")
    for col in ["Qty","Value (Rs)"]:
        sales[col] = pd.to_numeric(sales[col], errors="coerce").fillna(0)

    name_to_sku = dict(zip(items["Item Name"].astype(str).str.strip(), items["SKU"]))
    switches = {}
    for _, r in items[items["Alternate Flag"].eq("SWITCH")].iterrows():
        alt_name = str(r["Alternate Item (name)"]).strip()
        alt_sku = name_to_sku.get(alt_name)
        if alt_sku:
            switches[r["SKU"]] = alt_sku
        else:
            problems.append(f"{r['SKU']}: SWITCH alternate item name '{alt_name}' did not match an Item_Master item name; no switch merge was made.")
    for old, new in switches.items():
        if old == new:
            problems.append(f"{old}: SWITCH points to itself; no merge was made.")
            switches.pop(old, None)
            break

    known = set(items["SKU"])
    unknown_sales = sorted(set(sales["SKU"]) - known)
    unknown_stock = sorted(set(stock["SKU"]) - known)
    if unknown_sales:
        problems.append(f"Sales_Lines contains {len(unknown_sales)} SKU value(s) absent from Item_Master after normalization: {', '.join(unknown_sales)}.")
    if unknown_stock:
        problems.append(f"Stock_Aging contains {len(unknown_stock)} SKU value(s) absent from Item_Master: {', '.join(unknown_stock)}.")
    inactive = items[items["Status"].str.lower().eq("inactive")]["SKU"].tolist()
    inactive_stock = stock[stock["SKU"].isin(inactive)].copy()
    if not inactive_stock.empty:
        inactive_stock["Stock Qty"] = inactive_stock[STOCK_COLS].sum(axis=1)
        inactive_stock = inactive_stock[inactive_stock["Stock Qty"] > 0]
        for _, r in inactive_stock.iterrows():
            problems.append(f"Inactive item {r['SKU']} still holds {int(r['Stock Qty'])} unit(s); excluded from classification and aged-value summary.")
    if not problems:
        problems.append("No data-quality issues detected by the built-in checks.")

    # Aggregate stock and sales by canonical (new) SKU for SWITCH pairs.
    stockagg = stock.groupby("SKU", dropna=False)[STOCK_COLS].sum()
    for old, new in switches.items():
        if old in stockagg.index and new in stockagg.index:
            stockagg.loc[new, STOCK_COLS] = stockagg.loc[new, STOCK_COLS].astype(float) + stockagg.loc[old, STOCK_COLS].astype(float)
            stockagg.loc[old, STOCK_COLS] = 0
    sales["CanonicalSKU"] = sales["SKU"].map(lambda x: switches.get(x, x))
    salesqty = sales.groupby("CanonicalSKU")["Qty"].sum()
    customers["CustomerID"] = customers["CustomerID"].astype(str)
    customer_values = sales.groupby("CustomerID")["Value (Rs)"].sum()
    cust_index = customers.drop_duplicates("CustomerID").set_index("CustomerID")
    results = []
    active = items[items["Status"].str.lower().eq("active")].copy()
    active = active[~active["SKU"].isin(switches.keys())]
    for _, r in active.iterrows():
        sku = r["SKU"]
        if sku in stockagg.index:
            sr = stockagg.loc[sku]
            stock_qty = float(sr[STOCK_COLS].sum())
            aged_qty = float(sr["Qty 91-180 days"] + sr["Qty 180+ days"])
        else:
            stock_qty = aged_qty = 0.0
        sold = float(salesqty.get(sku, 0))
        pack = float(r["Master Pack Qty"]) if pd.notna(r["Master Pack Qty"]) and r["Master Pack Qty"] > 0 else float("inf")
        cost = float(r["Unit Cost (Rs)"]) if pd.notna(r["Unit Cost (Rs)"]) else 0
        avg = sold / 6
        months = stock_qty / avg if avg > 0 else float("inf")
        if sold < pack:
            classification = "Red List"
        elif aged_qty == 0 or months < 3:
            classification = "No Action"
        else:
            classification = "Push"
        results.append({
            "SKU":sku, "Item Name":r["Item Name"], "Category":r["Category"],
            "Master Pack Qty":pack, "Unit Cost (Rs)":cost, "Total Stock Qty":stock_qty,
            "Aged Qty":aged_qty, "Aged Value (Rs)":aged_qty*cost,
            "6-Month Sales Qty":sold, "Avg Monthly Sales Qty":avg,
            "Months to Clear":months, "Class":classification
        })
    result = pd.DataFrame(results)
    # Customer targets
    eligible = set(customer_values[customer_values > 100000].index)
    targets = []
    item_categories = items.set_index("SKU")["Category"].to_dict()
    for _, ir in result[result["Class"].eq("Push")].iterrows():
        sku = ir["SKU"]
        item_sale_rows = sales[sales["CanonicalSKU"].eq(sku)]
        buyers = set(item_sale_rows["CustomerID"])
        buyer_industries = set(cust_index.reindex(list(buyers))["Industry"].dropna().astype(str))
        category_skus = set(items.loc[items["Category"].eq(ir["Category"]), "SKU"])
        category_skus = {switches.get(x,x) for x in category_skus}
        same_category_buyers = set(sales[sales["CanonicalSKU"].isin(category_skus) & ~sales["CanonicalSKU"].eq(sku)]["CustomerID"])
        for cid in eligible:
            if cid in buyers:
                bucket = 1
            elif cid in cust_index.index and pd.notna(cust_index.loc[cid, "Industry"]) and cust_index.loc[cid, "Industry"] in buyer_industries:
                bucket = 2
            elif cid in same_category_buyers:
                bucket = 3
            else:
                continue
            info = cust_index.loc[cid] if cid in cust_index.index else pd.Series(dtype=object)
            targets.append({
                "SKU":sku, "Item Name":ir["Item Name"], "Bucket":bucket,
                "CustomerID":cid, "Customer Name":info.get("CustomerName",""),
                "Industry":info.get("Industry",""), "Area":info.get("Area",""),
                "Customer 6-Month Purchase Value (Rs)":float(customer_values.get(cid,0))
            })
    target_df = pd.DataFrame(targets)
    if not target_df.empty:
        target_df = target_df.sort_values(["SKU","Bucket","Customer 6-Month Purchase Value (Rs)"], ascending=[True,True,False])
    return result, target_df, problems, inactive_stock if 'inactive_stock' in locals() else pd.DataFrame(), tables

def excel_export(result, targets, problems):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="Item Review")
        targets.to_excel(writer, index=False, sheet_name="Customer Targets")
        pd.DataFrame({"Data Checks":problems}).to_excel(writer, index=False, sheet_name="Data Quality")
    bio.seek(0)
    return bio

upload = st.file_uploader("Upload the monthly Excel workbook (.xlsx)", type=["xlsx"])
if upload:
    try:
        result, targets, problems, inactive, tables = run_review(upload)
        st.success("Review completed.")
        c1,c2,c3 = st.columns(3)
        counts = result["Class"].value_counts()
        c1.metric("Push items", int(counts.get("Push",0)))
        c2.metric("No Action items", int(counts.get("No Action",0)))
        c3.metric("Red List items", int(counts.get("Red List",0)))
        push_value = result.loc[result["Class"].eq("Push"),"Aged Value (Rs)"].sum()
        st.metric("Total aged value in Push items", f"₹{push_value:,.0f}")
        st.subheader("Top 10 Push items by aged value")
        st.dataframe(result[result["Class"].eq("Push")].sort_values("Aged Value (Rs)",ascending=False).head(10), use_container_width=True, hide_index=True)
        st.subheader("Item list")
        a,b = st.columns(2)
        classes = a.multiselect("Filter by class", sorted(result["Class"].dropna().unique()), default=sorted(result["Class"].dropna().unique()))
        categories = b.multiselect("Filter by category", sorted(result["Category"].dropna().unique()), default=sorted(result["Category"].dropna().unique()))
        filtered = result[result["Class"].isin(classes) & result["Category"].isin(categories)]
        st.dataframe(filtered, use_container_width=True, hide_index=True)
        st.subheader("Customer targets (mobile-friendly)")
        push_skus = result.loc[result["Class"].eq("Push"),"SKU"].tolist()
        if push_skus:
            selected = st.selectbox("Choose a Push item", push_skus, format_func=lambda x: f"{x} — {result.loc[result.SKU.eq(x),'Item Name'].iloc[0]}")
            view = targets[targets["SKU"].eq(selected)].copy()
            st.dataframe(view[["Bucket","Customer Name","Industry","Area","Customer 6-Month Purchase Value (Rs)"]], use_container_width=True, hide_index=True)
            st.caption("Bucket 1 = bought this item/switch partner; Bucket 2 = same industry; Bucket 3 = bought another item in the same category.")
        st.subheader("Data-quality checks")
        for p in problems:
            st.write("• " + p)
        st.download_button("Download results as Excel", data=excel_export(result,targets,problems), file_name="companyx_slow_stock_review.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Could not process this workbook: {e}")
else:
    st.info("Upload the provided workbook to begin. Required sheets: Item_Master, Stock_Aging, Sales_Lines, Customers.")
