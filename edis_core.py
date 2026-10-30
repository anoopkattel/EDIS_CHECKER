"""Core logic for EDIS checking: reads CM1 + Meroshare files and compares them."""
import io
import re
import pandas as pd


def _clean_code(s):
    """Client codes are mixed (numbers and text like KDL03002), so keep them as clean text."""
    s = s.astype(str).str.strip()
    return s.str.replace(r"\.0$", "", regex=True)


def _clean_id(s):
    """Contract numbers. Also repairs values Excel turned into scientific notation
    (e.g. 2.026102901014816E+15) when the full digits are still stored in the file."""
    from decimal import Decimal, InvalidOperation

    def fix(v):
        v = str(v).strip()
        if re.fullmatch(r"-?\d+(\.\d+)?[eE][+-]?\d+", v):
            try:
                return format(Decimal(v), "f").split(".")[0]
            except InvalidOperation:
                return v
        return re.sub(r"\.0+$", "", v)
    return s.map(fix)


def read_any(file, header_keyword=None):
    """Read xls / xlsx / csv. If header_keyword is given, find the header row automatically
    (CM1 has title lines above the real header)."""
    name = getattr(file, "name", str(file)).lower()
    if name.endswith(".csv"):
        raw = pd.read_csv(file, header=None, dtype=str, keep_default_na=False)
    else:
        engine = "xlrd" if name.endswith(".xls") else None
        raw = pd.read_excel(file, header=None, dtype=str, engine=engine)
    raw = raw.fillna("")
    hdr = 0
    if header_keyword:
        found = None
        for i in range(min(40, len(raw))):
            if any(str(c).strip().lower() == header_keyword.lower() for c in raw.iloc[i]):
                found = i
                break
        if found is None:
            raise ValueError(f"Could not find a column named '{header_keyword}' in the file.")
        hdr = found
    df = raw.iloc[hdr + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[hdr]]
    return df.reset_index(drop=True)


def _col(df, *options):
    low = {c.lower(): c for c in df.columns}
    for o in options:
        if o.lower() in low:
            return low[o.lower()]
    raise ValueError(f"Missing column. Expected one of: {', '.join(options)}. Found: {', '.join(df.columns)}")


def prepare_cm1(file):
    df = read_any(file, "Client Code")
    out = pd.DataFrame({
        "contract": _clean_id(df[_col(df, "Contract No", "Contract Number")]),
        "settlement": _clean_code(df[_col(df, "Settlement ID", "Settlement Id")]),
        "client": _clean_code(df[_col(df, "Client Code")]),
        "scrip": df[_col(df, "Scrip Short Name", "Scrip")].astype(str).str.strip(),
        "qty": pd.to_numeric(df[_col(df, "Quantity")], errors="coerce"),
    })
    return out.dropna(subset=["qty"]).query("client != ''").reset_index(drop=True)


def prepare_mero(file):
    df = read_any(file, "Client Code")
    out = pd.DataFrame({
        "contract": _clean_id(df[_col(df, "Contract Number", "Contract No")]),
        "settlement": _clean_code(df[_col(df, "Settlement Id", "Settlement ID")]),
        "client": _clean_code(df[_col(df, "Client Code")]),
        "scrip": df[_col(df, "Scrip", "Scrip Short Name")].astype(str).str.strip(),
        "qty": pd.to_numeric(df[_col(df, "Quantity")], errors="coerce"),
    })
    return out.dropna(subset=["qty"]).query("client != ''").reset_index(drop=True)


MATCH_KEY = ["contract", "client", "scrip", "qty"]


def match_trades(cm1, mero):
    """Trade-by-trade matching on Contract No + Client Code + Scrip + Quantity.
    Each CM1 trade can be matched by exactly ONE Meroshare transfer and vice versa.
    If two trades are completely identical, the n-th copy in CM1 is paired with the
    n-th copy in Meroshare, so one transfer can never cover two trades."""
    c = cm1.copy()
    m = mero.copy()
    c["_n"] = c.groupby(MATCH_KEY).cumcount()
    m["_n"] = m.groupby(MATCH_KEY).cumcount()
    keys = MATCH_KEY + ["_n"]
    j = c.merge(m[keys].assign(_hit=True), on=keys, how="left")
    c["matched"] = j["_hit"].fillna(False).astype(bool).values
    m = m.merge(c[keys].assign(_hit=True), on=keys, how="left")
    m["matched"] = m["_hit"].fillna(False).astype(bool)
    return c.drop(columns="_n"), m.drop(columns=["_n", "_hit"])


def build_summary(cm1, mero):
    c, m = match_trades(cm1, mero)
    total_req, total_rec = float(cm1.qty.sum()), float(mero.qty.sum())

    c["received"] = c.qty.where(c.matched, 0)
    c["pending"] = c.qty - c.received
    d = c.groupby(["client", "scrip"], as_index=False).agg(
        required=("qty", "sum"), received=("received", "sum"), pending=("pending", "sum"))
    pending = d[d.pending > 0].copy()

    # transfers that did not match any CM1 trade (wrong contract / quantity / client / scrip)
    extra = m[~m.matched][["contract", "client", "scrip", "qty"]].rename(columns={"qty": "extra"})

    per_client = d.groupby("client").agg(required=("required", "sum"), received=("received", "sum"))
    status = pd.Series("done", index=per_client.index)
    status[(per_client.required > per_client.received) & (per_client.received <= 0)] = "not_done"
    status[(per_client.required > per_client.received) & (per_client.received > 0)] = "partial"
    return dict(total_req=total_req, total_rec=total_rec, diff=total_req - total_rec,
                detail=pending, extra=extra, status=status, per_client=per_client)


def with_contacts(clients, contacts):
    """clients: DataFrame with column 'client'. contacts: DataFrame with client, name, phone."""
    if contacts is None:
        clients = clients.assign(name="", phone="")
        return clients
    c = contacts.drop_duplicates("client")
    out = clients.merge(c, on="client", how="left")
    out["name"] = out["name"].fillna("Not found in contact list")
    out["phone"] = out["phone"].fillna("")
    return out


def client_table(summary, which, contacts=None):
    """which = 'not_done' or 'partial'. One row per client with pending scrips listed."""
    ids = summary["status"][summary["status"] == which].index
    det = summary["detail"][summary["detail"].client.isin(ids)]
    rows = []
    for cid, g in det.groupby("client"):
        rows.append({
            "client": cid,
            "pending_scrips": ", ".join(f"{r.scrip}: {int(r.pending)}" for r in g.itertuples()),
            "total_pending_qty": int(g.pending.sum()),
        })
    t = pd.DataFrame(rows, columns=["client", "pending_scrips", "total_pending_qty"])
    t = with_contacts(t, contacts)
    t = t[["client", "name", "phone", "pending_scrips", "total_pending_qty"]]
    t.columns = ["Client Code", "Name", "Phone", "Pending Scrips (qty)", "Total Pending Qty"]
    return t.sort_values("Client Code").reset_index(drop=True)


def prepare_contacts(file):
    """Contact list where the client code sits inside the name, e.g. 'RAM SHRESTHA [RS1234]'.
    If the file has a separate code column, that is used instead."""
    df = read_any(file)
    low = {c.lower(): c for c in df.columns}
    name_col = next((low[k] for k in low if "name" in k), None)
    phone_col = next((low[k] for k in low if any(w in k for w in ("mobile", "phone", "contact", "tel"))), None)
    if name_col is None or phone_col is None:
        raise ValueError(f"Need a name column and a mobile/phone column. Found: {', '.join(df.columns)}")
    code_col = next((low[k] for k in low if k in ("client code", "client id", "code", "client_code", "clientid")), None)
    names = df[name_col].astype(str).str.strip()
    if code_col:
        codes, clean = _clean_code(df[code_col]), names
    else:
        parts = names.str.extract(r"^(.*?)\s*\[([^\]]*)\]\s*$")
        if parts[1].notna().sum() == 0:
            raise ValueError("Could not find client codes. Expected names like 'RAM SHRESTHA [RS1234]' or a 'Client Code' column.")
        codes, clean = parts[1].str.strip(), parts[0].fillna(names)
    out = pd.DataFrame({"client": codes, "name": clean, "phone": df[phone_col].astype(str).str.strip()})
    return out[out.client.notna() & (out.client != "") & (out.client != "nan")]
