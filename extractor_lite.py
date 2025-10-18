import io
from datetime import datetime
from typing import List, Dict, Any
import pandas as pd
import pdfplumber
import re
import yaml

COLUMNS = ["Date","Description","Debit","Credit","Balance","CheckNo","Page","SourceLine"]

UNIVERSAL_TEMPLATE = {
  "line_regex": r"^(?P<date>(?:\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{2}\s+[A-Za-z]{3}\s+\d{4}|[A-Za-z]{3}\s+\d{1,2},\s*\d{4}))\s+(?P<desc>.+?)\s+(?P<amount>[\-\+\(\)]?\d{1,3}(?:,\d{3})*(?:\.\d{2})?)(?:\s+(?P<balance>\(?\-?\d{1,3}(?:,\d{3})*(?:\.\d{2})?\)?))?$",
  "ignore_patterns": [
    r"^Total\b", r"^Subtotal\b", r"^Account\b", r"^Important\b",
    r"^Service fee(s)?\b", r"^Opening balance\b", r"^Closing balance\b",
    r"^Balance brought forward\b", r"^Page\s+\d+\s+of\s+\d+\b", r"^\*+\s*Continued\s*\*+$"
  ]
}

class GenericParser:
    def __init__(self, template: Dict[str, Any] | None = None):
        import re
        self.template = template or {}
        self.line_regex = self.template.get("line_regex") or UNIVERSAL_TEMPLATE["line_regex"]
        self.ignore_pats = [re.compile(p) for p in self.template.get("ignore_patterns", [])]

    def _norm_amt(self, s: str | None):
        if not s: return None
        s = s.replace(",", "").strip()
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        try: return float(s)
        except: return None

    def _ignored(self, text: str) -> bool:
        for rx in self.ignore_pats:
            if rx.search(text):
                return True
        return False

    def parse(self, lines: List[Dict[str, Any]]) -> pd.DataFrame:
        out, unparsed, ignored = [], [], []
        import re
        rx = re.compile(self.line_regex)
        for ln in lines:
            text = " ".join((ln.get("text") or "").split())
            if not text: continue
            if self._ignored(text):
                ignored.append(ln); continue
            m = rx.search(text)
            if not m:
                unparsed.append(ln); continue
            gd = m.groupdict()
            amount = self._norm_amt(gd.get("amount"))
            balance = self._norm_amt(gd.get("balance"))
            debit = credit = None
            if amount is not None:
                if amount < 0: debit = abs(amount)
                else: credit = amount
            out.append({
                "Date": gd.get("date"),
                "Description": gd.get("desc"),
                "Debit": debit,
                "Credit": credit,
                "Balance": balance,
                "CheckNo": None,
                "Page": ln.get("page"),
                "SourceLine": text
            })
        df = pd.DataFrame(out, columns=COLUMNS)
        df.attrs["unparsed"] = unparsed
        df.attrs["ignored"] = ignored
        return df

def extract_lines_pdfplumber_bytes(pdf_bytes: bytes) -> List[Dict[str, Any]]:
    out = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(x_tolerance=2, y_tolerance=2) or []
            if not words:
                text = (page.extract_text() or "").splitlines()
                for t in text:
                    if t.strip():
                        out.append({"text": t.strip(), "x0": None, "x1": None, "top": None, "bottom": None, "page": i})
                continue
            words.sort(key=lambda w: (round(w["top"], 1), w["x0"]))
            current_top, buffer = None, []
            for w in words:
                t = round(w["top"], 1)
                if current_top is None: current_top = t
                if abs(t - current_top) <= 1.0: buffer.append(w)
                else:
                    if buffer:
                        text = " ".join(x["text"] for x in sorted(buffer, key=lambda z: z["x0"]))
                        x0 = min(x["x0"] for x in buffer); x1 = max(x["x1"] for x in buffer)
                        top = min(x["top"] for x in buffer); bottom = max(x["bottom"] for x in buffer)
                        out.append({"text": text, "x0": x0, "x1": x1, "top": top, "bottom": bottom, "page": i})
                    buffer = [w]; current_top = t
            if buffer:
                text = " ".join(x["text"] for x in sorted(buffer, key=lambda z: z["x0"]))
                x0 = min(x["x0"] for x in buffer); x1 = max(x["x1"] for x in buffer)
                top = min(x["top"] for x in buffer); bottom = max(x["bottom"] for x in buffer)
                out.append({"text": text, "x0": x0, "x1": x1, "top": top, "bottom": bottom, "page": i})
    return out

def process_pdf(file_bytes: bytes, template_yaml: str | None = None):
    lines = extract_lines_pdfplumber_bytes(file_bytes)
    if not lines:
        return None, {"error": "No extractable text detected. This looks like a scanned PDF. Cloud-lite supports only text PDFs."}
    template = None
    if template_yaml:
        try:
            template = yaml.safe_load(template_yaml) or {}
        except Exception:
            template = None
    parser = GenericParser(template=template or None)
    tx_df = parser.parse(lines)
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as xw:
        tx_df.to_excel(xw, index=False, sheet_name="Transactions")
        pd.DataFrame(lines, columns=["page","text","x0","x1","top","bottom"]).to_excel(xw, index=False, sheet_name="Raw_Extract")
        issues = []
        for up in tx_df.attrs.get("unparsed", []):
            issues.append({"type":"unparsed_line","page":up.get("page"),"text":up.get("text")})
        for ig in tx_df.attrs.get("ignored", []):
            issues.append({"type":"ignored_line","page":ig.get("page"),"text":ig.get("text")})
        (pd.DataFrame(issues) if issues else pd.DataFrame(columns=["type","page","text","detail"])).to_excel(xw, index=False, sheet_name="Issues")
        meta = {"generated_at": datetime.now().isoformat(timespec="seconds"), "method": "text-pdf + line-parse (lite)", "notes": ["Cloud-lite build (no OCR, no Camelot)."]}
        pd.DataFrame([meta]).to_excel(xw, index=False, sheet_name="Metadata")
    bio.seek(0)
    report = {"method": "text-pdf + line-parse (lite)", "notes": ["Cloud-lite build (no OCR, no Camelot)."], "transactions": int(len(tx_df)), "unparsed": int(len(tx_df.attrs.get('unparsed', []))), "ignored": int(len(tx_df.attrs.get('ignored', [])))}
    return bio.read(), report
