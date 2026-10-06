"""Turn a month of billable work into PDF invoices and an Excel register.

Input is two CSV files that any bookkeeping export or spreadsheet can produce:
  work_log.csv  date, client_id, description, qty, unit_price
  clients.csv   client_id, name, contact, email, address, terms_days, tax_rate

    python invoices.py --month 2026-09
    python invoices.py --month 2026-09 --start-number 1042
"""

import argparse
import csv
import json
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

CENT = Decimal("0.01")
ACCENT = colors.HexColor("#1F3A5F")
MUTED = colors.HexColor("#6B7280")


def money(value):
    return f"${value:,.2f}"


@dataclass
class Line:
    day: date
    description: str
    qty: Decimal
    unit_price: Decimal

    @property
    def amount(self):
        return (self.qty * self.unit_price).quantize(CENT, ROUND_HALF_UP)


@dataclass
class Invoice:
    number: str
    client: dict
    issued: date
    lines: list = field(default_factory=list)

    @property
    def subtotal(self):
        return sum((l.amount for l in self.lines), Decimal("0"))

    @property
    def tax(self):
        rate = Decimal(self.client.get("tax_rate") or "0")
        return (self.subtotal * rate / 100).quantize(CENT, ROUND_HALF_UP)

    @property
    def total(self):
        return self.subtotal + self.tax

    @property
    def due(self):
        return self.issued + timedelta(days=int(self.client.get("terms_days") or 30))


def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def build_invoices(work_rows, clients, month, start_number, issued):
    """Group the month's rows by client into numbered invoices."""
    by_client = defaultdict(list)
    for row in work_rows:
        day = date.fromisoformat(row["date"].strip())
        if f"{day:%Y-%m}" != month:
            continue
        cid = row["client_id"].strip()
        if cid not in clients:
            raise ValueError(f"Unknown client_id {cid!r} on {row['date']}")
        by_client[cid].append(Line(day, row["description"].strip(),
                                   Decimal(row["qty"]), Decimal(row["unit_price"])))

    invoices = []
    for i, cid in enumerate(sorted(by_client, key=lambda c: clients[c]["name"])):
        inv = Invoice(f"INV-{start_number + i}", clients[cid], issued,
                      sorted(by_client[cid], key=lambda l: l.day))
        invoices.append(inv)
    return invoices


def render_pdf(inv, company, path):
    doc = SimpleDocTemplate(str(path), pagesize=LETTER, leftMargin=0.8 * inch,
                            rightMargin=0.8 * inch, topMargin=0.7 * inch,
                            bottomMargin=0.7 * inch, title=inv.number)
    base = ParagraphStyle("base", fontName="Helvetica", fontSize=9.5, leading=13)
    small = ParagraphStyle("small", parent=base, textColor=MUTED, fontSize=8.5)
    big = ParagraphStyle("big", parent=base, fontName="Helvetica-Bold", fontSize=22,
                         leading=26, textColor=ACCENT)
    bold = ParagraphStyle("bold", parent=base, fontName="Helvetica-Bold")
    right = ParagraphStyle("right", parent=base, alignment=2)

    c = inv.client
    story = []
    head = Table([[
        Paragraph(f"<b>{company['name']}</b><br/>{company['address'].replace(chr(10), '<br/>')}"
                  f"<br/>{company['email']}", base),
        Paragraph("INVOICE", ParagraphStyle("t", parent=big, alignment=2)),
    ]], colWidths=[3.6 * inch, 3.3 * inch])
    head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story += [head, Spacer(1, 0.3 * inch)]

    meta = Table([[
        Paragraph(f"<font color='#6B7280'>BILL TO</font><br/><b>{c['name']}</b><br/>"
                  f"{c['contact']}<br/>{c['address']}<br/>{c['email']}", base),
        Table([
            [Paragraph("Invoice #", small), Paragraph(inv.number, right)],
            [Paragraph("Issued", small), Paragraph(f"{inv.issued:%b %d, %Y}", right)],
            [Paragraph("Due", small), Paragraph(f"<b>{inv.due:%b %d, %Y}</b>", right)],
            [Paragraph("Amount due", small), Paragraph(f"<b>{money(inv.total)}</b>", right)],
        ], colWidths=[1.1 * inch, 1.6 * inch]),
    ]], colWidths=[4.2 * inch, 2.7 * inch])
    meta.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story += [meta, Spacer(1, 0.35 * inch)]

    rows = [["Date", "Description", "Qty", "Rate", "Amount"]]
    for l in inv.lines:
        rows.append([f"{l.day:%m/%d}", Paragraph(l.description, base),
                     f"{l.qty.normalize():f}", money(l.unit_price), money(l.amount)])
    table = Table(rows, colWidths=[0.6 * inch, 3.6 * inch, 0.6 * inch, 0.9 * inch, 1.2 * inch],
                  repeatRows=1)
    table.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 9),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 0), (-1, 0), ACCENT),
        ("FONT", (0, 1), (-1, -1), "Helvetica", 9.5),
        ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F3F5F8")]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story += [table, Spacer(1, 0.15 * inch)]

    tax_label = f"Tax ({Decimal(c.get('tax_rate') or '0').normalize():f}%)"
    totals = Table([
        ["Subtotal", money(inv.subtotal)],
        [tax_label, money(inv.tax)],
        ["Total due", money(inv.total)],
    ], colWidths=[1.5 * inch, 1.2 * inch], hAlign="RIGHT")
    totals.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
        ("FONT", (0, 0), (-1, 1), "Helvetica", 9.5),
        ("FONT", (0, 2), (-1, 2), "Helvetica-Bold", 11),
        ("LINEABOVE", (0, 2), (-1, 2), 1, ACCENT),
        ("TOPPADDING", (0, 2), (-1, 2), 6),
    ]))
    story += [totals, Spacer(1, 0.4 * inch)]

    story += [Paragraph("Payment", bold),
              Paragraph(company["payment_instructions"].replace("\n", "<br/>"), base),
              Spacer(1, 0.2 * inch),
              Paragraph(company.get("footer", ""), small)]
    doc.build(story)


def write_register(invoices, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    ws.append(["Invoice #", "Client", "Email", "Issued", "Due", "Lines",
               "Subtotal", "Tax", "Total", "Status"])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3A5F")
    for inv in invoices:
        ws.append([inv.number, inv.client["name"], inv.client["email"], inv.issued,
                   inv.due, len(inv.lines), float(inv.subtotal), float(inv.tax),
                   float(inv.total), "Sent"])
    last = ws.max_row
    ws.append(["Total", "", "", "", "", f"=SUM(F2:F{last})", f"=SUM(G2:G{last})",
               f"=SUM(H2:H{last})", f"=SUM(I2:I{last})", ""])
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
    for row in ws.iter_rows(min_row=2):
        row[3].number_format = row[4].number_format = "mm/dd/yyyy"
        for cell in row[6:9]:
            cell.number_format = '"$"#,##0.00'
    widths = [12, 28, 30, 12, 12, 8, 13, 11, 13, 10]
    for col, w in zip("ABCDEFGHIJ", widths):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    wb.save(path)


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--month", required=True, help="billing month, YYYY-MM")
    parser.add_argument("--data", default="sample_data", help="folder with the CSV files")
    parser.add_argument("--out", default="output")
    parser.add_argument("--start-number", type=int, default=1001)
    parser.add_argument("--issued", help="issue date, YYYY-MM-DD (default: today)")
    args = parser.parse_args()

    data = Path(args.data)
    company = json.loads((data / "company.json").read_text())
    clients = {r["client_id"].strip(): r for r in read_csv(data / "clients.csv")}
    issued = date.fromisoformat(args.issued) if args.issued else date.today()
    invoices = build_invoices(read_csv(data / "work_log.csv"), clients, args.month,
                              args.start_number, issued)
    if not invoices:
        print(f"No billable work found for {args.month}.")
        return

    out = Path(args.out) / args.month
    out.mkdir(parents=True, exist_ok=True)
    for inv in invoices:
        slug = "".join(ch for ch in inv.client["name"] if ch.isalnum())
        render_pdf(inv, company, out / f"{inv.number}_{slug}.pdf")
        print(f"  {inv.number}  {inv.client['name']:<28} {money(inv.total):>12}")
    write_register(invoices, out / f"register_{args.month}.xlsx")
    grand = sum(i.total for i in invoices)
    print(f"{len(invoices)} invoices, {money(grand)} total -> {out}")


if __name__ == "__main__":
    main()
