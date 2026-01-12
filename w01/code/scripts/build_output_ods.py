#!/usr/bin/env python3
import argparse
import csv
import shutil
import time
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.dom import Node

from odf.element import Element, Text as OdfText
from odf.opendocument import load
from odf.table import Table
from odf.text import P

BASE_FIELDS = [
    "Serial No",
    "Order No",
    "Order Date",
    "Customer Name",
    "Grand Total",
]

CLEAR_CELLS = [
    (4, 5),   # D5
    (4, 6),   # D6
    (14, 6),  # N6
    (18, 5),  # R5
    (18, 7),  # R7
]

CLEAR_RANGES = [
    (3, 13, 28),   # C13:C28
    (6, 13, 28),   # F13:F28
    (12, 13, 28),  # L13:L28
    (14, 13, 28),  # N13:N28
    (15, 13, 28),  # O13:O28
    (15, 29, 29),  # O29
]

COLUMN_LETTER_MAP = {
    "C": 3,
    "D": 4,
    "F": 6,
    "L": 12,
    "N": 14,
    "O": 15,
    "R": 18,
}


def iter_table_rows(table):
    for node in table.childNodes:
        if getattr(node, "tagName", None) == "table:table-row":
            yield node


def row_repeat_count(row):
    repeat = row.getAttribute("numberrowsrepeated")
    return int(repeat) if repeat else 1


def ensure_physical_row(table, index):
    current = 1
    for row in iter_table_rows(table):
        repeat = row_repeat_count(row)
        if current <= index <= current + repeat - 1:
            if repeat == 1:
                return row
            offset = index - current
            new_rows = []
            if offset > 0:
                before = clone_node(row)
                before.setAttribute("numberrowsrepeated", str(offset))
                new_rows.append(before)
            target = clone_node(row)
            target.setAttribute("numberrowsrepeated", "1")
            new_rows.append(target)
            remaining = repeat - offset - 1
            if remaining > 0:
                after = clone_node(row)
                after.setAttribute("numberrowsrepeated", str(remaining))
                new_rows.append(after)
            for new_row in new_rows:
                table.insertBefore(new_row, row)
            table.removeChild(row)
            return target
        current += repeat
    return None


def find_row_at_index(table, index):
    current = 1
    for row in iter_table_rows(table):
        repeat = row_repeat_count(row)
        if current <= index <= current + repeat - 1:
            return row
        current += repeat
    return None


def iter_cells(row):
    for node in row.childNodes:
        if getattr(node, "tagName", None) in ("table:table-cell", "table:covered-table-cell"):
            yield node


def cell_repeat_count(cell):
    repeat = cell.getAttribute("numbercolumnsrepeated")
    return int(repeat) if repeat else 1


def ensure_cell(row, col_index):
    current = 1
    for cell in iter_cells(row):
        repeat = cell_repeat_count(cell)
        if current <= col_index <= current + repeat - 1:
            if repeat == 1:
                return cell
            offset = col_index - current
            new_cells = []
            if offset > 0:
                before = clone_node(cell)
                before.setAttribute("numbercolumnsrepeated", str(offset))
                new_cells.append(before)
            target = clone_node(cell)
            target.setAttribute("numbercolumnsrepeated", "1")
            new_cells.append(target)
            remaining = repeat - offset - 1
            if remaining > 0:
                after = clone_node(cell)
                after.setAttribute("numbercolumnsrepeated", str(remaining))
                new_cells.append(after)
            for new_cell in new_cells:
                row.insertBefore(new_cell, cell)
            row.removeChild(cell)
            return target
        current += repeat
    return None


def clone_node(node):
    if isinstance(node, OdfText):
        return OdfText(node.data)
    if getattr(node, "tagName", None):
        cloned = Element(qname=node.qname, attributes=dict(node.attributes), check_grammar=False)
        for child in node.childNodes:
            cloned.appendChild(clone_node(child))
        return cloned
    if node.nodeType == Node.TEXT_NODE:
        return OdfText(node.data)
    if hasattr(node, "cloneNode"):
        return node.cloneNode(True)
    raise TypeError(f"Unsupported node type: {type(node)!r}")


def clear_cell_content(cell):
    if cell is None:
        return
    for child in list(cell.childNodes):
        cell.removeChild(child)
    for attr in ("value", "value-type", "date-value"):
        if attr in cell.attributes:
            cell.attributes.pop(attr, None)


def set_cell_text(cell, text):
    clear_cell_content(cell)
    if text == "":
        return
    cell.setAttribute("valuetype", "string")
    cell.addElement(P(text=text))


def normalize_text(value):
    return value.strip() if value is not None else ""


def normalize_number_text(value):
    if value is None:
        return ""
    cleaned = value.replace(",", "").replace("$", "").strip()
    return cleaned


def parse_decimal(value, record, field, logs):
    cleaned = normalize_number_text(value)
    if cleaned == "":
        return Decimal("0")
    try:
        return Decimal(cleaned)
    except (InvalidOperation, ValueError):
        logs.append(
            format_log_entry(
                "PARSE_ERROR",
                record,
                Field=field,
                Value=value,
            )
        )
        return Decimal("0")


def format_log_entry(reason, record, **extra):
    parts = [f"Reason={reason}"]
    for key in BASE_FIELDS:
        parts.append(f"{key}={record.get(key, '')}")
    for key, value in extra.items():
        parts.append(f"{key}={value}")
    return ",".join(parts)


def apply_page_break(target_row, break_attrs):
    for key, value in break_attrs.items():
        if key == "numberrowsrepeated":
            continue
        target_row.setAttribute(key, value)


def build_output_ods(input_csv, template_path, output_path, log_path):
    start_time = time.time()
    logs = []

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(template_path, output_path)

    doc = load(str(output_path))
    spreadsheet = doc.spreadsheet

    print_all = None
    for table in list(spreadsheet.getElementsByType(Table)):
        if table.getAttribute("name") == "PRINT_ALL":
            print_all = table
        else:
            spreadsheet.removeChild(table)
    if print_all is None:
        raise RuntimeError("PRINT_ALL sheet not found")

    master_rows = [ensure_physical_row(print_all, idx) for idx in range(1, 32)]
    template_rows = [clone_node(row) for row in master_rows]

    break_row = ensure_physical_row(print_all, 1000)
    break_attrs = dict(break_row.attributes) if break_row else {}

    total_records = 0

    with input_csv.open(newline="", encoding="utf-8") as csv_file:
        reader = csv.DictReader(csv_file)
        for record in reader:
            total_records += 1
            start_row = 1 + (total_records - 1) * 31

            if total_records == 1:
                block_rows = master_rows
            else:
                insert_before = find_row_at_index(print_all, start_row)
                block_rows = [clone_node(row) for row in template_rows]
                if insert_before is None:
                    for row in block_rows:
                        print_all.addElement(row)
                else:
                    for row in block_rows:
                        print_all.insertBefore(row, insert_before)
                apply_page_break(block_rows[0], break_attrs)

            for col, row_start, row_end in CLEAR_RANGES:
                for row_idx in range(row_start, row_end + 1):
                    cell = ensure_cell(block_rows[row_idx - 1], col)
                    clear_cell_content(cell)

            for col, row_idx in CLEAR_CELLS:
                cell = ensure_cell(block_rows[row_idx - 1], col)
                clear_cell_content(cell)

            order_date = normalize_text(record.get("Order Date", ""))
            if order_date:
                order_date = order_date.split(" ")[0]
            customer_name = normalize_text(record.get("Customer Name", ""))
            serial_no = normalize_text(record.get("Serial No", ""))
            order_no = normalize_text(record.get("Order No", ""))
            customer_phone = normalize_text(record.get("Customer Phone", ""))

            set_cell_text(ensure_cell(block_rows[4], COLUMN_LETTER_MAP["D"]), order_date)
            set_cell_text(ensure_cell(block_rows[5], COLUMN_LETTER_MAP["D"]), customer_name)
            set_cell_text(ensure_cell(block_rows[5], COLUMN_LETTER_MAP["N"]), serial_no)
            set_cell_text(ensure_cell(block_rows[4], COLUMN_LETTER_MAP["R"]), order_no)
            set_cell_text(ensure_cell(block_rows[6], COLUMN_LETTER_MAP["R"]), customer_phone)

            totals_sum = Decimal("0")
            for index in range(1, 17):
                row_index = 12 + index
                sku = normalize_text(record.get(f"Product {index} SKU", ""))
                name = normalize_text(record.get(f"Product {index} Name", ""))
                quantity = normalize_number_text(record.get(f"Product {index} Quantity", ""))
                price = normalize_number_text(record.get(f"Product {index} Price", ""))
                total_value_raw = record.get(f"Product {index} Total", "")
                total_value = normalize_number_text(total_value_raw)
                totals_sum += parse_decimal(total_value_raw, record, f"Product {index} Total", logs)

                target_row = block_rows[row_index - 1]
                set_cell_text(ensure_cell(target_row, COLUMN_LETTER_MAP["C"]), sku)
                set_cell_text(ensure_cell(target_row, COLUMN_LETTER_MAP["F"]), name)
                set_cell_text(ensure_cell(target_row, COLUMN_LETTER_MAP["L"]), quantity)
                set_cell_text(ensure_cell(target_row, COLUMN_LETTER_MAP["N"]), price)
                set_cell_text(ensure_cell(target_row, COLUMN_LETTER_MAP["O"]), total_value)

            grand_total_raw = record.get("Grand Total", "")
            grand_total = parse_decimal(grand_total_raw, record, "Grand Total", logs)
            set_cell_text(ensure_cell(block_rows[28], COLUMN_LETTER_MAP["O"]), f"${grand_total}")

            if totals_sum != grand_total:
                diff = grand_total - totals_sum
                logs.append(
                    format_log_entry(
                        "TOTAL_MISMATCH",
                        record,
                        **{
                            "Sum(Product Totals)": totals_sum,
                            "Diff": diff,
                        },
                    )
                )

            has_products_17_25 = False
            for index in range(17, 26):
                for field in (
                    f"Product {index} Name",
                    f"Product {index} SKU",
                    f"Product {index} Quantity",
                    f"Product {index} Price",
                    f"Product {index} Total",
                ):
                    if normalize_text(record.get(field, "")):
                        has_products_17_25 = True
                        break
                if has_products_17_25:
                    break

            if has_products_17_25:
                logs.append(format_log_entry("HAS_PRODUCT_17_25", record))

    if total_records > 0:
        end_row = total_records * 31
        print_range = f"'PRINT_ALL'.A1:'PRINT_ALL'.R{end_row}"
        print_all.setAttribute("print-ranges", print_range)

    doc.save(str(output_path))

    elapsed = time.time() - start_time
    logs.append(
        format_log_entry(
            "SUMMARY",
            {
                "Serial No": "",
                "Order No": "",
                "Order Date": "",
                "Customer Name": "",
                "Grand Total": "",
            },
            TotalRecords=total_records,
            ElapsedSeconds=f"{elapsed:.2f}",
            Output=str(output_path),
        )
    )

    if logs:
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with log_path.open("a", encoding="utf-8") as log_file:
            for entry in logs:
                log_file.write(entry + "\n")


def main():
    parser = argparse.ArgumentParser(description="Build output.ods from CSV using ODF template.")
    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help="Path to input CSV",
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=None,
        help="Path to templates.ods",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Path to output.ods",
    )
    parser.add_argument(
        "--log",
        type=Path,
        default=None,
        help="Path to log.txt",
    )
    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parents[2]
    input_csv = args.input or (base_dir / "in" / "input.csv")
    template_path = args.template or (base_dir / "in" / "templates.ods")
    output_path = args.output or (base_dir / "out" / "output.ods")
    log_path = args.log or (base_dir / "out" / "log.txt")

    if not input_csv.exists():
        raise SystemExit(f"Input CSV not found: {input_csv}")

    build_output_ods(input_csv, template_path, output_path, log_path)


if __name__ == "__main__":
    main()
