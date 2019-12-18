from pathlib import Path, PurePath
import xlwt
import datetime
from decimal import Decimal
from fileutils import (
    list_xls, list_xls_recursively, list_mmo, get_codes_from_report,
    get_full_info_from_report, get_info_from_member_report, get_items_from_member_report,
)
from mmo import get_mmo_info, get_mmo_items


def make_reports(member_code, reports_path, invoices_path, result_path):
    reports = {}
    for report_path in list_xls(reports_path):
        reports[Path(report_path).name.replace('.xls', '')] = {
            'codes': get_codes_from_report(report_path),
            'items': [],
        }

    for invoice_path in list_mmo(invoices_path):
        info = get_mmo_info(invoice_path)
        morion_id_unique_index = {}
        for item in get_mmo_items(invoice_path):
            for rep_name, rep_dict in reports.items():
                if item['morion_id'] in rep_dict['codes']:
                    if morion_id_unique_index.get(item['morion_id'], None) is None:
                        rep_row = {
                            'supplier': info['supplier'],
                            'number': info['number'],
                            'date': info['date'],
                            'morion_id': item['morion_id'],
                            'title': item['title'],
                            'maker': item['maker'],
                            'price': item['price'],
                            'count': item['count'],
                            'units': item['units'],
                        }
                        morion_id_unique_index[item['morion_id']] = rep_row
                        rep_dict['items'].append(rep_row)
                    else:  # If morion id inserted multiple times to invoice
                        rep_row = morion_id_unique_index[item['morion_id']]
                        rep_row['count'] += item['count']

    for rep_name, rep_info in reports.items():
        if len(rep_info['items']) != 0:
            date_style = xlwt.XFStyle()
            date_style.num_format_str = "DD.MM.YYYY"
            w = xlwt.Workbook()
            ws = w.add_sheet('report')
            ws.write(0, 0, 'member_code')
            ws.write(0, 1, member_code)
            ws.write(1, 0, 'date')
            ws.write(1, 1, datetime.date.today(), date_style)
            for r, d in enumerate(rep_info['items'], start=2):
                ws.write(r, 0, d['supplier'])
                ws.write(r, 1, d['number'])
                ws.write(r, 2, d['date'], date_style)
                ws.write(r, 3, d['morion_id'])
                ws.write(r, 4, d['title'])
                ws.write(r, 5, d['maker'])
                ws.write(r, 6, d['price'])
                ws.write(r, 7, d['count'])
                ws.write(r, 8, d['units'])
            new_rep_name = Path(PurePath(result_path).joinpath(rep_name + '.xls'))
            if new_rep_name.exists():
                new_rep_name.unlink()
            w.save(str(new_rep_name))


def compose_report(start_date, end_date, report_path, members_path, result_path):
    composed_report = []
    unique_index = []
    members = []

    report_info = get_full_info_from_report(report_path)
    report_info_by_code = {}
    for code, title in report_info:
        if report_info_by_code.get(code, None) is None:
            report_info_by_code[code] = title
    report_codes = list(code for code, title in report_info)

    for member_report in list_xls_recursively(members_path):
        member_info = get_info_from_member_report(member_report)
        member_report_name = PurePath(member_report).name.lower().replace('.xls', '')
        for item in get_items_from_member_report(member_report):
            if item['morion_id'] in report_codes and item['invoice_date'] >= start_date and item['invoice_date'] < end_date:
                item_unique_key = (member_info['code'], item['supplier'], item['invoice_number'], item['morion_id'])
                if item_unique_key not in unique_index:
                    unique_index.append(item_unique_key)
                    if member_info['code'] not in members:
                        members.append(member_info['code'])
                    composed_report.append(dict(
                        item,
                        member_code=member_info['code'],
                        member_report=member_report_name,
                        member_report_date=member_info['date'],
                    ))
    w = xlwt.Workbook()
    date_style = xlwt.XFStyle()
    date_style.num_format_str = "DD.MM.YYYY"

    ws = w.add_sheet('table')
    ws.write(0, 0, 'start date')
    ws.write(0, 1, start_date, date_style)
    ws.write(1, 0, 'end date')
    ws.write(1, 1, end_date, date_style)

    col = {
        'morion_id': 0,
        'title': 1,
        'total_count': 2,
        'total_sum': 3,
    }
    ws.write(2, col['morion_id'], 'morion_id')
    ws.write(2, col['title'], 'title')
    ws.write(2, col['total_count'], 'total count')
    ws.write(2, col['total_sum'], 'total sum')
    m_col = {}
    for c, member_code in enumerate(members, start=0):
        m_col[member_code] = {
            'count': (c * 2) + 4,
            'sum': (c * 2) + 5,
        }
        ws.write(2, m_col[member_code]['count'], member_code + ' count')
        ws.write(2, m_col[member_code]['sum'], member_code + ' sum')

    # Prepare matrix for single write to xls
    morion_rows = {}
    for item in composed_report:
        if morion_rows.get(item['morion_id'], None) is None:
            row_l = [item['morion_id'], item['title'], 0, Decimal(0)]
            for member in members:
                row_l.append(0)
                row_l.append(Decimal(0))
            morion_rows[item['morion_id']] = row_l
        r = morion_rows[item['morion_id']]
        r[col['total_count']] += item['count']
        r[col['total_sum']] += Decimal(item['count'] * item['price']).quantize(Decimal('1.00'))
        r[m_col[item['member_code']]['count']] += item['count']
        r[m_col[item['member_code']]['sum']] += Decimal(item['count'] * item['price']).quantize(Decimal('1.00'))

    # Write matrix
    for r, item in enumerate(sorted(morion_rows.values(), key=lambda i: i[1]), start=3):  # Sort by title
        for c, v in enumerate(item, start=0):
            ws.write(r, c, v)

    ws = w.add_sheet('by_invoice')
    ws.write(0, 0, 'start date')
    ws.write(0, 1, start_date, date_style)
    ws.write(1, 0, 'end date')
    ws.write(1, 1, end_date, date_style)

    ws.write(2, 0, 'member_code')
    ws.write(2, 1, 'member_report')
    ws.write(2, 2, 'member_report_date')
    ws.write(2, 3, 'supplier')
    ws.write(2, 4, 'number')
    ws.write(2, 5, 'invoice_date')
    ws.write(2, 6, 'morion_id')
    ws.write(2, 7, 'title')
    ws.write(2, 8, 'maker')
    ws.write(2, 9, 'price')
    ws.write(2, 10, 'count')
    ws.write(2, 11, 'units')

    for r, item in enumerate(sorted(composed_report, key=lambda i: (i['invoice_date'], i['title'], i['member_code'])), start=3):
        ws.write(r, 0, item['member_code'])
        ws.write(r, 1, item['member_report'])
        ws.write(r, 2, item['member_report_date'], date_style)
        ws.write(r, 3, item['supplier'])
        ws.write(r, 4, item['invoice_number'])
        ws.write(r, 5, item['invoice_date'], date_style)
        ws.write(r, 6, item['morion_id'])
        ws.write(r, 7, item['title'])
        ws.write(r, 8, item['maker'])
        ws.write(r, 9, item['price'])
        ws.write(r, 10, item['count'])
        ws.write(r, 11, item['units'])

    w.get_sheet(0)

    report_name = PurePath(report_path).name.lower().replace('.xls', '')
    new_rep_name = Path(PurePath(result_path).joinpath(report_name + '.xls'))
    if new_rep_name.exists():
        new_rep_name.unlink()
    w.save(str(new_rep_name))
