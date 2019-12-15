from pathlib import Path, PurePath
import xlwt
import datetime
from fileutils import list_xls, list_mmo, get_codes_from_report
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
        for item in get_mmo_items(invoice_path):
            for rep_name, rep_dict in reports.items():
                if item['morion_id'] in rep_dict['codes']:
                    rep_dict['items'].append({
                        'supplier': info['supplier'],
                        'number': info['number'],
                        'date': info['date'],
                        'morion_id': item['morion_id'],
                        'title': item['title'],
                        'maker': item['maker'],
                        'price': item['price'],
                        'count': item['count'],
                        'units': item['units'],
                    })

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
