# -*- coding: utf-8 -*-
# @Time    : 2024/1/2 22:11
# @Author  : Zhang Huan
# @Email   : johnhuan@whu.edu.cn
# QQ       : 248404941
# @File    : main.py
import datetime
import ppgnss.gnss_time
from docx import Document


def generate_school_calendar(year, month, day, week_list):
    yr, doy_start = ppgnss.gnss_time.ymd2doy(year, month, day)  # 第一周开始的时间
    for i in range(0, 20):
        week_name = "第%d周" % (i + 1)
        document.add_heading(week_name, 2)
        # table = document.add_table(rows=3, cols=7, style='Table Grid')
        table = document.add_table(rows=3, cols=7, style='Colorful List')
        for j in range(0, 7):
            doy = doy_start + i * 7 + j
            yr, mo, dy = ppgnss.gnss_time.doy2ymd(year, doy)
            week_day = week_list[datetime.date(yr, mo, int(dy)).isoweekday() - 1]
            hdr_cells_mo_dy = table.rows[0].cells
            hdr_cells_mo_dy[j].text = "%d月%d" % (mo, dy)
            hdr_cells_week_dy = table.rows[1].cells
            hdr_cells_week_dy[j].text = week_day
            hdr_cells_work = table.rows[2].cells
            hdr_cells_work[j].text = ""


if __name__ == '__main__':
    document = Document()
    document.add_heading("工作计划", 0)
    """
    第一学期
    """
    week_list = ["星期二", "星期三", "星期四", "星期五", "星期六", "星期日", "星期一"]
    document.add_heading("2023~2024学年第一学期", 1)
    year = 2023
    month = 9
    day = 4
    generate_school_calendar(year, month, day, week_list)

    """
    第二学期
    """
    week_list = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
    document.add_heading("2023~2024学年第二学期", 1)
    year = 2024
    month = 2
    day = 26
    generate_school_calendar(year, month, day, week_list)

    document.save("2023~2024年度工作计划.docx")
