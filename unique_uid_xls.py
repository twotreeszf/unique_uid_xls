# encoding: utf-8
"""
File:       unique_uid_xls
Author:     twotrees.zf@gmail.com
Date:       2017/11/11 15:44
Desc:
"""

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

from openpyxl import load_workbook


class UidLocation:
    def __init__(self, sheet_name, cell):
        self.sheetName = sheet_name
        self.cell = cell


if __name__ == '__main__':
    workbookPath = u'站内签约主播规划.xlsx'
    print(u'正在打开文件：{0} 请稍等...'.format(workbookPath))
    workbook = load_workbook(filename=workbookPath)
    print(u'打开成功')

    print(u'正在提取uid ...')
    totalUids = {}
    for sheet in workbook.worksheets:
        title = sheet.title
        for h in range(ord('A'), ord('Z') + 1):
            h = chr(h)
            header = '{0}1'.format(h)
            value = sheet[header].value
            if value is None:
                print(u'「{0}」没有找到 UID'.format(title))
                break

            if value == u'UID':
                row = 2
                while True:
                    index = '{0}{1}'.format(h, str(row))
                    value = sheet[index].value
                    if value is None:
                        break

                    if type(value) == int:
                        pass
                    elif type(value) == long:
                        value = int(value)
                    elif type(value) == unicode:
                        value = int(value.encode('utf-8').strip())
                    else:
                        pass

                    location = UidLocation(title, unicode(index))
                    if value in totalUids:
                        totalUids[value].append(location)
                    else:
                        totalUids[value] = [location]
                    row += 1

                print(u'「{0}」找到 UID {1} 个'.format(title, row - 2))
                break

    print(u'分析重复...\n')

    valid = True
    for uid in totalUids:
        locations = totalUids[uid]
        if len(locations) > 1:
            valid = False
            print(u'发现重复UID: {0}'.format(uid))
            for loc in locations:
                print(u'位置：「{0}」->{1}'.format(loc.sheetName, loc.cell))
            print('\n')
    if valid:
        print(u'验证通过，没有重复UID')

    raw_input("Press Enter to continue...")