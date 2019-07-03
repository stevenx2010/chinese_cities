import xlrd
import pytest
import read_xls

book = 0
sheet = 0
col_town = []
@pytest.fixture()
def fixture01(request):
    global book
    global sheet
    global col_town

    book = xlrd.open_workbook('baidu_data_201904_modified.xlsx')
    sheet = book.sheets()[0]
    #get the town column & convert number to int
    col_temp = sheet.col_slice(7)
    for i in range(len(col_temp)):
        if isinstance(col_temp[i].value, float):
            col_town.append(int(col_temp[i].value))
        else:
            col_town.append(col_temp[i].value)

    def fin():
        book.unload_sheet(0)
        print('\nWork book released')

    request.addfinalizer(fin)

@pytest.mark.usefixtures('fixture01')
class TestGenerateData:
    def test_generate_data(self):
        for i in range(len(col_town)):
            row_town = []
            #get the town row & convert number to int
            row_temp = sheet.row_slice(i)
            for j in range(len(row_temp)):
                if isinstance(row_temp[j].value, float):
                    row_town.append(int(row_temp[j].value))
                else:
                    row_town.append(row_temp[j].value)

            #print(row_town)

            province_list = [read_xls.generate_data(sheet.cell(i,2).value, sheet.cell(i,1).value, 0, 'p')]
            city_list = [read_xls.generate_data(sheet.cell(i,4).value, sheet.cell(i,3).value, 1, sheet.cell(i,1).value)]
            district_list = [read_xls.generate_data(sheet.cell(i,6).value, sheet.cell(i,5).value, 1, sheet.cell(i,3).value)]
            town_list = [read_xls.generate_data(sheet.cell(i,8).value, sheet.cell(i,7).value, 1, sheet.cell(i,5).value)]

            #print(province_list)
            #print(province_list[0]['value'])
            #print(town_list[0])
            assert row_town[1] == province_list[0]['value']
            assert row_town[3] == city_list[0]['value']
            assert row_town[5] == district_list[0]['value']
            assert row_town[7] == town_list[0]['value']

            #test parentVal
            assert row_town[5] == town_list[0]['parentVal']
            assert row_town[3] == district_list[0]['parentVal']
            assert row_town[1] == city_list[0]['parentVal']
            
