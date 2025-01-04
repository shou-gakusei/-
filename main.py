import openpyxl


class Carbon_Emission:
    def __init__(self, year, emission_quantity):
        self.year = year
        # int
        self.emission_quantity = emission_quantity
        # float


class City:
    def __init__(self, city_name):
        self.name = city_name
        # string
        self.emission_data_list = []
        # list

    def expand_data(self, carbon_emission):
        self.emission_data_list.append(carbon_emission)

    def search_data(self, year):
        if not self.emission_data_list:
            return 0
        for emission_data in self.emission_data_list:
            if year == emission_data.year:
                return emission_data.emission_quantity


class Province:
    def __init__(self, name):
        self.city_list = []
        self.name = name

    def expand_city(self, city):
        self.city_list.append(city)

    def search_city(self, city_name):
        if not self.city_list:
            return 0
        for city in self.city_list:
            if city == city_name:
                return city


def list_exchange(list_input):
    # 交换
    for data_count in range(len(list_input) - 1):
        if list_input[data_count][1] <= list_input[data_count + 1][1]:
            temporary_variables_5 = list_input[data_count]
            list_input[data_count] = list_input[data_count + 1]
            list_input[data_count + 1] = temporary_variables_5


def list_sort(list_input):
    # 交换排序~~~
    for list_l in range(len(list_input)):
        list_exchange(list_input)


path = "homework.xlsx"
shell_start = 2
shell_end = 7414
workbook = openpyxl.load_workbook(path)
sheet = workbook['数据']
sheet_output = workbook['数据输出1']
sheet_output_pointer = 2
province_name_list = []
city_name_list = []
province_list = []
column_list = ['A', 'B', 'C', 'D']
for shell in sheet['A']:
    if (shell.value not in province_name_list
            and not shell.value == 'Province Name'):
        province_name_list.append(shell.value)
print(province_name_list)
for province_name in province_name_list:
    province_list.append(Province(province_name))

for pointer in range(shell_start, shell_end + 1):
    temporary_variables_1 = [sheet['A' + str(pointer)].value,
                             sheet['B' + str(pointer)].value]
    for province in province_list:
        if (province.name == temporary_variables_1[0] and
                not temporary_variables_1[1] in [city.name for city in province.city_list
                                                 if not province.city_list == []]):
            province.expand_city(City(temporary_variables_1[1]))
            break
for pointer in range(shell_start, shell_end + 1):
    temporary_variables_2 = [sheet['A' + str(pointer)].value,
                             sheet['B' + str(pointer)].value,
                             sheet['C' + str(pointer)].value,
                             sheet['D' + str(pointer)].value]
    for province in province_list:
        if province.name == temporary_variables_2[0]:
            for city in province.city_list:
                if city.name == temporary_variables_2[1]:
                    city.expand_data(Carbon_Emission(temporary_variables_2[2],
                                                     temporary_variables_2[3]))

# 第一题
temporary_variables_4 = []
for year in range(1997, 2017 + 1):
    temporary_variables_3 = []
    for province in province_list:
        province_carbon_emission_sum = 0
        for city in province.city_list:
            province_carbon_emission_sum += city.search_data(year)
        temporary_variables_3.append([province.name, province_carbon_emission_sum, year])
    temporary_variables_4.append(temporary_variables_3)
for year_data_count in range(len(temporary_variables_4)):
    year = temporary_variables_4[year_data_count][0][2]
    count = year - 1995
    sheet_output['A' + str(count)] = year
    co2_emission_max_10 = [[data[0], data[1]] for data in temporary_variables_4[year_data_count][0:10]]
    list_sort(co2_emission_max_10)
    for year_data in temporary_variables_4[year_data_count]:
        if year_data[1] >= co2_emission_max_10[9][1] and not [year_data[0], year_data[1]] in co2_emission_max_10:
            co2_emission_max_10[9] = [year_data[0], year_data[1]]
            list_sort(co2_emission_max_10)
    for co2_emission_max_10_count in range(len(co2_emission_max_10)):
        sheet_output['B' + str(year_data_count + 2)] = co2_emission_max_10[0][0]
        sheet_output['C' + str(year_data_count + 2)] = co2_emission_max_10[1][0]
        sheet_output['D' + str(year_data_count + 2)] = co2_emission_max_10[2][0]
        sheet_output['E' + str(year_data_count + 2)] = co2_emission_max_10[3][0]
        sheet_output['F' + str(year_data_count + 2)] = co2_emission_max_10[4][0]
        sheet_output['G' + str(year_data_count + 2)] = co2_emission_max_10[5][0]
        sheet_output['H' + str(year_data_count + 2)] = co2_emission_max_10[6][0]
        sheet_output['I' + str(year_data_count + 2)] = co2_emission_max_10[7][0]
        sheet_output['J' + str(year_data_count + 2)] = co2_emission_max_10[8][0]
        sheet_output['K' + str(year_data_count + 2)] = co2_emission_max_10[9][0]
    print(year, co2_emission_max_10)
workbook.save(path)


# 第二题
def int_to_chr(int_input):
    output = ''
    if int_input <= 26:
        output = chr(64 + int_input)
    elif 26 < int_input < 52:
        int_input -= 26
        output = 'A' + chr(int_input + 64)
    return output


sheet_output_2 = workbook['数据输出2']
for year in range(1997, 2017 + 1):
    city_max_50_list = []
    for province in province_list:
        for city in province.city_list:
            city_max_50_list.append([city.name, city.search_data(year)])
    list_sort(city_max_50_list)
    city_max_50_list = city_max_50_list[0:50]
    count = year - 1995
    sheet_output_2['A' + str(count)] = year
    for city_max_50_list_count in range(len(city_max_50_list)):
        sheet_output_2[int_to_chr(city_max_50_list_count + 2) + str(count)] = city_max_50_list[city_max_50_list_count][
            0]
    print(year, city_max_50_list)
workbook.save(path)

# 第三题
sheet_output_3 = workbook['数据输出3']
# 长三角
Yangtze_River_delta = ['上海市', '江苏省', '浙江省', '安徽省']
for year in range(1997, 2017 + 1):
    province_co2_sum = 0
    for province_name in Yangtze_River_delta:
        for province in province_list:
            if province.name == province_name:
                for city in province.city_list:
                    province_co2_sum += city.search_data(year)
        sheet_output_3['A' + str(year - 1995)] = year
        sheet_output_3['B' + str(year - 1995)] = province_co2_sum
# 京津冀
Beijing_Tianjin_Hebei_Urban_Agglomeration = ['北京市', '天津市', '河北省']
for year in range(1997, 2017 + 1):
    province_co2_sum = 0
    for province_name in Beijing_Tianjin_Hebei_Urban_Agglomeration:
        for province in province_list:
            if province.name == province_name:
                for city in province.city_list:
                    province_co2_sum += city.search_data(year)
    sheet_output_3['C' + str(year - 1995)] = province_co2_sum
# 成渝
Chengdu_Chongqing_urban_agglomeration = ['四川省', '重庆市']
for year in range(1997, 2017 + 1):
    province_co2_sum = 0
    for province_name in Chengdu_Chongqing_urban_agglomeration:
        for province in province_list:
            if province.name == province_name:
                for city in province.city_list:
                    province_co2_sum += city.search_data(year)
        sheet_output_3['D' + str(year - 1995)] = province_co2_sum
workbook.save(path)

# 第四题
Hebei = '河北省'
sheet_output_4 = workbook['数据输出4']
for year in range(1997, 2017 + 1):
    province_co2_sum = 0
    for province in province_list:
        if province.name == Hebei:
            for city in province.city_list:
                province_co2_sum += city.search_data(year)
    sheet_output_4['A' + str(year - 1995)] = year
    sheet_output_4['B' + str(year - 1995)] = province_co2_sum
workbook.save(path)
