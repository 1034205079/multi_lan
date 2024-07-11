import openpyxl
import xml.etree.ElementTree as ET
import html  # 导入html模块用于还原HTML实体

class MULTI_LAN:

    def __init__(self, original_file):
        self.original_file = original_file  # 绑定在实例上

        try:
            self.original_wb = openpyxl.load_workbook(self.original_file)  # 加载原始excel文件
        except FileNotFoundError:
            print("原始excel文件不存在！请检查命名或文件")
            exit()

        try:
            self.sheet = self.original_wb["Sheet1"]
        except KeyError:
            print(f"工作表 Sheet1 不存在！请检查工作表名称")
            return

    def get_key_from_origin(self):
        """获取key值"""
        self.col_a_values = [cell.value.lower() for cell in self.sheet['A'] if cell.value is not None][1:]  # 获取列A的值,排除第一个值
        print(f"从原始文档获取到key值,并转成小写,共计{len(self.col_a_values)}个\n")
        return self.col_a_values

    def get_countries_from_origin(self):
        """获取国家列表"""
        row1_values = list(self.sheet.values)[0]  # 获取第一行的值
        self.countries = [val for val in row1_values[1:] if val]  # 排除第一个值和空值
        print(f"获取到了国家列表：{self.countries}\n")
        return self.countries

    def read_strings_from_xml(self):
        """从xml读取字符串"""
        self.all_values_from_xml = {}

        for country in self.countries:
            if country.lower() in ["values", "value"]:
                file_path = "res/values/strings.xml"
            else:
                file_path = f"res/values-{country}/strings.xml"  # 确定文件路径

            try:
                tree = ET.parse(file_path)  # 解析xml文件
                root = tree.getroot()  # 获取根节点

                key_value = {}
                name_list = []
                for string in root.findall('string'):  # 遍历所有string节点
                    name = string.get('name')
                    if name:
                        name_list.append(name.lower())  # 把name插入到列表

                        if name.lower() in self.col_a_values:  # 如果xml中的name值在key值中
                            value = html.unescape(string.text) if string.text else ""  # 获取值并还原HTML实体
                            key_value[name.lower()] = value  # 把对应的name和字符串存入字典

                diff = set(self.col_a_values).difference(set(name_list))  # 找出xml中没有的key值
                if diff:
                    print(f"国家{country}中这些key：{diff}, 不存在于xml中")
                else:
                    print(f"国家{country}的key都存在于xml中")

                self.all_values_from_xml[country] = key_value  # 再把国家代号作为键，刚才的字典作为value，避免重复

            except FileNotFoundError:
                print(f"xml文件 {file_path} 不存在！请检查命名或文件")
                continue
            except ET.ParseError:
                print(f"xml文件 {file_path} 解析失败！请检查格式")
                continue
        return self.all_values_from_xml

    def get_excel_value_by_key_and_country(self, key, country):
        """根据key和国家获取Excel中的值"""
        country_index = self.countries.index(country) + 2  # +2 because column index starts from 1 and first column is A
        for row in self.sheet.iter_rows(min_row=2, min_col=1, max_col=country_index, max_row=self.sheet.max_row):
            if row[0].value and row[0].value.lower() == key:
                return row[country_index - 1].value  # 返回对应excel中的值，方便调用
        return None

    def compare_and_write_to_excel(self):
        """比较并写入到excel"""
        wb_new = openpyxl.Workbook()
        ws_new = wb_new.active
        ws_new.title = "Differences"

        wb_same = openpyxl.Workbook()
        ws_same = wb_same.active
        ws_same.title = "same_result"

        ws_new.append(["Country", "Key", "Excel Value", "XML Value"])
        ws_same.append(["Country", "Key", "Excel Value", "XML Value", "二次校验结果"])

        row_num = 2
        for country, strings in self.all_values_from_xml.items():  # 拿到对应国家的字典
            for key, xml_value in strings.items():  # 再从字典拿到对应的xml值
                excel_value = self.get_excel_value_by_key_and_country(key, country)  # 调用函数,传入参数
                if xml_value != excel_value:  # 如果xml值和excel值不一样
                    ws_new.append([country, key, excel_value, xml_value])
                else:
                    ws_same.append([country, key, excel_value, xml_value, f'=D{row_num}=C{row_num}'])
                    row_num += 1

        wb_new.save("differences.xlsx")
        wb_same.save("same_result.xlsx")
        print("\n差异已写入 differences.xlsx\n相同的结果已写入 same_result.xlsx \n请核实两个文件")
        print("\n请计算两个文件中各个国家的key数量加上上方打印的是否等于原始文档的数量！！！")


if __name__ == '__main__':
    wb = MULTI_LAN("origin.xlsx")
    wb.get_key_from_origin()
    wb.get_countries_from_origin()
    wb.read_strings_from_xml()
    wb.compare_and_write_to_excel()
    input("\nPress Enter to exit...")
