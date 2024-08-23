import openpyxl
import xml.etree.ElementTree as ET
import html
import difflib


class MULTI_LAN:

    def __init__(self, original_file):
        self.original_file = original_file

        try:
            self.original_wb = openpyxl.load_workbook(self.original_file)
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
        self.col_a_values = [cell.value.lower() for cell in self.sheet['A'] if cell.value is not None][1:]
        print(f"从原始文档获取到key值,并转成小写,共计{len(self.col_a_values)}个\n")
        return self.col_a_values

    def get_countries_from_origin(self):
        """获取国家列表"""
        row1_values = list(self.sheet.values)[0]
        self.countries = [val for val in row1_values[1:] if val]
        print(f"获取到了国家列表：{self.countries}\n")
        return self.countries

    def clean_value(self, value):
        """通用的清理函数"""
        if isinstance(value, str):
            # 处理Excel中的"\n"文本
            value = value.replace('\\n', '\n')

            # 统一换行符
            value = value.replace('\r\n', '\n')
            value = value.replace('\r', '\n')

            # 去除所有反斜杠，但保留换行符
            value = value.replace('\\', '')

            # 去除引号和首尾空格
            value = value.strip('"').strip()

            # 处理每一行，去除多余空格
            value = ' '.join(line.strip() for line in value.split('\n'))

            # 去除连续的空行
            value = ' '.join(line for line in value.split('\n') if line.strip())

        return value

    def read_strings_from_xml(self):
        """从xml读取字符串"""
        self.all_values_from_xml = {}

        for country in self.countries:
            if country.lower() in ["values", "value"]:
                file_path = "res/values/strings.xml"
            else:
                file_path = f"res/values-{country}/strings.xml"

            try:
                tree = ET.parse(file_path)
                root = tree.getroot()

                key_value = {}
                name_list = []
                for string in root.findall('string'):
                    name = string.get('name')
                    if name:
                        name_list.append(name.lower())

                        if name.lower() in self.col_a_values:
                            value = html.unescape(string.text) if string.text else ""
                            value = self.clean_value(value)
                            key_value[name.lower()] = value

                diff = set(self.col_a_values).difference(set(name_list))
                if diff:
                    print(f"国家{country}中这些key：{diff}, 不存在于xml中")
                else:
                    print(f"国家{country}的key都存在于xml中")

                self.all_values_from_xml[country] = key_value

            except FileNotFoundError:
                print(f"xml文件 {file_path} 不存在！请检查命名或文件")
                continue
            except ET.ParseError:
                print(f"xml文件 {file_path} 解析失败！请检查格式")
                continue
        return self.all_values_from_xml

    def get_excel_value_by_key_and_country(self, key, country):
        """根据key和国家获取Excel中的值"""
        country_index = self.countries.index(country) + 2
        for row in self.sheet.iter_rows(min_row=2, min_col=1, max_col=country_index, max_row=self.sheet.max_row):
            if row[0].value and row[0].value.lower() == key:
                excel_value = row[country_index - 1].value
                return self.clean_value(excel_value)
        return None

    def compare(self, value1, value2):
        """直接比较两个已清理的值"""
        # 确保两者都是字符串
        if not isinstance(value1, str):
            value1 = str(value1) if value1 is not None else ""
        if not isinstance(value2, str):
            value2 = str(value2) if value2 is not None else ""

        ratio = difflib.SequenceMatcher(None, value1, value2).ratio()
        return ratio == 1

    def compare_and_write_to_excel(self):
        """比较并写入到excel"""
        wb_new = openpyxl.Workbook()
        ws_new = wb_new.active
        ws_new.title = "对比差异结果"

        wb_same = openpyxl.Workbook()
        ws_same = wb_same.active
        ws_same.title = "对比相同结果"

        ws_new.append(["Country", "Key", "Excel Value", "XML Value"])
        ws_same.append(["Country", "Key", "Excel Value", "XML Value", "二次校验结果"])

        row_num = 2
        for country, strings in self.all_values_from_xml.items():
            for key, xml_value in strings.items():
                excel_value = self.get_excel_value_by_key_and_country(key, country)
                if not self.compare(xml_value, excel_value):
                    ws_new.append([country, key, excel_value, xml_value])
                else:
                    ws_same.append([country, key, excel_value, xml_value, f'=D{row_num}=C{row_num}'])
                    row_num += 1

        wb_new.save("对比差异结果.xlsx")
        wb_same.save("对比相同结果.xlsx")
        print("\n差异已写入 对比差异结果.xlsx\n相同的结果已写入 对比相同结果.xlsx \n请核实两个文件")


if __name__ == '__main__':
    wb = MULTI_LAN("origin.xlsx")
    wb.get_key_from_origin()
    wb.get_countries_from_origin()
    wb.read_strings_from_xml()
    wb.compare_and_write_to_excel()
    input("\nPress Enter to exit...")
