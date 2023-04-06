import openpyxl

# 读取文件并将数据整理成列表
def read_data_from_file(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        data_list = []
        data = {}
        responsible_person = None
        for line in f:
            line = line.strip()
            if not line:
                if data:
                    data["负责人"] = responsible_person
                    data_list.append(data)
                    data = {}
                responsible_person = None
                continue
            if "：" in line:
                key, value = line.split("：", 1)
                if not value.strip():  # 如果冒号后面的内容为空，则认为该行是负责人
                    responsible_person = key.strip()
                else:
                    data[key.strip()] = value.strip()
        if data:
            data["负责人"] = responsible_person
            data_list.append(data)
        return data_list


# 将数据写入 Excel 文件
def write_data_to_excel(data_list, excel_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # 写表头
    headings = ['日期', '时间', '教室', '专业年级班级', '课程名称', '迟到', '请假', '带早餐', '负责人','踢球','旷课']
    sheet.append(headings)

    # 写数据
    for data in data_list:
        row = []
        row.append(data.get('日期', ''))
        row.append(data.get('时间', ''))
        row.append(data.get('教室', ''))
        row.append(data.get('专业年级班级', ''))
        row.append(data.get('课程名称', '')or data.get('课程', ''))
        row.append(data.get('迟到', ''))
        row.append(data.get('请假', ''))
        row.append(data.get('带早餐', ''))
        row.append(data.get('负责人', ''))
        row.append(data.get('踢球', ''))
        row.append(data.get('旷课', ''))
        sheet.append(row)

    workbook.save(excel_file)
    print(f"已将数据写入 Excel 文件 {excel_file}")

# 测试代码
if __name__ == "__main__":
    data_list = read_data_from_file("data.txt")
    write_data_to_excel(data_list, "data.xlsx")