import os
import pandas as pd
from openpyxl import load_workbook

def get_file_names_without_extension(folder_path):
    """
    获取文件夹下所有文件名（去掉扩展名）。

    :param folder_path: 文件夹路径
    :return: 文件名（无扩展）的列表
    """
    file_names = []

    # 遍历文件夹
    for root, _, files in os.walk(folder_path):
        for file_name in files:
            # 分离文件名和扩展名
            name_without_ext, _ = os.path.splitext(file_name)
            file_names.append(name_without_ext)

    return file_names

def search_elements_in_files(folder_path, elements):
    """
    检查列表中的元素是否存在于文件夹下的所有文件内容中。

    :param folder_path: 文件夹路径
    :param elements: 要检索的字符串列表
    :return: 包含元素和对应文件的字典
    """
    result = {element: [] for element in elements}  # 用于存储结果

    # 遍历文件夹中的所有文件
    for root, _, files in os.walk(folder_path):
        for file_name in files:
            file_path = os.path.join(root, file_name)

            # 尝试读取文件内容
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()

                    # 检查每个元素是否在文件内容中
                    for element in elements:
                        if element in content:
                            result[element].append(file_path)
            except Exception as e:
                print(f"无法读取文件: {file_path}, 错误: {e}")

    return result


def export_to_excel(data, headers, output_path, sheet_name="Sheet1"):
    """
    将数据导出到 Excel 文件，如果文件存在则添加新工作表，如果文件不存在则创建新文件。

    :param data: 数据列表
    :param headers: 列标题列表
    :param output_path: Excel 文件路径
    :param sheet_name: 工作表名称
    """
    # 创建 DataFrame
    df = pd.DataFrame(data, columns=headers)

    if os.path.exists(output_path):
        # 文件存在，加载现有的 Excel 文件
        with pd.ExcelWriter(output_path, engine="openpyxl", mode="a") as writer:
            # 打开现有的工作簿
            book = load_workbook(output_path)

            # 检查工作表是否已存在
            if sheet_name in book.sheetnames:
                print(f"工作表 '{sheet_name}' 已存在，正在添加数据到该工作表...")
            else:
                print(f"工作表 '{sheet_name}' 不存在，正在创建新工作表...")

            # 将数据添加到新的工作表中
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            print(f"数据已成功保存到现有文件 {output_path}，工作表 '{sheet_name}'")

    else:
        # 文件不存在，直接创建并保存
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            print(f"新文件已创建并保存到 {output_path}，工作表 '{sheet_name}'")


if __name__=="__main__":

    source_folder_path = r"D:\download\work\SP"
    destination_folder_path = r"D:\download\work\物件収支\SRU"
    source_file_list = get_file_names_without_extension(source_folder_path)

    matches = search_elements_in_files(destination_folder_path, source_file_list)

    paris_list = []
    source_extension = os.path.basename(source_folder_path)
    destination_extension = os.path.basename(destination_folder_path)
    headers = [destination_extension,source_extension,]
    sheet_name = destination_extension + "→" + source_extension
    #将ID名存入list中
    for element, files in matches.items():
        if files:
         file_name = os.path.splitext(os.path.basename(files[0]))[0]
         paris_list.append([file_name,element])

    export_to_excel(paris_list,headers,r"D:\download\work\物件収支\物件収支_output.xlsx",sheet_name)