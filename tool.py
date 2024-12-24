import os

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


source_folder_path = "D:\\download\\work\\物件収支\\SRD"
destination_folder_path = "D:\\download\\work\\物件収支\\SRW"
source_file_list = get_file_names_without_extension(source_folder_path)

matches = search_elements_in_files(destination_folder_path, source_file_list)

# 打印结果
for element, files in matches.items():
    if files:
        print(f"'{element}' 在以下文件中找到: {files}")
    else:
        print(f"'{element}' 未在任何文件中找到。")


