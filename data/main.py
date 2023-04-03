# coding:utf-8
import os
import re
import pandas as pd
import logging

# 定义文件输出路径
output_path = os.path.dirname(__file__)

logging.basicConfig(level=logging.INFO,  filename=os.path.join(output_path,"output",'running_error.log'), encoding="utf-8", filemode='w')


def get_excel_file_list():
    """
        获取当前文件夹下的所有.xlsx文件名路径，存入列表中

    Returns:
        _type_:list 获取当前文件夹下的所有.xlsx文件地址列表
    """
    # 获取当前文件夹下的所有.xlsx文件名路径，存入列表中
    base_path = os.path.abspath(os.path.dirname(__file__))
    file_list = [file_name for file_name in os.listdir(base_path) if re.match(r'.*\.xlsx', file_name)]
    file_path_list = [os.path.join(base_path, file_name) for file_name in file_list]
    # print(file_path_list)
    return file_list, file_path_list

# 按照指定分隔符将字符串分割成列表
def split_str(file_name, row_index, target_str, delimiter_list):
    """
        按照指定分隔符将字符串分割成列表。
        如果不能切割成指定的列表长度，则返回空, 并打印相关信息到日志文件中


    Args:
        target_str (str): 将要被分割的字符串
        delimiter_list (list): 用于分割字符串的字符列表

    Returns:
        list: 分割后返回的内容
    """
    # assert len(target_str) > 0
    # assert len(delimiter_list) > 0
    try:
        result_str = re.split('|'.join(delimiter_list), target_str)
        fact = "本案现已审理终结。" + result_str[1] + "本院认为"
        criteria = "本院认为：" + result_str[2] + "，判决如下："
        return fact, criteria
    except:
        logging.info("{}的第{}行分割异常！".format(file_name, row_index))
        return "None", "None"


if __name__ == "__main__":


    # 读取当前文件夹中的额所有excel文件
    file_list, file_path_list = get_excel_file_list()

    # 定义分隔符
    delimiter_list1 = ["本案现已审理终结。", "本院认为，", "，判决如下"]
    delimiter_list2 = ["本案现已审理终结", "本院认为", "判决如下"]

    # 新建一个分割后的文件模板
    columns = ["CASENO", "courtname", "judgeresult", "content", "fact", "criteria"]
    data = {
        "CASENO": [],
        "courtname": [],
        "judgeresult": [],
        "content": [],
        "fact": [],
        "criteria": []
    }

    # 遍历每一个excel文件列表
    for file_index in range(len(file_path_list)):
        # 读取excel文件内容
        df = pd.read_excel(file_path_list[file_index])
        # 逐行遍历excel文件内容
        for row_index, row in df.iterrows():
            data["CASENO"].append(row["CASENO"])
            data["courtname"].append(row["courtname"])
            data["judgeresult"].append(row["judgeresult"])
            data["content"].append(row["content"])
            fact, criteria = split_str(file_name = file_list[file_index], row_index=row_index + 2, target_str=row["content"], delimiter_list=delimiter_list1)
            data["fact"].append(fact)
            data["criteria"].append(criteria)
            content = row["content"]
    
        new_excel = pd.DataFrame(data, columns=columns)
        # 使用 ExcelWriter 上下文管理器
        with pd.ExcelWriter(os.path.join(output_path,"output", "new_{}".format(file_list[file_index])), engine='xlsxwriter') as writer:
            # 遍历 DataFrame 的每一行，逐行写入 Excel 文件
            new_excel.to_excel(writer, sheet_name='Sheet1', index=False)
            # 保存文件
            writer.save()
        writer.close()
    print('文件处理完成')
