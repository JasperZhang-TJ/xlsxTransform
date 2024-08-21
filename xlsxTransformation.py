import pandas as pd

def generate_nested_mapping(config_path):
    """
    从 config.xlsx 生成嵌套字典结构，并处理元素映射。

    Args:
    - config_path (str): 配置文件（Excel）的路径。

    Returns:
    - nested_mapping (dict): 一个三层嵌套的字典结构，格式为
      {
          (source_sheet, target_sheet): {
              (source_title, target_title): {
                  'rule': trans_type,
                  'elements': [
                      (source_element, target_element)
                  ]
              }
          }
      }
    """
    nested_mapping = {}

    # 读取配置文件的所有sheet
    config_df = pd.read_excel(config_path, sheet_name=None)

    for sheet_name, df in config_df.items():
        # 获取每个sheet对应的source和target sheet名
        left_sheet, right_sheet = sheet_name.split('-')
        left_sheet = left_sheet.strip()
        right_sheet = right_sheet.strip()

        if (left_sheet, right_sheet) not in nested_mapping:
            nested_mapping[(left_sheet, right_sheet)] = {}

        for _, row in df.iterrows():
            left_title = row['LeftTitle'].strip()
            right_title = row['RightTitle'].strip()
            trans_type = row['TransType'].strip()

            if (left_title, right_title) not in nested_mapping[(left_sheet, right_sheet)]:
                nested_mapping[(left_sheet, right_sheet)][(left_title, right_title)] = {
                    'rule': trans_type,
                    'elements': []
                }

            # 处理TransType后的元素映射，存储为元组列表
            elements_mapping = nested_mapping[(left_sheet, right_sheet)][(left_title, right_title)]['elements']
            if trans_type == 'Translate':
                for element_pair in row[3:]:  # 从第四列开始处理映射关系
                    if pd.notna(element_pair):
                        left_element, right_element = element_pair.split('-')
                        elements_mapping.append((left_element.strip(), right_element.strip()))

    return nested_mapping

def reverse_nested_mapping(nested_mapping):
    """
    生成一个新的嵌套字典，颠倒元组内部元素的位置。

    Args:
    - nested_mapping (dict): 原始嵌套字典。

    Returns:
    - reversed_mapping (dict): 元组内部元素颠倒后的新嵌套字典。
    """
    reversed_mapping = {}

    for (source_sheet, target_sheet), titles_mapping in nested_mapping.items():
        reversed_mapping[(target_sheet, source_sheet)] = {}

        for (source_title, target_title), mapping_info in titles_mapping.items():
            reversed_mapping[(target_sheet, source_sheet)][(target_title, source_title)] = {
                'rule': mapping_info['rule'],
                'elements': [(target, source) for source, target in mapping_info['elements']]
            }

    return reversed_mapping

def read_source_data(source_path):
    """
    从 source.xlsx 读取数据并生成字典结构。

    Args:
    - source_path (str): 源文件（Excel）的路径。

    Returns:
    - source_data (dict): 一个嵌套的字典结构，格式为：
      {
          sheet_name: {
              title_name: [element1, element2, ..., elementN]
          }
      }
    """
    source_data = {}

    # 读取源文件的所有sheet
    source_df_dict = pd.read_excel(source_path, sheet_name=None)

    for sheet_name, df in source_df_dict.items():
        source_data[sheet_name] = {}

        for title in df.columns:
            # 获取列中的所有元素，将NaN替换为None
            elements = df[title].fillna('null').tolist()
            source_data[sheet_name][title] = elements

    return source_data

def find_unmapped_data(source_data, nested_mapping):
    """
    识别 source_data 中无法找到映射的部分，并打印其路径。

    Args:
    - source_data (dict): 源数据字典。
    - nested_mapping (dict): 嵌套映射字典。

    Returns:
    - None
    """
    for sheet_name, titles in source_data.items():
        # 查找匹配的 sheet
        matched_sheet = None
        for (source_sheet, target_sheet) in nested_mapping.keys():
            if sheet_name == source_sheet:
                matched_sheet = (source_sheet, target_sheet)
                break

        if not matched_sheet:
            print(f"未找到映射的 Sheet: {sheet_name}")
            continue

        for title, elements in titles.items():
            # 查找匹配的 title
            matched_title = None
            for (source_title, target_title) in nested_mapping[matched_sheet].keys():
                if title == source_title:
                    matched_title = (source_title, target_title)
                    break

            if not matched_title:
                print(f"未找到映射的 Title: {sheet_name} -> {title}")
                continue

def map_elements_by_rule(elements, rule, element_mapping):
    """
    根据不同的映射规则处理元素，并生成新的元素列表。

    Args:
    - elements (list): 需要映射的源元素列表。
    - rule (str): 映射规则。
    - element_mapping (list of tuples): 映射关系，格式为 [(source_element, target_element), ...]

    Returns:
    - mapped_elements (list): 映射后的目标元素列表。
    """
    if rule == 'Copy':
        # 直接复制元素
        return elements
    elif rule == 'Translate':
        # 使用映射关系翻译元素
        mapping_dict = dict(element_mapping)
        return [mapping_dict.get(elem, 'null') for elem in elements]
    else:
        # 可以在这里添加更多的映射规则
        raise ValueError(f"未知的映射规则: {rule}")

def generate_new_data(source_data, nested_mapping):
    """
    根据映射关系生成新的数据字典。

    Args:
    - source_data (dict): 源数据字典。
    - nested_mapping (dict): 嵌套映射字典。

    Returns:
    - new_data (dict): 生成的目标数据字典，格式为：
      {
          target_sheet: {
              target_title: [mapped_element1, mapped_element2, ..., mapped_elementN]
          }
      }
    """
    new_data = {}

    for sheet_name, titles in source_data.items():
        # 查找匹配的 sheet
        matched_sheet = None
        for (source_sheet, target_sheet) in nested_mapping.keys():
            if sheet_name == source_sheet:
                matched_sheet = (source_sheet, target_sheet)
                break

        if not matched_sheet:
            continue

        if target_sheet not in new_data:
            new_data[target_sheet] = {}

        for title, elements in titles.items():
            # 查找匹配的 title
            matched_title = None
            for (source_title, target_title) in nested_mapping[matched_sheet].keys():
                if title == source_title:
                    matched_title = (source_title, target_title)
                    break

            if not matched_title:
                continue

            # 根据不同的映射规则生成目标元素列表
            rule = nested_mapping[matched_sheet][matched_title]['rule']
            element_mapping = nested_mapping[matched_sheet][matched_title]['elements']
            mapped_elements = map_elements_by_rule(elements, rule, element_mapping)

            new_data[target_sheet][target_title] = mapped_elements

    return new_data

def save_to_excel(new_data, target_path):
    """
    将生成的目标数据字典保存到新的 Excel 表格中。

    Args:
    - new_data (dict): 生成的目标数据字典。
    - target_path (str): 保存新 Excel 文件的路径。

    Returns:
    - None
    """
    with pd.ExcelWriter(target_path, engine='openpyxl') as writer:
        for sheet_name, titles in new_data.items():
            # 创建一个 DataFrame 用于存储该 sheet 的数据
            df = pd.DataFrame(titles)
            # 将 DataFrame 写入 Excel 文件中的相应 sheet
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"新表格已保存至 {target_path}")

def main():
    config_path = 'config.xlsx'  # 配置文件路径
    target_path = 'source.xlsx'  # 保存新 Excel 文件的路径

    # 生成嵌套字典
    nested_mapping = generate_nested_mapping(config_path)

    # 询问用户映射方向
    direction = input("请选择映射方向：1. 从左往右 2. 从右往左\n请输入选项（1 或 2）：")

    if direction == '2':
        # 生成颠倒后的嵌套字典
        nested_mapping = reverse_nested_mapping(nested_mapping)
        print("已生成从右往左的映射结构。")
    else:
        print("保持从左往右的映射结构。")

    # 打印结果
    print(nested_mapping)


    #--------------------------------------读取原表格部分开发
    source_path = 'target.xlsx'  # 源文件路径

    # 读取源文件的数据
    source_data = read_source_data(source_path)

    # 打印结果
    print("源文件结构和内容",source_data)

    # 识别并打印未映射的部分
    find_unmapped_data(source_data, nested_mapping)

    # 根据映射关系生成新的数据字典
    new_data = generate_new_data(source_data, nested_mapping)

    # 打印生成的目标数据字典
    print("生成的目标数据字典：", new_data)

    # 将生成的数据保存到新的 Excel 文件中
    save_to_excel(new_data, target_path)

if __name__ == "__main__":
    main()
