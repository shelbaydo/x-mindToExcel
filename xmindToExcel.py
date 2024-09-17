import xmindparser
import openpyxl

# 假设你的 Excel 标题在这个常量数组中
EXCEL_HEADERS = ["function module", "title", "preconditon", "Step", "Expected Result"]

# 递归函数遍历 XMind 的分支
def traverse_branch(branch, path=None):
    if path is None:
        path = []
    
    # 当前分支的标题
    title = branch.get('title', '')
    
    # 将当前分支的标题添加到路径中
    current_path = path + [title]
    
    # 初始化结果列表
    all_cases = []
    
    

    
    if 'topics' in branch:
        for sub_branch in branch['topics']:
            all_cases.extend(traverse_branch(sub_branch, current_path))
    else:
        # 没有子分支时，将完整路径添加到结果中
        all_cases.append(current_path)
    
    return all_cases



# 读取 XMind 文件并解析
def parse_xmind_to_cases(xmind_file):
    xmind_data = xmindparser.xmind_to_dict(xmind_file)
    sheets = xmind_data[0]['topic']['topics']  # 假设这是 XMind 的顶层结构
    
    all_cases = []
    for sheet in sheets:
        all_cases.extend(traverse_branch(sheet))
    
    return all_cases

# 将解析的数据写入 Excel 文件
def write_cases_to_excel(cases, excel_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    
    # 写入 Excel 表头
    sheet.append(EXCEL_HEADERS)
    
    # 写入每一行数据
    for case in cases:
        sheet.append(case)
    
    # 保存 Excel 文件
    workbook.save(excel_file)

# 主函数
def xmind_to_excel(xmind_file, excel_file):
    cases = parse_xmind_to_cases(xmind_file)
    write_cases_to_excel(cases, excel_file)

# 使用例子
xmind_file = 'xmindfiles\AVM.xmind'
excel_file = 'AVM_TESTCASE.xlsx'
xmind_to_excel(xmind_file, excel_file)
