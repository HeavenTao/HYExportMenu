#!/usr/bin/env python3
"""
根据menu.json生成Excel文件
支持单元格合并，过滤无url的行，文本左对齐垂直居中
一级菜单包含图标（与名称换行显示）
"""
import json
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment


# 全局对齐样式：水平居左，垂直居中
CELL_ALIGNMENT = Alignment(horizontal='left', vertical='center')
# 表头样式
HEADER_ALIGNMENT = Alignment(horizontal='left', vertical='center')


def load_menu_data(filepath):
    """加载menu.json数据"""
    with open(filepath, 'r', encoding='utf-8') as f:
        return json.load(f)


def get_first_level_display(first_level):
    """获取一级菜单显示文本（名称+图标，换行）"""
    name = first_level.get('Name', '')
    icon = first_level.get('Icon')
    
    if icon and str(icon).strip():
        return f"{name}\n{icon}"
    return name


def generate_menu_data(menu_data):
    """
    根据menu.json生成菜单数据，过滤掉没有url地址的行
    返回:
        rows: 列表，每个元素是 (一级菜单显示文本, 二级菜单名, 三级菜单名, 地址)
        first_merges: 一级菜单合并范围 [(显示文本, 起始行, 结束行), ...]
        second_merges: 二级菜单合并范围 [(菜单名, 起始行, 结束行), ...]
    """
    rows = []
    
    for first_level in menu_data:
        first_display = get_first_level_display(first_level)
        second_menus = first_level.get('SubMenus', [])
        
        for second_level in second_menus:
            second_name = second_level.get('Name', '')
            third_menus = second_level.get('SubMenus', [])
            
            for third_level in third_menus:
                third_name = third_level.get('Name', '')
                url = third_level.get('NavigateUrl', '')
                
                # 只保留有url的行（非空字符串且非None）
                if url and url.strip():
                    rows.append((first_display, second_name, third_name, url))
    
    # 计算合并范围
    first_merges = []
    second_merges = []
    
    if not rows:
        return rows, first_merges, second_merges
    
    # 计算一级菜单合并范围
    current_first = rows[0][0]
    first_start = 3  # Excel行号从3开始（1=说明，2=表头）
    
    for idx, row in enumerate(rows):
        if row[0] != current_first:
            first_end = idx + 2  # 上一行的行号
            first_merges.append((current_first, first_start, first_end))
            current_first = row[0]
            first_start = idx + 3  # 当前行的行号
    
    # 最后一个一级菜单
    first_merges.append((current_first, first_start, len(rows) + 2))
    
    # 计算二级菜单合并范围
    current_second = rows[0][1]
    current_first_for_second = rows[0][0]
    second_start = 3
    
    for idx, row in enumerate(rows):
        # 当二级菜单变化，或一级菜单变化（同名二级菜单在不同一级下）
        if row[1] != current_second or row[0] != current_first_for_second:
            second_end = idx + 2
            second_merges.append((current_second, second_start, second_end))
            current_second = row[1]
            current_first_for_second = row[0]
            second_start = idx + 3
    
    # 最后一个二级菜单
    second_merges.append((current_second, second_start, len(rows) + 2))
    
    return rows, first_merges, second_merges


def set_cell_value(ws, row, col, value, font=None):
    """设置单元格值和对齐方式"""
    cell = ws.cell(row=row, column=col, value=value)
    cell.alignment = CELL_ALIGNMENT
    if font:
        cell.font = font
    return cell


def setup_worksheet(ws):
    """设置工作表格式"""
    # 设置列宽
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 47
    ws.column_dimensions['D'].width = 66
    ws.column_dimensions['E'].width = 48
    
    # 设置说明行
    ws.merge_cells('A1:C1')
    cell = ws.cell(row=1, column=1, value='说明：绿色标注为要部署的菜单')
    cell.alignment = CELL_ALIGNMENT
    
    # 设置表头
    headers = ['一级菜单', '二级菜单', '三级菜单', '地址']
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.alignment = HEADER_ALIGNMENT
        cell.font = Font(bold=True)


def create_excel(menu_data, output_path):
    """创建Excel文件"""
    # 创建新的工作簿
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    
    # 设置工作表格式
    setup_worksheet(ws)
    
    # 生成菜单数据
    rows, first_merges, second_merges = generate_menu_data(menu_data)
    
    # 写入数据
    for idx, (first, second, third, url) in enumerate(rows, start=3):
        set_cell_value(ws, idx, 1, first)
        set_cell_value(ws, idx, 2, second)
        set_cell_value(ws, idx, 3, third)
        set_cell_value(ws, idx, 4, url)
    
    # 一级菜单合并单元格
    for first_name, start_row, end_row in first_merges:
        if end_row > start_row:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
    
    # 二级菜单合并单元格
    for second_name, start_row, end_row in second_merges:
        if end_row > start_row:
            ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
    
    # 保存文件
    wb.save(output_path)
    print(f"Excel文件已生成: {output_path}")
    print(f"总共 {len(rows)} 行菜单数据")
    print(f"一级菜单合并: {len(first_merges)} 处")
    print(f"二级菜单合并: {len(second_merges)} 处")


def main():
    menu_data = load_menu_data('menu.json')
    create_excel(menu_data, '菜单导出.xlsx')


if __name__ == '__main__':
    main()
