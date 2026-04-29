#!/usr/bin/env python3
"""
根据menu.json生成Excel文件，参照模板.xlsx的格式
"""
import json
import openpyxl
from openpyxl import Workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from copy import copy


def load_menu_data(filepath):
    """加载menu.json数据"""
    with open(filepath, 'r', encoding='utf-8') as f:
        return json.load(f)


def copy_cell_style(src_cell, dst_cell):
    """复制单元格样式"""
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.alignment = copy(src_cell.alignment)
        dst_cell.border = copy(src_cell.border)
        dst_cell.number_format = copy(src_cell.number_format)


def copy_worksheet_format(src_ws, dst_ws):
    """复制工作表的格式（列宽、行高、合并单元格等）"""
    # 复制列宽
    for col_letter, col_dim in src_ws.column_dimensions.items():
        if col_dim.width:
            dst_ws.column_dimensions[col_letter].width = col_dim.width
    
    # 复制行高
    for row_num, row_dim in src_ws.row_dimensions.items():
        if row_dim.height:
            dst_ws.row_dimensions[row_num].height = row_dim.height
    
    # 不复制合并单元格，新数据会重新计算


def generate_menu_rows(menu_data):
    """
    根据menu.json生成菜单行数据
    返回列表，每个元素是 (一级菜单, 二级菜单, 三级菜单, 地址)
    """
    rows = []
    
    for first_level in menu_data:
        first_name = first_level.get('Name', '')
        second_menus = first_level.get('SubMenus', [])
        
        if not second_menus:
            # 没有二级菜单
            rows.append((first_name, None, None, None))
            continue
        
        for second_level in second_menus:
            second_name = second_level.get('Name', '')
            third_menus = second_level.get('SubMenus', [])
            
            if not third_menus:
                # 没有三级菜单
                rows.append((first_name, second_name, None, None))
            else:
                # 有三级菜单
                for third_level in third_menus:
                    third_name = third_level.get('Name', '')
                    url = third_level.get('NavigateUrl', '')
                    # 处理空字符串url
                    if url == '':
                        url = None
                    rows.append((first_name, second_name, third_name, url))
    
    return rows


def optimize_rows(rows):
    """
    优化行数据：
    - 相同的一级菜单，只在第一次出现时保留，后续设为None
    - 相同的二级菜单（在同一个一级菜单下），只在第一次出现时保留，后续设为None
    """
    optimized = []
    prev_first = None
    prev_second = None
    
    for row in rows:
        first, second, third, url = row
        
        # 处理一级菜单
        if first == prev_first:
            first = None
        else:
            prev_first = first
            prev_second = None  # 新一级菜单时重置二级菜单记忆
        
        # 处理二级菜单
        if second == prev_second and first is None:
            second = None
        else:
            prev_second = second
        
        optimized.append((first, second, third, url))
    
    return optimized


def create_excel(menu_data, template_path, output_path):
    """创建Excel文件"""
    # 加载模板
    template_wb = openpyxl.load_workbook(template_path)
    template_ws = template_wb.active
    
    # 创建新的工作簿
    wb = Workbook()
    ws = wb.active
    ws.title = template_ws.title
    
    # 复制工作表格式
    copy_worksheet_format(template_ws, ws)
    
    # 复制说明行和表头行
    from openpyxl.cell.cell import MergedCell
    for row_idx in range(1, 3):
        for col_idx in range(1, template_ws.max_column + 1):
            src_cell = template_ws.cell(row=row_idx, column=col_idx)
            dst_cell = ws.cell(row=row_idx, column=col_idx)
            if not isinstance(dst_cell, MergedCell):
                dst_cell.value = src_cell.value
                copy_cell_style(src_cell, dst_cell)
    
    # 生成菜单数据
    rows = generate_menu_rows(menu_data)
    rows = optimize_rows(rows)
    
    # 写入数据
    for idx, row in enumerate(rows, start=3):
        first, second, third, url = row
        ws.cell(row=idx, column=1, value=first)
        ws.cell(row=idx, column=2, value=second)
        ws.cell(row=idx, column=3, value=third)
        ws.cell(row=idx, column=4, value=url)
    
    # 如果有Sheet2，复制它
    if 'Sheet2' in template_wb.sheetnames:
        template_ws2 = template_wb['Sheet2']
        ws2 = wb.create_sheet(title='Sheet2')
        copy_worksheet_format(template_ws2, ws2)
        for row_idx in range(1, template_ws2.max_row + 1):
            for col_idx in range(1, template_ws2.max_column + 1):
                src_cell = template_ws2.cell(row=row_idx, column=col_idx)
                dst_cell = ws2.cell(row=row_idx, column=col_idx)
                if not isinstance(dst_cell, MergedCell):
                    dst_cell.value = src_cell.value
                    copy_cell_style(src_cell, dst_cell)
    
    # 保存文件
    wb.save(output_path)
    print(f"Excel文件已生成: {output_path}")
    print(f"总共 {len(rows)} 行菜单数据")


def main():
    menu_data = load_menu_data('menu.json')
    create_excel(menu_data, '模板.xlsx', '菜单导出.xlsx')


if __name__ == '__main__':
    main()
