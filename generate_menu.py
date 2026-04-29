#!/usr/bin/env python3
"""
根据menu.json生成Excel文件，参照模板.xlsx的格式
支持单元格合并
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
    """复制工作表的格式（列宽、行高等）"""
    # 复制列宽
    for col_letter, col_dim in src_ws.column_dimensions.items():
        if col_dim.width:
            dst_ws.column_dimensions[col_letter].width = col_dim.width
    
    # 复制行高
    for row_num, row_dim in src_ws.row_dimensions.items():
        if row_dim.height:
            dst_ws.row_dimensions[row_num].height = row_dim.height


def generate_menu_data(menu_data):
    """
    根据menu.json生成菜单数据
    返回:
        rows: 列表，每个元素是 (三级菜单, 地址)
        first_merges: 列表，每个元素是 (一级菜单名, 起始行, 结束行)
        second_merges: 列表，每个元素是 (二级菜单名, 起始行, 结束行)
    """
    rows = []  # [(三级菜单, 地址), ...]
    first_merges = []  # [(一级菜单名, 起始行, 结束行), ...]
    second_merges = []  # [(二级菜单名, 起始行, 结束行), ...]
    
    current_row = 3  # 从第3行开始（第1行说明，第2行表头）
    
    for first_level in menu_data:
        first_name = first_level.get('Name', '')
        second_menus = first_level.get('SubMenus', [])
        
        first_start = current_row
        
        if not second_menus:
            # 没有二级菜单，只有一级菜单
            rows.append((None, None))
            current_row += 1
        else:
            for second_level in second_menus:
                second_name = second_level.get('Name', '')
                third_menus = second_level.get('SubMenus', [])
                
                second_start = current_row
                
                if not third_menus:
                    # 没有三级菜单
                    rows.append((None, None))
                    current_row += 1
                else:
                    # 有三级菜单
                    for third_level in third_menus:
                        third_name = third_level.get('Name', '')
                        url = third_level.get('NavigateUrl', '')
                        # 处理空字符串url
                        if url == '':
                            url = None
                        rows.append((third_name, url))
                        current_row += 1
                
                second_end = current_row - 1
                if second_end >= second_start:
                    second_merges.append((second_name, second_start, second_end))
        
        first_end = current_row - 1
        if first_end >= first_start:
            first_merges.append((first_name, first_start, first_end))
    
    return rows, first_merges, second_merges


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
    
    # 复制说明行和表头行（处理合并单元格）
    for row_idx in range(1, 3):
        for col_idx in range(1, template_ws.max_column + 1):
            src_cell = template_ws.cell(row=row_idx, column=col_idx)
            dst_cell = ws.cell(row=row_idx, column=col_idx)
            if not isinstance(dst_cell, MergedCell):
                dst_cell.value = src_cell.value
                copy_cell_style(src_cell, dst_cell)
    
    # 复制说明行的合并单元格 A1:C1
    ws.merge_cells('A1:C1')
    
    # 生成菜单数据
    rows, first_merges, second_merges = generate_menu_data(menu_data)
    
    # 写入三级菜单和地址数据
    for idx, (third, url) in enumerate(rows, start=3):
        ws.cell(row=idx, column=3, value=third)
        ws.cell(row=idx, column=4, value=url)
    
    # 写入一级菜单并合并单元格
    for first_name, start_row, end_row in first_merges:
        ws.cell(row=start_row, column=1, value=first_name)
        if end_row > start_row:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
    
    # 写入二级菜单并合并单元格
    for second_name, start_row, end_row in second_merges:
        ws.cell(row=start_row, column=2, value=second_name)
        if end_row > start_row:
            ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
    
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
    print(f"一级菜单合并: {len(first_merges)} 处")
    print(f"二级菜单合并: {len(second_merges)} 处")


def main():
    menu_data = load_menu_data('menu.json')
    create_excel(menu_data, '模板.xlsx', '菜单导出.xlsx')


if __name__ == '__main__':
    main()
