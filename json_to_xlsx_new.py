def create_schedule_sheet_from_assignments_format(ws, data):
    """
    处理assignments格式 - 完全泛化的实现
    严格按照columns定义创建列结构，智能判断工作日/周末
    """
    font_family = '新宋体'
    sheet_structure = data.get('sheet_structure', {})
    assignments = data.get('assignments', [])

    title = sheet_structure.get('title_row', '应用變更人員時間安排表')
    columns = sheet_structure.get('columns', [])
    time_row_data = sheet_structure.get('time_row_data', {})

    # 严格按照columns定义创建列结构
    dates = []
    date_col_mapping = {}
    current_col = 2

    # 按日期分组columns（保持原始顺序）
    date_columns = {}  # {date: [col_names]}
    for col in columns[1:]:  # 跳过第一列"执行人"
        if '_' in col:
            date_part = col.split('_')[0]
        else:
            date_part = col

        if date_part not in date_columns:
            date_columns[date_part] = []
        date_columns[date_part].append(col)

    # 按columns中第一次出现的顺序创建列
    seen_dates = []
    for col in columns[1:]:
        if '_' in col:
            date_part = col.split('_')[0]
        else:
            date_part = col
        if date_part not in seen_dates:
            seen_dates.append(date_part)

    for date in seen_dates:
        date_col_mapping[date] = []
        for col in date_columns[date]:
            date_col_mapping[date].append(current_col)
            ws.column_dimensions[get_column_letter(current_col)].width = 20
            current_col += 1
        if date not in dates:
            dates.append(date)

    ws.column_dimensions['A'].width = 33
    total_cols = current_col - 1

    # 标题行
    ws.row_dimensions[1].height = 45
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = title
    title_cell.font = Font(name=font_family, size=30, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    title_fill_color = fix_color_code('#E6F3FF')
    title_cell.fill = PatternFill(start_color=title_fill_color, end_color=title_fill_color, fill_type='solid')
    if total_cols > 1:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)

    # 日期行
    ws.row_dimensions[2].height = 35
    date_fill_color = fix_color_code('#E6F3FF')
    for date in dates:
        cols = date_col_mapping.get(date, [])
        if not cols:
            continue
        start_c = cols[0]
        end_c = cols[-1]
        cell = ws.cell(row=2, column=start_c)
        cell.value = date
        cell.font = Font(name=font_family, size=26, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.fill = PatternFill(start_color=date_fill_color, end_color=date_fill_color, fill_type='solid')
        if start_c != end_c:
            ws.merge_cells(start_row=2, start_column=start_c, end_row=2, end_column=end_c)

    # 时间行（智能判断工作日/周末）
    ws.row_dimensions[3].height = 30
    for date in dates:
        cols = date_col_mapping.get(date, [])
        time_slots = time_row_data.get(date, [])

        # 判断是否为工作日：所有时间槽相同且不包含'-'
        is_workday = (len(time_slots) > 0 and
                     all(slot == time_slots[0] for slot in time_slots) and
                     '-' not in str(time_slots[0]))

        if is_workday:
            # 工作日：查找最长的时间范围
            best_start = None
            best_end = None
            best_duration = 0

            for assignment in assignments:
                for task in assignment.get('tasks', []):
                    if task.get('date') == date:
                        start = task.get('original_start', '')
                        end = task.get('original_end', '')
                        if start and end:
                            try:
                                start_min = int(start.split(':')[0]) * 60 + int(start.split(':')[1])
                                end_min = int(end.split(':')[0]) * 60 + int(end.split(':')[1])
                                duration = end_min - start_min
                                if duration > best_duration:
                                    best_duration = duration
                                    best_start = start
                                    best_end = end
                            except:
                                pass

            time_range = f"{best_start}-{best_end}" if best_start and best_end else time_slots[0]

            # 只在第一列显示时间范围
            for i, col in enumerate(cols):
                cell = ws.cell(row=3, column=col)
                if i == 0:
                    cell.value = time_range
                else:
                    cell.value = ''  # 其他列为空
                cell.font = Font(name=font_family, size=14)
                cell.alignment = Alignment(horizontal='center', vertical='center')
        else:
            # 周末：按列显示时间点
            for i, time_slot in enumerate(time_slots):
                if i >= len(cols):
                    break
                cell = ws.cell(row=3, column=cols[i])
                cell.value = time_slot
                cell.font = Font(name=font_family, size=14)
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # 人员行（正确分配工单）
    for row_idx, assignment in enumerate(assignments, start=4):
        person_name = assignment.get('personnel', '')
        color = assignment.get('color', '#FFFFFF')
        tasks = assignment.get('tasks', [])
        ws.row_dimensions[row_idx].height = 90
        bg_color = fix_color_code(color)

        # 姓名单元格
        person_cell = ws.cell(row_idx, column=1)
        person_cell.value = person_name
        person_cell.font = Font(name=font_family, size=14, bold=True)
        person_cell.alignment = Alignment(horizontal='center', vertical='center')
        person_cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')

        # 任务单元格
        for task in tasks:
            date = task.get('date', '')
            work_order = task.get('work_order', '')
            original_start = task.get('original_start', '')
            original_end = task.get('original_end', '')
            merge_columns = task.get('merge_columns', [])

            if date not in date_col_mapping:
                continue

            cols = date_col_mapping[date]
            time_slots = time_row_data.get(date, [])

            # 判断是否为工作日
            is_workday = (len(time_slots) > 0 and
                         all(slot == time_slots[0] for slot in time_slots) and
                         '-' not in str(time_slots[0]))

            if is_workday:
                # 工作日：使用第一列
                target_col = cols[0]
                cell = ws.cell(row_idx, column=target_col)
                if cell.value:  # 如果已经有内容，添加换行
                    cell.value = str(cell.value) + '\n' + work_order
                else:
                    cell.value = work_order
                cell.font = Font(name=font_family, size=12)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
            else:
                # 周末：使用merge_columns或匹配时间点
                target_cols = []

                if merge_columns:
                    # 使用merge_columns确定目标列
                    for merge_col in merge_columns:
                        if '_' in merge_col:
                            date_part = merge_col.split('_')[0]
                            slot_num = int(merge_col.split('_')[1]) - 1
                            if date_part in date_col_mapping:
                                cols_for_date = date_col_mapping[date_part]
                                if slot_num < len(cols_for_date):
                                    target_cols.append(cols_for_date[slot_num])
                else:
                    # 通过时间匹配找到目标列
                    start_time = task.get('original_start', '')
                    end_time = task.get('original_end', '')

                    if start_time and end_time:
                        # 转换时间为分钟
                        def time_to_minutes(time_str):
                            try:
                                # Handle dashed time ranges like '18:00-19:00' by extracting start time
                                if '-' in str(time_str):
                                    time_str = time_str.split('-')[0]
                                h, m = map(int, time_str.split(':'))
                                return h * 60 + m
                            except:
                                return -1

                        start_minutes = time_to_minutes(start_time)
                        end_minutes = time_to_minutes(end_time)

                        # 检查跨越哪些时间点
                        for i, slot in enumerate(time_slots):
                            slot_minutes = time_to_minutes(slot)
                            if start_minutes <= slot_minutes <= end_minutes:
                                if i < len(cols):
                                    target_cols.append(cols[i])

                # 合并单元格并放置工单
                if target_cols:
                    target_cols = sorted(list(set(target_cols)))
                    if len(target_cols) > 1:
                        min_col = min(target_cols)
                        max_col = max(target_cols)
                        ws.merge_cells(start_row=row_idx, start_column=min_col, end_row=row_idx, end_column=max_col)
                        cell = ws.cell(row_idx, column=min_col)
                    else:
                        cell = ws.cell(row_idx, column=target_cols[0])

                    if work_order:
                        cell.value = work_order
                        cell.font = Font(name=font_family, size=12)
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')

    # 边框
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    for row in ws.iter_rows(min_row=1, max_row=len(assignments)+3, min_col=1, max_col=total_cols):
        for cell in row:
            cell.border = thin_border
