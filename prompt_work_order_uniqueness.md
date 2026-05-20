# 🚨 工单全局唯一性致命规则 🚨

## 最高优先级规则（违反即视为致命错误）

**AI必须逐字逐句记住并执行以下规则：**

### 1. 工单全局唯一性铁律

**每张工单在全局排班表中只能出现一次，绝对禁止重复分配！**

**🚨🚨🚨 致命警告（最高优先级）🚨🚨🚨**
**每张工单全局只能出现一次，绝对禁止重复分配！**
**如果发现重复分配，立即删除所有重复项，只保留第一次分配！**

### 2. 致命错误示例（绝对禁止）

#### ❌ 错误示例：工单重复出现
```
工单：2367 GBS 葉俊濠
错误情况：
- 第一次出现在：05-21（周四，工作日）
- 第二次出现在：05-24（周六，周末）
❌❌❌ 致命错误！同一工单出现了两次！

正确做法：
- 只保留第一次分配（05-21）
- 删除第二次分配（05-24）
- 确保工单2367在全局排班表中只出现一次
```

#### ❌ 错误示例：工单被分配给多个人
```
工单：2381 DWS 王文斌
错误情况：
- 分配给人员A：岑惠韜
- 分配给人员B：羅家勝
❌❌❌ 致命错误！同一工单分配给了多个人！

正确做法：
- 只分配给一个人（比如岑惠韜）
- 确保工单2381在全局排班表中只出现一次
```

### 3. 强制检查逻辑

**AI在生成排班表后必须执行以下检查：**

```python
# 伪代码：AI必须执行的工单唯一性检查
def check_work_order_uniqueness(personnel_assignments):
    # 第一步：收集所有工单
    all_work_orders = {}  # {工单号: [出现次数, 出现位置]}

    for person in personnel_assignments:
        for assignment in person['assignments']:
            work_order = assignment['work_order']
            work_order_id = work_order.split()[0]  # 提取工单号（如"2367"）

            if work_order_id in all_work_orders:
                # 发现重复！
                all_work_orders[work_order_id][0] += 1
                all_work_orders[work_order_id][1].append({
                    'person': person['name'],
                    'date': assignment['date'],
                    'time': f"{assignment['original_start']}-{assignment['original_end']}"
                })
            else:
                all_work_orders[work_order_id] = [1, [{
                    'person': person['name'],
                    'date': assignment['date'],
                    'time': f"{assignment['original_start']}-{assignment['original_end']}"
                }]]

    # 第二步：检查是否有重复
    duplicates = {k: v for k, v in all_work_orders.items() if v[0] > 1}

    if duplicates:
        # 发现重复工单！
        for work_order_id, info in duplicates.items():
            print(f"❌ 致命错误：工单{work_order_id}出现了{info[0]}次！")
            print(f"出现位置：")
            for i, occurrence in enumerate(info[1]):
                print(f"  {i+1}. {occurrence['person']} - {occurrence['date']} ({occurrence['time']})")

            # 立即修复：只保留第一次出现
            print(f"修复：只保留第一次分配，删除其他{info[0]-1}次重复分配")

        return False  # 发现重复，需要修复

    return True  # 通过检查，无重复
```

### 4. 生成前强制检查清单

**AI在输出JSON前必须询问自己：**

- [ ] 是否每个工单编号在全局排班表中只出现一次？
- [ ] 是否有任何工单被分配给了多个不同的人员？
- [ ] 是否有任何工单在多个不同的日期出现？
- [ ] 如发现重复工单，是否立即删除所有重复项？
- [ ] 最终结果中是否每个工单编号出现次数=1？

### 5. 常见错误模式（必须避免）

#### ❌ 错误模式1：因为人员不足就重复分配工单
```
错误思维：工单太多，人员不足，所以把同一工单分配给多个人
正确做法：绝对禁止！宁可一个人承担多个工单，也不能重复分配同一工单
```

#### ❌ 错误模式2：因为时间冲突就重复安排时间
```
错误思维：工单时间冲突，所以在多个日期都安排同一工单
正确做法：绝对禁止！同一工单只能在全局出现一次，时间冲突需要另想办法
```

#### ❌ 错误模式3：不小心忽略了已有分配
```
错误思维：我忘记了已经分配过这个工单，所以又分配了一次
正确做法：绝对禁止！必须记录已分配的工单，确保不重复分配
```

### 6. 修复重复工单的标准流程

**如果AI发现自己重复分配了工单，必须按以下流程修复：**

1. **识别重复**：找出所有出现次数>1的工单
2. **保留第一次**：只保留该工单的第一次出现
3. **删除重复项**：删除所有后续的重复分配
4. **验证修复**：确保最终结果中每个工单只出现一次

**示例修复流程：**
```
发现问题：工单2367出现了两次
  1. 第一次：李曼菁 - 05-21 (18:00-18:30)
  2. 第二次：羅家勝 - 05-24 (18:00-18:30)  ❌ 删除这个

修复后：工单2367只出现一次
  1. 第一次：李曼菁 - 05-21 (18:00-18:30)  ✅ 保留这个
```

### 7. 零容忍政策

**工单唯一性规则执行零容忍政策：**
- 一次重复 = 整个排班表作废
- 不允许"不小心"或"忘记了"
- 不允许"人员不足"的借口
- 不允许"时间冲突"的理由
- 必须百分之百遵守工单唯一性规则

### 8. 记忆强化

**AI必须将以下规则刻入核心逻辑：**

```
FOR each work_order IN all_work_orders:
    count = 0
    FOR each assignment IN personnel_assignments:
        IF assignment.work_order_id == work_order:
            count += 1
            IF count > 1:
                DELETE assignment  # 立即删除重复项

    IF count == 0:
        FATAL_ERROR(f"工单{work_order}未分配！")
    ELIF count > 1:
        FATAL_ERROR(f"工单{work_order}重复分配了{count}次！")
```

### 9. 与其他规则的关系

**工单唯一性规则优先级最高：**
- 当工单唯一性规则与其他规则冲突时，工单唯一性规则优先
- 例如：宁可违反均衡分配规则，也不能重复分配同一工单
- 例如：宁可违反工作日/周末分离规则，也不能重复分配同一工单

### 10. 最终校验

**AI在输出最终JSON前必须执行的最后检查：**

```python
# 最终校验逻辑
def final_validation(personnel_assignments):
    # 收集所有工单编号
    work_order_ids = []
    for person in personnel_assignments:
        for assignment in person['assignments']:
            work_order = assignment['work_order']
            work_order_id = work_order.split()[0]
            work_order_ids.append(work_order_id)

    # 检查是否有重复
    unique_ids = set(work_order_ids)
    if len(work_order_ids) != len(unique_ids):
        duplicates = [id for id in unique_ids if work_order_ids.count(id) > 1]
        raise FatalError(f"发现重复工单：{duplicates}，请立即修复！")

    print("✅ 通过最终校验：所有工单都是唯一的")
    return True
```

---

**最后警告：工单全局唯一性是排班的基本原则，违反此规则将被视为最严重的致命错误，必须重新生成整个排班表！**

**每张工单全局只能出现一次，绝对禁止重复分配！**
