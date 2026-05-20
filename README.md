# 计划生成工具 (Schedule Generator)

自动生成人员排班表的智能工具，支持工作日和周末的时间安排。

## 功能特点

- ✅ **智能排班**: 自动分析工单并合理分配给人员
- ✅ **周末支持**: 正确处理周末的时间槽和工单合并
- ✅ **Excel导出**: 生成格式化的Excel排班表
- ✅ **数据验证**: 自动检查工单重复和分配均衡性
- ✅ **批量处理**: 支持批量处理多个工单

## 最新修复

### v2.0 - 2026-05-20

**修复**: 周末时间槽单元格合并错误

- **问题**: 当时间槽包含时间范围（如 "18:00-19:00"）时，导致列排序错误
- **影响**: 周末工单的单元格合并到错误的列
- **修复**: 改进时间解析函数，正确处理时间范围格式
- **结果**: 单元格合并现在正确映射到对应的时间槽

## 快速开始

### Windows 用户

1. 下载最新的 `schedule-v2.exe` 文件
2. 双击运行 `run_schedule.bat`
3. 查看生成的 `output_transformed.xlsx` 文件

### Mac/Linux 用户

1. 确保已安装 Python 3.8+
2. 安装依赖: `pip install -r requirements.txt`
3. 运行: `python call_ai.py`

## 文件说明

### 核心文件

- `call_ai.py` - 主程序入口
- `json_to_xlsx.py` - JSON转Excel转换器（已修复）
- `config.ini` - 配置文件
- `run_schedule.bat` - Windows批处理脚本

### 输出文件

- `output_transformed.xlsx` - 生成的排班表
- `ai_responses/schedule_result.json` - AI响应数据

### 配置文件

- `data/变更明細數據報表_20250518.xlsx` - 数据源

## 系统要求

- **Python**: 3.8 或更高版本
- **操作系统**: Windows 10+, macOS 10.14+, Linux
- **依赖包**: openpyxl, pandas, requests

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用说明

### 1. 准备数据

将你的工单数据放在 `data/` 目录下的Excel文件中。

### 2. 配置参数

编辑 `config.ini` 文件，设置API密钥和其他参数。

### 3. 运行程序

**Windows:**
```bash
run_schedule.bat
```

**Mac/Linux:**
```bash
python call_ai.py
```

### 4. 查看结果

打开生成的 `output_transformed.xlsx` 文件查看排班表。

## 故障排除

### 问题: 时间槽合并错误

如果发现周末工单的单元格合并到错误的列，请确保你使用的是 v2.0 或更高版本。

### 问题: 编码错误

**Windows**: 确保批处理文件以 UTF-8 编码保存
**Mac/Linux**: 设置环境变量 `export LANG=zh_CN.UTF-8`

### 问题: 依赖包安装失败

尝试使用国内镜像源:
```bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

## 开发说明

### 构建可执行文件

```bash
# 安装 PyInstaller
pip install pyinstaller

# 构建 exe
pyinstaller --onefile --console --name schedule-v2 call_ai.py

# Mac/Linux 构建
pyinstaller --onefile --console call_ai.py
```

### 代码结构

```
schedule-v2/
├── call_ai.py              # 主程序
├── json_to_xlsx.py         # Excel生成器
├── config.ini              # 配置文件
├── requirements.txt        # Python依赖
├── run_schedule.bat        # Windows启动脚本
├── data/                   # 数据目录
├── ai_responses/          # AI响应存储
└── output_transformed.xlsx # 输出文件
```

## 更新日志

### v2.0 (2026-05-20)
- 🔧 修复周末时间槽单元格合并错误
- ✨ 改进时间范围解析逻辑
- 📝 添加详细的使用文档

### v1.0 (2026-05-18)
- 🎉 初始版本发布
- ✅ 基础排班功能
- 📊 Excel导出支持

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！

## 联系方式

如有问题，请通过 GitHub Issues 联系。
