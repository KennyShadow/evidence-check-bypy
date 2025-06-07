# 收入证据管理程序

<div align="center">

![Python](https://img.shields.io/badge/python-v3.7+-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![GUI](https://img.shields.io/badge/GUI-CustomTkinter-orange.svg)
![Data](https://img.shields.io/badge/data-Pandas-red.svg)
![Excel](https://img.shields.io/badge/excel-OpenPyXL-green.svg)
![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)

**企业级收入证据管理系统** - 专业的财务数据管理与证据归档解决方案

[🎯 项目简介](#-项目简介) •
[✨ 功能特性](#-功能特性) •
[🚀 快速开始](#-快速开始) •
[📖 用户手册](#-用户手册) •
[🛠 开发文档](#-开发文档) •
[🔧 部署指南](#-部署指南) •
[❓ 常见问题](#-常见问题)

</div>

---

## 📋 项目简介

收入证据管理程序是一个专业的财务数据管理工具，旨在帮助用户高效地管理收入数据与相关证据文件。该程序提供了直观的图形界面，支持Excel数据导入、多维度数据筛选、附件管理、差异分析等功能，大大提升了财务数据处理的效率和准确性。

### 🎯 设计理念

- **数据安全**：本地存储，保护用户隐私和数据安全
- **易于使用**：现代化UI设计，操作简单直观
- **功能完整**：覆盖数据导入、处理、分析、导出的完整流程
- **性能优秀**：高效的数据处理和响应速度

## ✨ 功能特性

### 📊 数据管理
- **Excel数据导入**
  - 支持 `.xlsx` 和 `.xls` 格式
  - 智能列名映射和数据验证
  - 多工作表选择和预览
  - 批量数据处理

- **多项目管理**
  - 支持创建和管理多个独立项目
  - 项目间数据完全隔离
  - 项目删除和备份功能

- **数据编辑**
  - 灵活的记录增删改功能
  - 实时数据验证和保存
  - 批量操作支持

### 🔍 高级筛选
- **多维度筛选**
  - 按合同号、客户名称筛选
  - 按差异状态、附件状态筛选
  - 自定义日期范围筛选
  - 多条件组合筛选

- **智能搜索**
  - 实时关键词搜索
  - 模糊匹配算法
  - 高亮显示搜索结果

### 📁 附件管理
- **灵活的附件关联**
  - 支持拖拽上传附件
  - 多种文件格式支持（PDF、图片、Office文档等）
  - 一对多附件关联
  - 自动文件夹组织

- **附件操作**
  - 快速打开附件文件夹
  - 附件重命名和删除
  - 附件数量统计显示

### 📈 数据分析
- **差异计算**
  - 自动计算收入差异
  - 可视化差异状态显示
  - 差异原因备注功能

- **统计分析**
  - 实时数据统计
  - 多维度数据分析
  - 图表可视化展示

- **版本管理**
  - 多版本数据对比
  - 数据变更追踪
  - 新增记录自动标识

### 💾 数据导出
- **灵活导出**
  - Excel格式导出
  - 自定义导出字段
  - 筛选结果导出
  - 保持数据格式

- **数据备份**
  - 自动数据备份
  - 手动备份创建
  - 备份文件管理

## 🚀 快速开始

### 系统要求

- **操作系统**：Windows 10/11
- **Python版本**：3.7 及以上
- **内存**：建议 4GB 及以上
- **存储空间**：至少 100MB 可用空间

### 安装步骤

1. **克隆仓库**
   ```bash
   git clone https://github.com/your-username/income-evidence-manager.git
   cd income-evidence-manager
   ```

2. **创建虚拟环境**（推荐）
   ```bash
   python -m venv venv
   venv\Scripts\activate  # Windows
   ```

3. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

4. **运行程序**
   ```bash
   python main.py
   ```

### 首次使用

1. 程序启动后会显示项目选择器
2. 点击"创建新项目"，输入项目名称和描述
3. 进入主界面，开始使用各项功能

## 📖 使用指南

### 项目管理

#### 创建新项目
- 在项目启动器中点击"新建项目"
- 输入项目名称（必填）和描述（可选）
- 系统会自动创建项目目录和配置文件

#### 切换项目
- 在项目启动器中选择已有项目
- 双击项目名称或点击"打开项目"

### 数据导入

#### Excel文件导入
1. 准备Excel文件，确保包含以下必要列：
   - 合同号（必须唯一）
   - 客户名称
   - 本年确认的收入

2. 点击"导入Excel"按钮
3. 选择Excel文件和目标工作表
4. 确认列映射关系
5. 点击"导入"完成数据导入

#### 数据验证规则
- 合同号不能为空且必须唯一
- 客户名称不能为空
- 收入金额必须为有效数字
- 重复合同号会提示处理方式

### 数据操作

#### 添加记录
- 点击"添加记录"按钮
- 填写必要信息并保存
- 新增记录会自动标记

#### 编辑记录
- 双击表格行或点击"编辑"按钮
- 修改信息后保存
- 支持批量编辑操作

#### 删除记录
- 选择要删除的记录
- 点击"删除"按钮
- 确认删除操作

### 附件管理

#### 上传附件
1. 选择目标记录
2. 点击"管理附件"按钮
3. 使用以下方式上传：
   - 点击"添加附件"按钮选择文件
   - 直接拖拽文件到附件列表

#### 附件操作
- **查看附件**：双击附件名称
- **重命名**：右键选择重命名
- **删除附件**：选择附件后点击删除
- **打开文件夹**：快速访问附件存储位置

### 数据筛选

#### 基础筛选
- 使用顶部筛选器设置条件
- 支持合同号、客户名称模糊搜索
- 按差异状态、附件状态筛选

#### 高级筛选
- 组合多个筛选条件
- 自定义日期范围
- 保存和恢复筛选设置

### 数据分析

#### 差异分析
- 输入"附件确认的收入"
- 系统自动计算差异金额
- 添加差异原因备注

#### 统计信息
- 左侧面板显示实时统计
- 总记录数、差异统计
- 附件关联状态统计

### 数据导出

#### 导出Excel
1. 设置筛选条件（可选）
2. 点击"导出Excel"按钮
3. 选择导出路径和文件名
4. 确认导出设置

## 🛠 开发指南

### 项目结构

```
收入证据管理程序/
├── main.py                 # 程序入口
├── requirements.txt        # 依赖列表
├── src/                   # 源代码
│   ├── config.py          # 配置文件
│   ├── gui/               # GUI界面模块
│   │   ├── main_window.py         # 主窗口
│   │   ├── project_launcher.py    # 项目启动器
│   │   ├── record_dialog.py       # 记录编辑对话框
│   │   ├── attachment_dialog.py   # 附件管理对话框
│   │   └── ...                    # 其他GUI组件
│   ├── data/              # 数据处理模块
│   │   ├── data_processor.py      # 数据处理器
│   │   ├── excel_handler.py       # Excel处理
│   │   ├── file_manager.py        # 文件管理
│   │   └── project_manager.py     # 项目管理
│   └── models/            # 数据模型
│       ├── database.py            # 数据库模型
│       ├── income_record.py       # 收入记录模型
│       └── attachment.py          # 附件模型
└── data/                  # 数据目录（被gitignore忽略）
    ├── app.log           # 应用日志
    ├── database.pkl      # 数据库文件
    └── attachments/      # 附件存储
```

### 技术栈

- **GUI框架**：[CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) - 现代化的Tkinter界面库
- **数据处理**：[Pandas](https://pandas.pydata.org/) - 数据分析和处理
- **Excel操作**：[OpenPyXL](https://openpyxl.readthedocs.io/) - Excel文件读写
- **图像处理**：[Pillow](https://pillow.readthedocs.io/) - 图像处理库
- **数据持久化**：Python Pickle - 对象序列化

### 核心模块说明

#### 数据模型 (models/)
- `database.py`: 数据库操作和管理
- `income_record.py`: 收入记录数据结构
- `attachment.py`: 附件信息数据结构

#### 数据处理 (data/)
- `data_processor.py`: 核心数据处理逻辑
- `excel_handler.py`: Excel文件导入导出
- `file_manager.py`: 文件和附件管理
- `project_manager.py`: 项目管理功能

#### 用户界面 (gui/)
- `main_window.py`: 主界面和核心交互
- `project_launcher.py`: 项目选择和创建
- `record_dialog.py`: 记录编辑界面
- `attachment_dialog.py`: 附件管理界面

### 开发环境配置

1. **安装开发依赖**
   ```bash
   pip install -r requirements.txt
   pip install pytest  # 测试框架（可选）
   ```

2. **代码规范**
   - 遵循 PEP 8 Python编码规范
   - 使用类型注解提高代码可读性
   - 模块化设计，单一职责原则

3. **调试模式**
   - 修改 `src/config.py` 中的日志级别为 `DEBUG`
   - 查看 `data/app.log` 获取详细日志信息

### 扩展开发

#### 添加新功能
1. 在相应模块中添加功能函数
2. 更新数据模型（如需要）
3. 创建或修改GUI界面
4. 更新配置文件
5. 添加相应的错误处理

#### 自定义配置
- 修改 `src/config.py` 中的配置项
- 重启程序使配置生效

## ❓ 常见问题

### 安装问题

**Q: 安装依赖时出现错误怎么办？**
A: 
- 确保Python版本为3.7+
- 尝试升级pip：`python -m pip install --upgrade pip`
- 使用清华镜像源：`pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple`

**Q: 程序无法启动？**
A:
- 检查是否正确安装所有依赖
- 查看控制台错误信息
- 检查 `data/app.log` 日志文件

### 使用问题

**Q: Excel导入失败？**
A:
- 确保Excel文件未被其他程序占用
- 检查文件格式是否为.xlsx或.xls
- 确保文件中包含必要的列（合同号、客户名、收入）

**Q: 附件上传失败？**
A:
- 检查文件是否存在且未被占用
- 确保有足够的磁盘空间
- 检查文件路径中是否包含特殊字符

**Q: 数据丢失怎么办？**
A:
- 检查 `data/backups/` 目录中的备份文件
- 程序会自动创建数据备份
- 避免直接删除 `data/` 目录

### 性能问题

**Q: 程序运行缓慢？**
A:
- 使用筛选功能减少显示数据量
- 定期清理过期的备份文件
- 关闭不必要的其他程序释放内存

**Q: 大文件导入慢？**
A:
- 将大文件分割成多个小文件导入
- 关闭其他占用内存的程序
- 确保足够的磁盘空间

## 📄 许可证

本项目采用 MIT 许可证。详情请参阅 [LICENSE](LICENSE) 文件。

## 🤝 贡献

欢迎提交Issue和Pull Request来改进这个项目！

### 贡献方式
1. Fork 这个仓库
2. 创建你的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交你的修改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开一个 Pull Request

## 📞 支持

如果你遇到问题或有建议，请：

1. 查看本文档的常见问题部分
2. 搜索已有的 [Issues](https://github.com/your-username/income-evidence-manager/issues)
3. 创建新的 Issue 详细描述问题

---

<div align="center">

**感谢使用收入证据管理程序！**

如果这个项目对你有帮助，请考虑给它一个 ⭐

</div> 