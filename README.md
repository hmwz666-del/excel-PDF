# Excel 转 PDF 批量转换工具

> 一键将文件夹中的 Excel 文件批量转换为 PDF，支持多进程并发、自动异常跳过、空白页过滤。

## 系统要求

- Windows 10 / 11
- Microsoft Office (Excel) 已安装
- Python 3.8+（开发/调试时需要，使用 EXE 则不需要）

## 快速开始

### 方式一：直接运行（需 Python 环境）

```bash
# 1. 安装依赖
pip install pywin32

# 2. 运行
python main.py
```

### 方式二：打包为 EXE（推荐给业务人员）

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 打包（二选一）
pyinstaller build.spec
# 或
pyinstaller -F -w --name "Excel转PDF工具" main.py

# 3. 生成的 EXE 在 dist/ 目录下
```

## 使用说明

1. 双击运行"Excel转PDF工具.exe"
2. 点击"浏览..."选择 Excel 文件所在的文件夹
3. 设置输出目录（默认为输入目录下的 `PDF_Output`）
4. 选择并发进程数（推荐 4）
5. 点击"🚀 开始转换"
6. 等待转换完成，查看统计结果

## 项目结构

```
EXCEL转PDF/
├── main.py          # 程序入口
├── config.py        # 配置常量
├── converter.py     # 核心转换引擎 (win32com)
├── worker.py        # 多进程管理
├── gui.py           # GUI 界面 (tkinter)
├── build.spec       # PyInstaller 打包配置
├── requirements.txt # 依赖清单
└── docs/            # 项目文档
```

## 性能参考

| 文件数量 | 并发进程 | 预计耗时 |
|----------|---------|---------|
| 100 个   | 4 进程  | ~2 分钟 |
| 1000 个  | 4 进程  | ~15-20 分钟 |
| 3000 个  | 4 进程  | ~45-60 分钟 |

## 注意事项

- 转换过程中会在后台启动 Excel 实例，请勿手动操作 Excel
- 密码保护和损坏的文件会自动跳过，详见日志文件
- 日志文件自动保存在 EXE 所在目录
