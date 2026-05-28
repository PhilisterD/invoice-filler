# 早鸟天筹明细生成器

将 Excel 明细表内容自动填充到 Word 模板中，按"受托方"分组合并生成文档。

## 功能

- 按"受托方"分组，同一受托方的记录合并到一个 Word 文档
- 自动替换委托方名称、日期
- 动态插入表格数据行（支持同一受托方多行记录）
- 自动计算总计金额并生成中文大写
- 支持单个下载或打包下载全部

## 项目结构

```
invoice-filler/
├── app.py              # Flask Web 应用入口
├── filler.py           # Excel -> Word 核心逻辑
├── utils.py            # 工具函数（大写金额、日期转换）
├── build_windows.py    # Windows 打包脚本
├── requirements.txt    # Python 依赖
├── templates/
│   ├── index.html      # 上传页面
│   └── result.html     # 下载页面
├── static/
├── uploads/            # 上传文件临时目录
└── output/             # 生成的 Word 文档
```

## 开发环境运行

```bash
pip install -r requirements.txt
python app.py
```

自动打开浏览器访问 http://127.0.0.1:5000/

## Windows 打包（生成 exe）

在 Windows 电脑上：

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 打包
python build_windows.py

# 3. 输出文件
dist/早鸟天筹明细生成器.exe
```

双击 exe 即可运行，无需安装 Python。

## 使用说明

1. 准备 Excel 明细表，需包含以下列：
   - 序号、姓名、受托方、服务内容、数量/个、单价/元、合计/元、总计/元、日期
2. 准备 Word 模板，标黄（高亮）的部分会被自动替换
3. 双击 exe，在浏览器中上传两个文件
4. 点击"生成文档"，下载结果

## Excel 列名说明

程序会自动识别包含以下关键字的列：
- "受托方" -> 分组依据
- "服务内容" -> 表格第一列
- "数量" -> 表格第二列
- "单价" -> 表格第三列
- "合计" -> 表格第四列
- "总计" -> 表格第五列
- "日期" -> 替换文档中的日期
