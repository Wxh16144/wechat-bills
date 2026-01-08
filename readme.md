# 微信账单 PDF 转 Excel 工具

这是一个将微信账单 PDF 文件批量转换为 Excel (.xlsx) 格式的小工具。  
本分支修改为命令行版本（CLI），去除了 GUI 依赖，方便在控制台环境下使用。

## 源代码

原项目地址：[https://gitee.com/tangx_666/wechat-bills](https://gitee.com/tangx_666/wechat-bills)

## 安装步骤 (macOS/Linux)

在 macOS 上，推荐使用 Python 虚拟环境来避免依赖冲突。

1.  **打开终端并运行以下命令**：

    ```bash
    # 进入项目目录
    cd /path/to/wechat-bills

    # 创建虚拟环境 (仅需一次)
    python3 -m venv venv

    # 激活虚拟环境 (每次新开终端都需要执行)
    source venv/bin/activate
    ```

    *激活成功后，你的终端提示符前会出现 `(venv)` 字样。*

2.  **安装依赖**：

    ```bash
    pip install -r requirements.txt
    ```

## 使用方法

脚本支持直接指定输入文件或目录。

### 1. 转换单个文件

```bash
python WeChatBillsPdf2Xlsx.py ./账单2025.pdf
```

这将在同目录下生成 Excel 文件。

### 2. 转换整个目录

```bash
python WeChatBillsPdf2Xlsx.py ./所有账单/
```

这将会扫描该目录下所有的 `.pdf` 文件并进行转换。

### 3. 指定输出路径 (可选)

你可以显式指定输出文件的位置：

```bash
python WeChatBillsPdf2Xlsx.py ./my_bill.pdf ./output/2026_bill.xlsx
```

## 注意事项

*   PDF 文件必须是微信导出的原生电子账单。
*   涉及个人隐私，请妥善保管转换后的文件。
*   脚本完全在本地运行，不会上传任何数据。
