# 微信账单 PDF 转 Excel 工具

这是一个将微信账单 PDF 文件批量转换为 Excel (.xlsx) 格式的小工具。本项目基于 [tangx_666/wechat-bills](https://gitee.com/tangx_666/wechat-bills) 二次开发，去除了原有的桌面 GUI 依赖，改为提供 **Web** 和 **CLI (命令行)** 两种版本，以满足不同场景的使用需求。

> 如需 **JSON, CSV, YAML** 等其他格式，请参考 [支持更多格式 (Issue #1)](https://github.com/Wxh16144/wechat-bills/issues/1)。

## 🌐 Web 版 (推荐)

如果您不想配置本地环境，可以直接使用 **Web 版**。
它完全基于浏览器运行，**所有数据处理均在您的设备本地完成**，不会上传任何账单文件，确保您的财务隐私安全。

**👉 立即访问：[https://wxh16144.github.io/wechat-bills/](https://wxh16144.github.io/wechat-bills/)**

<div align="center">
    <img src="https://i.imgur.com/fDfezvD.png" alt="Web 版截图"/>
</div>

## 🚀 CLI 版 (命令行)

专为开发者及需批量处理的用户设计，纯 Python 实现，无 GUI 依赖，轻量且便于脚本集成。

### 安装步骤 (macOS/Linux)

在 macOS/Linux 上，推荐使用 Python 虚拟环境来避免依赖冲突。

1.  **克隆项目并配置环境**：

    ```bash
    # 克隆项目
    git clone --depth 1 https://github.com/Wxh16144/wechat-bills.git
    cd wechat-bills

    # 创建虚拟环境 (仅需一次)
    python3 -m venv venv

    # 激活虚拟环境 (每次新开终端都需要执行)
    source venv/bin/activate
    ```

2.  **安装依赖**：

    ```bash
    pip install -r requirements.txt
    ```

### 使用方法

脚本支持直接指定输入文件或目录。

#### 1. 转换单个文件

```bash
python WeChatBillsPdf2Xlsx.py ./账单2025.pdf
```

这将在同目录下生成 Excel 文件。

#### 2. 转换整个目录

```bash
python WeChatBillsPdf2Xlsx.py ./所有账单/
```

这将会扫描该目录下所有的 `.pdf` 文件并进行转换。

#### 3. 指定输出路径 (可选)

你可以显式指定输出文件的位置：

```bash
python WeChatBillsPdf2Xlsx.py ./my_bill.pdf ./output/2026_bill.xlsx
```

## 注意事项

*   PDF 文件必须是微信导出的原生电子账单。
*   涉及个人隐私，请妥善保管转换后的文件。
*   脚本完全在本地运行，不会上传任何数据。
