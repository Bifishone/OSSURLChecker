# 🐟 URL 处理工具集 (URL Toolkit)

一个功能完整的OSS中的URL处理工具箱，包含 URL提取、Host信息抽取和有效性检测三大核心功能，帮助你高效处理关于OSS中的URL相关任务。

<img width="1706" height="766" alt="image" src="https://github.com/user-attachments/assets/6595121f-023c-46b3-87ba-04a115928841" />


## 📝 项目介绍

本工具集由三个紧密协作的 Python 脚本组成，形成了从 URL 信息提取到有效性检测的完整工作流：

<img width="1520" height="1383" alt="image" src="https://github.com/user-attachments/assets/9f103d67-7461-4d03-8405-5699f8e7a25d" />




- 从指定 URL 提取关键标签信息并生成结构化 Excel
- 从 Excel 文件中批量抽取 Host 信息并汇总
- 批量检测 URL 有效性并生成可视化检测报告



适用于数据采集、链接验证、信息整理等场景，操作简单且结果可视化程度高。

------

## 🔧 功能模块

### 1. KeyExtract.py - URL 标签提取工具 🔍

**核心功能**：从用户输入的 URL 中提取 XML 标签信息（如 Key、Size、Type 等），并生成格式化的 Excel 文件。



**特性**：

- 自动解析 XML 内容，提取`<Key>`及关联标签（Size/Type/ID/LastModified）
- 生成包含完整 Host 链接的结构化数据
- 自动生成基于域名和路径的安全文件名，避免重复
- 输出 Excel 自动美化（表头样式、边框、列宽自适应）
- 实时显示处理进度和统计信息

### 2. ExtractHost.py - Host 信息抽取工具 📊

**核心功能**：从当前目录下的所有 Excel 文件中提取 Host 列信息，去重后汇总保存。



**特性**：

- 自动扫描当前目录所有`.xlsx`文件
- 精准提取包含 "Host" 列的内容，忽略无 Host 列的文件
- 多文件 Host 信息自动去重合并
- 结果保存为`url.txt`，便于后续批量处理
- 详细日志输出，清晰展示每个文件的处理结果

### 3. OSSURLChecker.py - URL 有效性检测工具 ✅

**核心功能**：批量检测`url.txt`中的 URL 有效性，识别有效链接、无效链接和访问受限链接。



**特性**：

- 基于 Selenium 的无头浏览器检测，模拟真实访问
- 驱动池管理，复用浏览器实例，提升检测效率
- 多线程并发处理，支持动态调整线程数
- 详细的检测结果（状态码、错误信息、资源 ID 等）
- 结果 Excel 自动美化，区分不同状态的 URL
- 实时进度条展示，直观了解检测进度

------

## 🚀 安装说明

### 前置依赖

1. Python 3.8 及以上版本
2. Chrome 浏览器（用于 URL 检测模块）
3. ChromeDriver（需与 Chrome 版本匹配，放置于项目根目录）



### 安装步骤

1. 克隆或下载项目代码到本地

   ```bash
   git clone https://github.com/yourusername/url-toolkit.git
   cd url-toolkit
   ```

2. 安装依赖包

   ```bash
   pip install pandas openpyxl requests beautifulsoup4 selenium
   ```

3. 配置 ChromeDriver

   - 下载与本地 Chrome 版本匹配的 ChromeDriver：[下载地址](https://sites.google.com/chromium.org/driver/)
   - 将下载的`chromedriver.exe`（Windows）或`chromedriver`（Linux/Mac）放入项目根目录

------

## 📋 使用流程

### 步骤 1：提取 URL 标签信息（KeyExtract.py）

```bash
python KeyExtract.py
```

- 输入目标 URL（支持 HTTP/HTTPS 协议）
- 程序自动解析并提取标签信息
- 结果保存为格式化 Excel（如`domain_path_result.xlsx`）

### 步骤 2：抽取 Host 信息（ExtractHost.py）

```bash
python ExtractHost.py
```

- 自动扫描当前目录所有 Excel 文件
- 提取并合并所有 Host 信息（自动去重）
- 结果保存为`url.txt`（每行一个 Host）

### 步骤 3：检测 URL 有效性（OSSURLChecker.py）

```bash
python OSSURLChecker.py
```

- 读取`url.txt`中的所有 URL
- 批量检测并分类（有效 / 无效 / 访问拒绝）
- 结果保存为`result.xlsx`（含详细状态信息）

------

## 📌 注意事项

- 确保 Excel 文件中存在 "Host" 列，否则 ExtractHost.py 会跳过该文件
- ChromeDriver 版本必须与本地 Chrome 浏览器版本匹配，否则检测模块无法运行
- 网络环境可能影响 URL 检测结果，建议在稳定网络下运行
- 大批量 URL 检测时，可能需要适当延长超时时间（可修改代码中`set_page_load_timeout`参数）
- 生成的 Excel 文件支持直接用 Excel 或 WPS 打开，格式已优化

------

## 👨‍💻 作者信息

**作者**：一只鱼（Bifishone）
**功能说明**：本工具集旨在简化 URL 相关的数据处理流程，适用于数据分析师、开发工程师等需要批量处理 URL 的场景。

------



💡 如有问题或建议，欢迎提交 issue 或联系作者！
