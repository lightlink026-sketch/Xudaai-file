# Xudaai 本地产品资料整理工具

Windows 桌面程序，用于把已解压的供应商产品文件夹与补充 Excel 整理为标准产品档案，人工复核后上传到运行中的 Xudaai 产品资料库并建立索引。

## 直接使用

下载 `product_importer/dist/Xudaai产品资料整理工具-AI对话版.exe`，双击运行。无需在使用电脑安装 Python。首次使用时填写现有 Xudaai 服务地址和登录凭证。

## 源码和重新打包

源码、依赖和打包脚本在 `product_importer/`。从该文件夹运行 `build_windows.ps1` 可重新生成 exe；也可运行 `start_windows.bat` 在有 Python 3.11 的电脑上直接启动源码版。

## 导入流程

1. 选择供应商提供并已解压的单个产品文件夹；可选择另一个 Excel 作为补充资料。
2. 扫描文件，录入或使用兼容 OpenAI Chat Completions 的模型提取字段。
3. 人工确认产品分类、供应商、型号、SKU、价格、MOQ、包装和重量等资料，并选择主图。
4. 生成标准档案，检查后上传到现有 Xudaai 服务。

源文件不会被修改。1688 链接由使用者自行下载为 Excel 和文件后导入。AI 提取是可选的；使用时，提取的文字会发送到所配置的模型服务。

上传需要目标 Xudaai 服务已部署 `/api/products/import` 接口。该接口会保存产品记录和文件并建立索引。当前已接入原 Xudaai ERP 项目；若部署到其他服务实例，请先合并对应服务端接口。

