# Word文档格式规范工具

这是一个用于标准化Word文档格式的工具，可以识别和处理标题、图片说明和冗余信息，生成格式统一的文档。

## 功能特点

- 自动识别标题和正文内容
- 去除图片说明和冗余信息
- 统一字体和格式设置
- 支持AI智能分析关键词
- 批量处理多个文档
- 配置保存和加载
- 实时预览效果

## 安装方法

### 方法一：直接使用

1. 安装依赖包：
   ```
   pip install -r requirements.txt
   ```

2. 运行应用：
   ```
   streamlit run WordFormatter_GUI.py
   ```

### 方法二：打包为可执行文件

1. 安装Nuitka和依赖：
   ```
   pip install nuitka
   pip install -r requirements.txt
   ```

2. 运行打包脚本：
   ```
   python build_app.py
   ```

3. 打包完成后，可执行文件将位于`dist`目录下

## 使用说明

1. 单文件处理：上传Word文档，点击"开始处理"
2. 批量处理：选择多个文档，点击"开始批量处理"
3. 设置界面：
   - 基本设置：调整字体、大小和关键词
   - AI智能：配置OpenAI API用于智能分析
   - 系统设置：配置文件目录和恢复默认值
   - AI助手：与AI聊天，优化关键词识别

## 打包注意事项

打包后的应用包含以下文件：
- WordFormatter_launcher.exe - 主程序
- 其他依赖库和资源文件

如需自定义图标，请在根目录放置favicon.ico文件。

## 配置文件

配置文件默认保存在：`C:\Users\用户名\AppData\Roaming\WordFormatter\config.json`

可通过系统设置更改配置文件位置。 