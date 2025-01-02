# Word_Creator
此批处理脚本用于根据自定义模板自动在桌面上创建一个 Word 文件，用户可以选择文件名，脚本会将模板文件复制并重命名。默认模板文件应为 Word 文档，解决了新建word文档要点<b>7下<b>的痛点（嗯嗯的）。

### 功能：
- 用户输入文件名，或使用默认值。
- 复制自定义 Word 模板到桌面，并根据文件名生成新文档。
- 若文件已存在，会自动添加后缀（如：0 (1).docx）。
- 文件创建成功后会自动打开 Word 文档。

## 配置说明
此脚本需要手动指定两个文件路径：

- **桌面路径 (`desktopPath`)**：  
  默认设置为 `D:\我的文档\Desktop`，请确保该路径是你的桌面路径。如果你的桌面路径不同，或者使用不同的硬盘驱动器，请修改批处理文件中的 `desktopPath` 变量。

- **模板文件路径 (`templatePath`)**：  
  默认模板文件路径为 `C:\Users\DELL\AppData\Roaming\Microsoft\Templates\新建文档脚本专用（修改默认模板后请复制本文件名重新生成）.docx`，请确保该文件存在。如果你使用的是不同的模板文件或不同的 Windows 用户账户，必须修改该路径以指向正确的模板文件。

### 修改路径：
打开批处理文件 (`.bat` 文件)，在文件顶部找到以下两行：

```batch
set "desktopPath=D:\我的文档\Desktop"
set "templatePath=C:\Users\DELL\AppData\Roaming\Microsoft\Templates\新建文档脚本专用（修改默认模板后请复制本文件名重新生成）.docx"
