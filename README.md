# stden
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

spreadSheet To Data Exchange Methods

可以将 Excel 表格转化为有复杂层次结构的 json 的项目

## 开始使用

### 安装

**目前尚未发布，请使用 uv 自行打包安装。**

使用 uv 安装。
```bash
uv build
uv tool install dist/stdem-0.1.0-py3-none-any.whl
```

### 使用

```bash
stdem -dir EXCEL_PATH -o JSON_PATH
```

将 EXCEL_PATH 和 JSON_PATH 分别替换为具体的目录，即可开始使用。

**注意：上述命令会清空 JSON_PATH 中的所有文件，并尝试转换EXCEL_PATH中的所有文件**

## 许可证

本项目基于 MIT 许可证开源，详情请见 [LICENSE](LICENSE) 文件。
