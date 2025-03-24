# stden
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/iccues/stdem/blob/main/LICENSE)
[![PyPI - Version](https://img.shields.io/pypi/v/stdem)](https://pypi.org/project/stdem/)
![PyPI - Python Version](https://img.shields.io/pypi/pyversions/stdem)


spreadSheet To Data Exchange Methods

可以将 Excel 表格转化为有复杂层次结构的 json 的项目

![](https://github.com/iccues/stdem/blob/main/docs/image/example.png)

可以转化为以下 json 文件

```json
{
    "Nyxra": {
        "hp": 10000,
        "attack": 200.0,
        "skills": ["Shadowstep", "Twilight Veil", "Void Requiem"]
    },
    "Orin": {
        "hp": 15000,
        "attack": 100.0,
        "skills": ["Mana Surge", "Celestial Wrath"]
    }
}
```

## 开始使用

### 安装

使用 pip 安装。

```bash
pip install stdem
```

### 使用

```bash
stdem -dir EXCEL_PATH -o JSON_PATH
```

将 EXCEL_PATH 和 JSON_PATH 分别替换为具体的目录，即可开始使用。

**注意：上述命令会清空 JSON_PATH 中的所有文件，并尝试转换EXCEL_PATH中的所有文件**

## 许可证

本项目基于 MIT 许可证开源，详情请见 [LICENSE](https://github.com/iccues/stdem/blob/main/LICENSE) 文件。
