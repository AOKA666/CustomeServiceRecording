# 客服记录台

一个本地运行的客服通话记录工具，用表格方式快速记录来源、项目、电话、微信、省市、备注、项目标签、校区和行为，并支持一键保存到 `info.xlsx`、提交到 CRM。

## 功能

- 表格化录入通话信息
- 来源、项目、项目标签、校区候选项自动加载
- 输入校区后自动带出省、市
- 输入城市后自动匹配省份
- `400` 来源自动使用默认校区：`高顿网校 SEO`
- 点击“保存”后追加写入 `info.xlsx`
- 每次通过 `start.bat` 重启后，`info.xlsx` 会清空并重新生成表头
- 支持配置 CRM 所需的 `auth_token`、`cookie` 和影刀触发 URL

## 目录结构

```text
.
├── crm_tool/
│   ├── app.py          # Flask 主程序，包含页面、接口和 CRM 提交流程
│   ├── start.bat       # Windows 启动脚本
│   └── config.json     # 本地配置文件，不建议同步
├── 校区.txt             # 校区、省、市映射数据
├── 线索宝.txt           # 来源/线索相关映射数据
├── 项目分类.txt         # 项目候选项
├── 项目标签.txt         # 项目标签候选项
├── 紧急联系人.xlsx      # 本地联系人数据，不建议同步
└── info.xlsx            # 运行时保存结果，不建议同步
```

## 启动

在 Windows 上双击：

```text
crm_tool/start.bat
```

或在命令行运行：

```bash
cd crm_tool
python app.py
```

启动后打开：

```text
http://127.0.0.1:5000
```

## 依赖

需要 Python 3.8+。

主要依赖：

```bash
pip install flask requests openpyxl
```

`start.bat` 会自动安装 `flask` 和 `requests`。如果首次运行提示缺少 `openpyxl`，请手动执行上面的安装命令。

## 本地配置

页面右上角“配置”中填写：

- 影刀触发 URL
- CRM `auth_token`
- CRM `cookie`

配置会保存到：

```text
crm_tool/config.json
```

该文件包含本机登录信息，不应提交到 Git。

## 数据文件说明

- `项目分类.txt`：控制“项目”输入框候选项
- `项目标签.txt`：控制不同项目对应的标签候选项
- `校区.txt`：控制校区候选项，并用于自动填充省、市
- `线索宝.txt`：控制部分来源/线索映射
- `info.xlsx`：保存按钮写入的运行时结果

## Git 忽略建议

这些文件通常不应同步：

```gitignore
info.xlsx
infol.xlsx
__pycache__/
*.py[cod]
crm_tool/config.json
紧急联系人.xlsx
```

## 注意事项

- `info.xlsx` 是运行时临时结果；重启程序会清空旧数据。
- 如果修改了 `app.py`，需要重启 `start.bat` 才能加载新代码。
- 不要把 `auth_token`、`cookie` 或联系人数据提交到公开仓库。
