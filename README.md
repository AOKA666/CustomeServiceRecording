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

## 局域网共享给同事

程序现在会监听局域网地址。同事可以在同一个网络下访问你的电脑：

```text
http://你的电脑局域网IP:5000
```

查看你的电脑 IP：

```powershell
ipconfig
```

找到 `IPv4 地址`，例如：

```text
192.168.1.23
```

同事访问：

```text
http://192.168.1.23:5000
```

控制权在你的电脑上：

- 关闭 `start.bat` 窗口后，同事无法继续使用。
- 电脑休眠、关机或断网后，同事无法继续使用。
- Windows 防火墙阻止 5000 端口后，同事无法访问。

## 访问密码

如果要限制同事访问，在项目根目录创建 `.env`：

```text
APP_PASSWORD=你的密码
```

设置后，浏览器打开页面时会弹出账号密码框。用户名可以随便填，密码填 `APP_PASSWORD` 的值。

如果不设置 `APP_PASSWORD`，则局域网内知道地址的人都可以访问。

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

## 当前配置方式

右上角“配置”现在只填写“姓名”。姓名保存在每个同事自己的浏览器里，用于替换提交文案中的固定姓名。

高顿 token 不再通过页面填写，统一写在项目根目录 `.env`：

```text
GAODUN_AUTH_TOKEN=Basic xxx
```

可选访问密码也写在 `.env`：

```text
APP_PASSWORD=change-me
```
