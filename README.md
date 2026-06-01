# 基线测算网页工具

上传 Excel 后，按基线规则自动计算各户号 96 点/48 点基线，并导出结果 Excel。

## 功能

- 支持上传单个 `.xlsx` 文件
- 多用户账号体系（用户名+密码）
- 用户隔离工作空间（每个用户仅可查看自己的任务与结果）
- 自动识别列：`户号`、`户名`、`DATA_DATE`（或 `日期`）、`H00`~`H95`
- 支持：
  - 正调节（默认）：阈值规则 `Pavi < 0.75*Pav` 剔除
  - 负调节：阈值规则 `Pavi > 1.25*Pav` 剔除
- 自动按响应日类型选参考日：
  - 工作日：前 5 个工作日
  - 周六：前 3 个周六
  - 周日：前 3 个周日
- 参考日阈值校验后，剔除计算时段平均负荷最低日，再取其余参考日平均得到基线
- 导出：
  - 计算口径
  - 计算说明
  - 参考日 Pavi 明细
  - 各户号基线宽表（96 点/48 点）
  - 各户号基线长表
  - 异常信息（若有）

## 时间映射口径

- `H00 = 00:15`, `H95 = 24:00`
- 48 点转换：第 i 点 = `(H[2i] + H[2i+1]) / 2`
  - 时刻为：`00:30, 01:00, ..., 24:00`

## 启动方式

```bash
cd /Users/1916597037qq.com/baseline_web_app
python3 -m pip install -r requirements.txt
python3 app.py
```

启动后打开浏览器：

[http://127.0.0.1:8501](http://127.0.0.1:8501)

## 对外发布（Render）

你的代码已经在 GitHub，可直接用 Render 发布成公网地址，其他电脑可直接访问。

### 一次性配置

1. 打开 [Render Dashboard](https://dashboard.render.com/)
2. 新建 `Web Service`，选择仓库：`AimePython/Virtual-power-plant_ZJ`
3. 设置：
   - Root Directory: `baseline_web_app`
   - Runtime: `Python`
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn -w 2 -k gthread --threads 4 -b 0.0.0.0:$PORT app:app`
4. 点击 Deploy，等待构建完成

部署成功后会拿到类似 `https://xxxx.onrender.com` 的网址，可在任意电脑访问。

### 已准备好的部署文件

- `Procfile`
- `render.yaml`
- `runtime.txt`
- `requirements.txt`（已包含 `gunicorn`）

### 访问控制建议（可选）

- 如果担心外部随意访问，建议再加一层登录密码（Basic Auth）或仅对内网/白名单开放。
- 当前版本已内置登录页，可通过环境变量配置密码：
  - `APP_PASSWORD`：旧版单密码入口（已兼容，多用户版可忽略）
  - `APP_SECRET_KEY`：会话签名密钥（建议配置随机长字符串）
  - `ADMIN_USERNAME`：初始化管理员用户名（默认 `admin`）
  - `ADMIN_PASSWORD`：初始化管理员密码（默认 `admin123456`，请务必修改）
  - 本地示例：
    ```bash
    export APP_PASSWORD='你的强密码'
    export APP_SECRET_KEY='随机长字符串'
    python3 app.py
    ```

## 备注

- 若响应日前最近日期数据不完整（例如缺少临近几天），工具会使用文件中可用日期内满足规则的最近参考日进行计算。
- 建议正式结算前补齐响应日前完整历史数据后重算。
