# Table Sync & Excel Export

一个 Chrome 插件：在网页右下角提供两个按钮：

- 同步飞书：将当前页面主表格写入飞书多维表格
- 导出Excel：将当前页面主表格下载为 Excel 文件（`.xls`）

## 1. 安装

1. 打开 Chrome `chrome://extensions/`
2. 打开右上角“开发者模式”
3. 点击“加载已解压的扩展程序”
4. 选择目录：`feishu-bitable-sync`

## 2. 飞书配置（首次）

1. 打开插件弹窗，点击“打开飞书设置”
2. 默认只填并保存：
   - `飞书表格完整链接`
   - `应用密钥 App Secret`（来自飞书开放平台应用，不是机器人 Webhook 密钥）
3. 首次若未配置过 App ID，请在“高级设置”补充一次：
   - `App ID`
4. 其余高级项（`user_id_type`、`ignore_consistency_check`）按默认即可

建议勾选：
- 自动创建缺失字段
- 写入来源页面和抓取时间
- 同步前去重（检测到重复记录则跳过）
- 跳过“操作/Action”列

## 3. 使用方式

1. 打开目标网页
2. 使用页面右下角按钮，或插件弹窗按钮：
   - `同步飞书`
   - `导出Excel`

## 4. 导出说明

- 导出为 `.xls`（HTML 表格格式，Excel 可直接打开）
- 自动包含两列元信息：`来源页面`、`抓取时间`

## 5. 注意

- 当前只处理“当前页面可见数据”，不会自动翻页。
- `App Secret` 存在本地浏览器存储；生产场景建议改服务端中转。
- 新版 `batch_create` 请求已携带 `client_token`、`ignore_consistency_check`、`user_id_type`。
