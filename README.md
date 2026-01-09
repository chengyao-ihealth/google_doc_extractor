# 提取 Google Doc 内容到 Google Sheets

这个脚本用于从 Google Sheets 的 R 列提取 Google Doc 链接，读取文档内容，然后写入 S 列。

## 设置步骤

### 1. 安装依赖

进入 `google_doc_extractor` 文件夹并安装依赖：

```bash
cd google_doc_extractor
pip install -r requirements.txt
```

或者直接安装：

```bash
pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

### 2. 设置 Google API 凭证

1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 创建新项目或选择现有项目
3. 启用以下 API：
   - Google Sheets API
   - Google Docs API
4. 创建 OAuth 2.0 凭证：
   - 转到 "Credentials" → "Create Credentials" → "OAuth client ID"
   - 应用类型选择 "Desktop app"
   - 下载 `credentials.json` 文件
5. 将 `credentials.json` 文件放在项目根目录

### 3. 配置脚本

脚本已经配置了你的 Google Sheets ID：
- Spreadsheet ID: `1u1W9nV26a8-nvEx8R_bN6tEm3PrK6Z3O6A2YYXlPSLw`
- 源列: R (第18列)
- 目标列: S (第19列)

如果需要修改，可以编辑脚本中的以下变量：
```python
SPREADSHEET_ID = 'your-spreadsheet-id'
R_COLUMN = 17  # Column R
S_COLUMN = 18  # Column S
```

### 4. 运行脚本

```bash
python3 extract_google_doc_content.py
```

首次运行时会：
1. 打开浏览器窗口进行 Google 账户认证
2. 请求访问 Google Sheets 和 Google Docs 的权限
3. 授权后保存凭证到 `google_api_token.pickle` 文件
4. 后续运行无需再次认证（除非凭证过期）

## 脚本功能

- 从 R 列读取所有 Google Doc 链接
- 提取每个链接中的文档内容
- 将内容写入对应的 S 列单元格
- 跳过空单元格
- 处理各种 Google Doc URL 格式
- 支持从表格中提取文本内容

## 注意事项

1. **权限要求**：确保你的 Google 账户有权限访问目标 Google Sheet 和其中的 Google Doc 链接
2. **API 配额**：Google API 有使用配额限制，大量文档可能需要分批处理
3. **错误处理**：脚本会记录无法访问的文档链接
4. **数据格式**：提取的内容是纯文本格式，不包含格式信息

## 故障排除

### 问题：找不到 credentials.json
**解决方案**：确保已下载凭证文件并放在脚本同一目录

### 问题：权限被拒绝
**解决方案**：确保已启用 Google Sheets API 和 Google Docs API，并且凭证配置正确

### 问题：无法访问文档
**解决方案**：确保文档是公开可访问的，或者你的账户有查看权限

### 问题：Sheet 名称不匹配
**解决方案**：脚本默认使用第一个 sheet，如果需要指定特定 sheet，可以修改脚本中的 `sheet_name` 变量
