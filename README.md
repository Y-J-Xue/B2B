# B2B 下单模板生成器

一个基于 Streamlit 的 Excel 转换工具：上传原始 Excel，填写少量参数，自动生成可上传到公司系统的下单模板。

## 在线部署建议

这个项目最适合部署到 **Hugging Face Spaces**（Streamlit 模板）或 **Streamlit Community Cloud**。

### Hugging Face Spaces

1. 登录 Hugging Face，创建一个新的 Space
2. 选择 **Streamlit** 模板
3. 将本仓库的内容推送到 Space 对应的 Git 仓库
4. 入口文件使用 `app.py`
5. `requirements.txt` 会自动安装依赖

### 本地运行（开发用）

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 目录说明

- `app.py`：网页入口
- `converter.py`：Excel 解析与模板生成逻辑
- `assets/template.xlsx`：下单模板
- `.streamlit/config.toml`：Streamlit 配置
