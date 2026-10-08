# Bilibili 评论爬取网页端

用 Python 分页拉取 Bilibili 视频评论和楼中楼回复。本地网页可按用户名筛选、收集单条评论及其全部回复，并把当前已加载内容导出为 Excel。

数据来自 Bilibili 公开接口，仅供个人学习研究。请自行遵守平台服务条款，控制请求频率。

## 功能

- 输入 BV 号或视频链接，按热度分页加载评论
- 楼中楼回复不足时自动补全
- 按用户名筛选（主评论或回复中出现即保留）
- 收集单条评论及其全部回复，保存为 JSON
- 导出当前已加载评论为 Excel（有筛选词时只导出筛选结果）

## 环境

- Python 3.10+

## 安装与运行

```bash
pip install -r requirements.txt
python bilibili_comments_web.py
```

浏览器打开 http://127.0.0.1:8000

## 使用说明

1. 输入 BV 号，例如 `BV1xx411c7mD`，或直接粘贴视频链接，点击「开始爬取」。
2. 需要更多评论时点击「加载下一页」。导出的是已经加载到页面里的内容，不是视频下的全部评论。
3. 在筛选框输入用户名后点击「筛选」。
4. 「收集此评论及全部回复」会把该楼层写入 `collected/`。
5. 「导出 Excel」会下载 `.xlsx`。工作表「说明」记录视频信息，「评论」按主评论、回复分层列出。

## 目录

```text
.
├── bilibili_comments_web.py   # 服务入口与评论接口
├── templates/index.html       # 页面
├── static/style.css           # 样式
├── collected/                 # 运行后生成的 JSON（不纳入版本库）
├── requirements.txt
├── LICENSE
└── README.md
```

## 许可证

[MIT](LICENSE)
