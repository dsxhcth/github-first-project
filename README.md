# 运动学精讲题 · 讲义生成器

用 Python 生成一份排版好的高中物理「运动学」专题讲义（Word 文档），按手机阅读的版式设计。

## 特点

- 用 `python-docx` 直接生成 `.docx`，不需要提前准备模板文件
- 中文字体、标题层级、行距、页边距全部在代码里定义，换机器也能得到同样效果
- 公式、答案块、要点列表各有独立的排版函数，改样式只改一处
- 页脚自动插入页码

## 运行

```bash
pip install -r requirements.txt
python make_handout.py
```

生成的文档输出到脚本顶部 `OUT_DIR` 指定的目录，文件名是 `运动学精讲题思路_手机版.docx`。

## 代码结构

| 函数 | 作用 |
| --- | --- |
| `add_title` | 主标题 + 副标题 |
| `add_body` | 正文段落，可控制加粗、字号、颜色、缩进 |
| `add_formula` | 居中的公式行 |
| `add_answer` | 浅黄底色的答案块 |
| `add_tip` | 带编号的要点 |
| `add_bullet` | 多级项目符号 |
| `add_page_number_footer` | 页脚页码 |
| `set_cn_font` | 统一设置中文字体（含 eastAsia） |
| `shade` | 给段落加底色 |

## 备注

输出路径写死在代码顶部，换电脑运行时记得改 `OUT_DIR`。

## 许可证

个人学习用途。
