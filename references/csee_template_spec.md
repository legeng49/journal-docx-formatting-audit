# 当前学报模板版式档案

数值以 `references/csee_profile.json` 为准。导入新模板时，只覆盖该模板实际写明或在对应样本上设置的字段；模板没有的字段保持原值，不回退到 Word 默认值，也不从档案中删除。

栏底不齐、留白或末页溢出时，用改写表述改变行数。不得把行距压到档案规定值以下，不得缩小图，不得把图内或表内字体降到档案中的图表字号以下。

`audit` 的 `stop.allowed` 为 true 之前，任务不得结束。档案里的每一条闸门都要满足，包括这次模板没有改写、因此仍然有效的旧规定。

## 页面

- 纸宽：`11906` twips（21.0 cm）
- 纸高：`16838` twips（29.7 cm）
- 上边距：`1814` twips（3.2 cm）
- 下边距：`1417` twips（2.5 cm）
- 左边距：`1134` twips（2.0 cm）
- 右边距：`1134` twips（2.0 cm）
- 页眉距边界：`1134` twips（2.0 cm）
- 页脚距边界：`0` twips（0.0 cm）
- 栏距：`442` twips（0.78 cm）
- 文档网格：每页 42 行，每行 45 字，半栏 21 字，行距基准 `324` twips，与网格对齐：True
- 第 1 节 1 栏，第 2 节 2 栏。模板若没有第 2 节，不改已有的第 2 节栏数。
- 页眉：9 pt 宋体+Times New Roman，居中制表位 8.5 cm，居右制表位 16.8 cm。奇偶页眉 True，首页不同 True。

## 角色

- `front_common`：line=single；snap=True
- `front_meta`：size_pt=9；eastAsia=黑体；ascii=Times New Roman；line=single；snap=True
- `title_zh`：size_pt=22；eastAsia=黑体；ascii=Arial；before_pt=12.0；align=center；line=single；snap=True
- `authors`：size_pt=14；eastAsia=仿宋_GB2312；align=center；line=single；snap=True；ascii=仿宋_GB2312
- `affiliation_zh`：size_pt=10.5；eastAsia=楷体_GB2312；ascii=Times New Roman；align=center；line=single；snap=True
- `title_en`：size_pt=12；ascii=Times New Roman；bold=True；align=center；line=single；snap=True
- `authors_en`：size_pt=10.5；ascii=Times New Roman；align=center；line=single；snap=True
- `affiliation_en`：size_pt=9；ascii=Times New Roman；after_pt=12.0；align=center；line=single；snap=True
- `abstract_en`：size_pt=9；ascii=Times New Roman；line_exact_pt=14.0；align=both
- `keywords_en`：size_pt=9；ascii=Times New Roman；line_exact_pt=14.0；before_pt=6.0；align=both
- `abstract_zh`：size_pt=9；eastAsia=宋体；ascii=Times New Roman；label_eastAsia=黑体；label_bold=True；line_exact_pt=14.0；before_pt=6.0；align=both
- `keywords_zh`：size_pt=9；eastAsia=宋体；ascii=Times New Roman；label_eastAsia=黑体；label_bold=True；line_exact_pt=14.0；before_pt=6.0；align=both
- `heading1`：size_pt=12；eastAsia=黑体；ascii=Arial；before_pt=6.0；after_pt=6.0；hanging_twips=0；line=single；snap=True；keep_next=True；bold=True
- `heading2`：size_pt=10.5；eastAsia=黑体；ascii=Arial；before_pt=0；after_pt=0；hanging_twips=0；line=single；snap=True；keep_next=True
- `heading3`：size_pt=10.5；eastAsia=宋体；ascii=Times New Roman；before_pt=0；after_pt=0；hanging_twips=0；line=single；snap=True
- `body`：size_pt=10.5；eastAsia=宋体；ascii=Times New Roman；first_line_twips=420；line=single；snap=True；align=both
- `equation`：line=single；snap=True；before_twips=0；after_twips=0；align=right；mt_full_pt=10.5；mt_sub_pt=6.5；mt_ssub_pt=4.5；mt_symbol_pt=15.0；mt_subsymbol_pt=12.0
- `caption_zh`：size_pt=9；eastAsia=黑体；ascii=Times New Roman；bold=True；align=center；line_exact_pt=14.0；before_pt=3.0；keep_next=True
- `caption_en`：size_pt=9；ascii=Times New Roman；bold=True；align=center；line_exact_pt=14.0；after_pt=3.0；keep_next=True
- `table_cell`：size_pt=7.5；eastAsia=宋体；ascii=Times New Roman；line_exact_pt=14.0；align=center；border_top_bottom_eighths=6；border_inside_eighths=4
- `figure_interior`：size_pt=7.5；eastAsia=宋体；ascii=Times New Roman
- `references_heading`：size_pt=12.0；eastAsia=黑体；ascii=Times New Roman；before_pt=4.0；after_pt=4.0；line=single
- `references`：size_pt=9；eastAsia=宋体；ascii=Times New Roman；line_exact_pt=14.0；hanging_twips=357
- `biography`：size_pt=9；eastAsia=宋体；ascii=Times New Roman；line_exact_pt=14.0；first_line_twips=357；photo_min_px=[260, 330]
- `acknowledgment_heading`：size_pt=10.5；eastAsia=黑体；align=center；before_pt=6.0；line=single；ascii=Times New Roman
- `acknowledgment`：size_pt=9.0；eastAsia=楷体_GB2312；first_line_twips=420；line=single；align=both；ascii=楷体_GB2312
- `appendix_heading`：size_pt=9.0；eastAsia=黑体；ascii=Arial；before_pt=3.0；after_pt=3.0；line=single；snap=False
- `appendix`：size_pt=9；eastAsia=宋体；ascii=Times New Roman；line_exact_pt=14.0；hanging_twips=0
- `editor`：size_pt=10.5；label_eastAsia=黑体；name_eastAsia=楷体_GB2312；align=right；before_pt=6.0；line=single

## 正文、图表、参考文献与页眉

- 段末不单字成行：`True`。参考文献序号和数字后的单位不放在行首：`True`。
- 数字自小数点起每 3 位留 1/4 字空：`True`。数字与单位之间留 1/4 字空：`True`。列项标记：`1）`。
- 图用嵌入型并居中：`True` / `True`。中英文题成对：`True`。主副线、坐标标目、录波格值须目视核对：`True` / `True` / `True`。
- 表头单位用分数，百分数写成 `φB/%`，列用制表位而不是空格对齐：`True` / `True` / `True`。
- 参考文献编号后空 2 小格；括号半角、其余全角：`True`；GB/T 7714 类型标志：`True`；外文刊名不缩写：`True`。
- 符号按 GB 3100～3102 且全文一致，缩写首次出现时给出中文全称(英文全称，缩写)：`True` / `True`。
- 页眉须包含：中国电机工程学报、Proceedings of the CSEE。并核对页眉字号与居中、居右制表位。

## 图表字号

- 图题、表题见 `caption_zh` / `caption_en`。图内文字见 `figure_interior`，当前为 7.5 pt。表内文字见 `table_cell`，当前为 7.5 pt，行距 14.0 pt。
- 图内文字若只存在于位图中，OOXML 读不到字号。必须逐图核对为档案中的图内字号后，审计才可以带 `--figure-fonts-verified`。该开关不跳过任何其他闸门，也不能用来接受更小的字号。

## 版面与停止

- 调整手段：`prose`。禁止压行距：True。禁止缩图：True。禁止缩小图表字体：True。
- 双栏：左右栏顶、底坐标差绝对值不超过 5.25 pt；两栏都有文字的页，栏底不低于 745 pt。
- 编号公式的行距、段前段后和 MathType 尺寸以 `roles.equation` 为准。禁止用更大的固定行距把公式行撑开，也禁止为了省行把行距降到该值以下。
