### 一. 环境准备

- **前端**：现代浏览器（Chrome / Firefox / Edge / Safari）
- **Python 脚本**（可选，仅当需要转换 Word 时）：
  - Python 3.6+
  - 安装依赖：`pip install python-docx`

### 二. 使用预置题库（直接练习）

如果你已经有 `questions.json` 文件：
- 将 `index.html` 和 `questions.json` 放在同一目录。
- **必须通过本地服务器访问**（不能直接双击打开，否则 fetch 会失败）：
  ```bash
  # 方法一：使用 Python 自带服务器
  python -m http.server
  # 然后浏览器打开 http://localhost:8000#
  方法二：使用 VS Code Live Server 插件
# 右键 index.html → Open with Live Server

三. 从 Word 题库生成 JSON
1.将你的 Word 题库文件（.docx）放入同一目录，假设文件名为 原始题库.docx。

2.运行转换脚本：
python word_to_json.py 原始题库.docx

3.脚本会在当前目录生成 questions.json。

4.刷新前端页面即可使用新题库。

四.点击页面工具栏的 📂 导入题库 按钮，可选择任意符合格式的 JSON 文件覆盖当前题库。

点击 💾 导出题库 可将当前题库保存为 JSON 文件备份。

五.Word 题库格式要求（示例）
为了脚本能正确解析，建议 Word 文档采用以下简洁格式：

判断题
1. 将物体一个方向的面及其两个坐标轴与投影面平行，投影线与投影面斜交进行投影而得到的图称为轴测斜投影图。( √ )
2. 画管道斜等轴测图常把 OX 轴选定为前后走向的轴。( × )
单选题
1. 工业用户及单独的锅炉房内燃气管道的最高压力不应大于（A ）。
A. 0.8MPa
B. 0.4MPa
C. 0.2MPa
D. 0.1MPa
选项行需以 A.、B. 等开头。答案标记 (A) 可紧跟在题干后。

画图题
题目段落中包含图片（支持内联图片）。

答案单独一行，可包含“答案：”字样及图片。

可根据实际需求调整脚本中的解析逻辑。

六.技术栈
组件	技术
前端	HTML5 / CSS3 / 原生 JavaScript (ES6)
存储	LocalStorage
数据处理	Python 3 + python-docx
部署	静态托管（GitHub Pages、Netlify、Vercel 等）

七.注意事项
不要直接双击 index.html 打开（file:// 协议会导致 fetch 失败）。请使用本地 HTTP 服务器。

Python 脚本依赖 python-docx，请提前安装：pip install python-docx。

图片转换后 Base64 会增大 JSON 文件体积，但便于单文件分发。

如果 Word 文档结构特殊，可自行修改 word_to_json.py 中的解析规则。
