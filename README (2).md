# 📝 AI 试卷排版大师 (AI Exam Typesetting Master)

![Version](https://img.shields.io/badge/Version-v2.0_Pro-blue.svg?style=for-the-badge)
![Python](https://img.shields.io/badge/Python-3.10+-success.svg?style=for-the-badge)
![Framework](https://img.shields.io/badge/Streamlit-1.30+-red.svg?style=for-the-badge)
![AI Engine](https://img.shields.io/badge/AI_Engine-Gemini_2.5_Flash-orange.svg?style=for-the-badge)
![Database](https://img.shields.io/badge/Database-SQLite3-lightgrey.svg?style=for-the-badge)

🚀 **专为教务系统打造的工业级 AI 试卷处理 SaaS 平台**。

本项目旨在彻底解决一线教师在整理、排版、组装试卷时面临的排版错乱、公式乱码、图片发黑等痛点。通过引入先进的多模态大模型视觉能力与纯原生 Numpy 矩阵图像算法，实现从“杂乱原卷/截图”到“教务处标准 A3/A4 Word 试卷”的毫秒级转换。

---

## ✨ 核心特性 (Core Features)

### 🧠 一、多模态智能解析引擎
* **Gemini 视觉直连**：支持直接上传 PDF / Docx / Txt 或纯文本粘贴，精准提取文字与图片素材。
* **JSON 智能纠错与修复**：底层内置容错解析器，自动补齐大模型偶尔遗漏的括号，防止解析崩溃。
* **套题/阅读理解级联抽取**：完美识别“一段材料+多道小题”的复杂结构，自动维护父子题目层级。

### 📐 二、究极 Word 排版引擎
* **双栏自适应防爆锁**：针对 A3 双栏排版优化，动态计算选项长度，智能切换 `同行 / 矩阵 / 独立成行`，彻底告别选项重叠。
* **大题号绝对继承**：智能保留原卷“一、二、三”等原始题号格式，不强行重置。
* **教务处级精细渲染**：支持理科“作图区”自动生成占位方框，支持红字高亮生成“教师解析版”，支持填空题红字填入。

### 🖼️ 三、Super Clean 3.0 图像洗白算法
* 无需额外依赖，纯 Numpy 驱动的高斯除法洗图算法。
* 自动识别黑底白字的 CAD 截图或灰暗的手机拍照题干，瞬间反转并漂白底色，确保打印机输出清晰省墨。

### 🗄️ 四、本地化知识库与双向流转
* **SQLite 树状题库**：内置轻量级数据库，支持给题目打标签（如“期中考 / 第一章”），随时检索拉取历史经典题目。
* **Excel 数据打通**：支持将排版好的试题库一键导出为 `.xlsx` 格式，无缝对接各类网课平台或在线刷题系统。

### 🛡️ 五、高阶定制与防伪机制
* **校徽与校名定制**：支持上传学校高清 Logo，并自定义学校名称，自动完美居中于试卷页眉。
* **保密水印系统**：一键开启“机密★内部文件”防伪水印。
* **超页防呆警告**：基于字符容量动态预估排版页数，A3 规格超出 4 页自动标红警告，节约纸张成本。

---

## 🛠️ 快速部署 (Deployment)

本项目采用最轻量化的 Docker Compose 部署方案，开箱即用。

### 1. 环境准备
确保您的 VPS 或本地服务器已安装 `docker` 与 `docker-compose`。

### 2. 获取代码与配置
```bash
# 克隆仓库
git clone https://github.com/kzhx666/exam_ai.git
cd exam_ai

# 创建环境变量文件
cat << 'EOF' > .env
GEMINI_API_KEY="填写您的_Gemini_API_Key"
# 可选：配置本地 OCR 备用接口
OCR_API_URL="http://paddleocr:8866/predict/system"
EOF
```

### 3. 一键启动
```bash
docker-compose up -d
```
启动后，浏览器访问 `http://您的IP:8501` 即可进入系统大屏。

---

## 📖 使用工作流 (Workflow)

系统分为三大核心模块，流转极其顺滑：

1.  **⚡ 全自动智能解析**：
    * 直接拖拽原卷文件或粘贴文本，点击“一键提取”，AI 自动完成题目切分与图文剥离。
2.  **🌉 桥接人机协作 (免 API)**：
    * 针对没有 API Key 的网络环境，支持本地提取文件内容并生成标准 Prompt，发送给外部大模型后，将生成的 JSON 贴回系统即可完成入库。
3.  **🛒 试卷排版组装 & 数据导出**：
    * 在“试题手术台”内进行所见即所得的微调。
    * 从左侧 SQLite 题库搜索历史题目加入购物车。
    * 配置好考试时间、副标题、学校 Logo 与智能分值后，一键下载标准 `.docx` 试卷或 `.xlsx` 数据表。

---
*Built with ❤️ for better education efficiency.*
