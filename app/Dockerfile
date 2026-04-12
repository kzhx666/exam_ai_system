FROM python:3.10-slim
WORKDIR /app

# 安装底层依赖：增加 catdoc 用于解析老版 .doc 文件
RUN apt-get update && apt-get install -y catdoc && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt

COPY . .
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.address=0.0.0.0"]
