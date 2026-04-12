FROM python:3.10-slim
WORKDIR /app
# 安装图像处理所需的底层 C++ 库
RUN apt-get update && apt-get install -y libgl1 libglib2.0-0 libgomp1 && rm -rf /var/lib/apt/lists/*
# 使用国内源极速安装 Python 依赖
RUN pip install --no-cache-dir -i https://pypi.tuna.tsinghua.edu.cn/simple fastapi uvicorn python-multipart paddlepaddle paddleocr
COPY main.py .
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8866"]