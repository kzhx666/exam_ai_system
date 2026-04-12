import os
import shutil
from fastapi import FastAPI, UploadFile, File
from paddleocr import PaddleOCR

app = FastAPI()

# 启动时初始化加载一次大模型，后续调用秒出结果
ocr = PaddleOCR(use_angle_cls=True, lang="ch")

@app.post("/predict/system")
async def extract_text(file: UploadFile = File(...)):
    # 将上传的图片暂存到临时目录
    temp_path = f"/tmp/{file.filename}"
    with open(temp_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    
    # 调用飞桨进行识别
    result = ocr.ocr(temp_path, cls=True)
    
    # 清理临时图片
    if os.path.exists(temp_path):
        os.remove(temp_path)
    
    # 提取纯文本内容
    text_lines = []
    if result and result[0]:
        for line in result[0]:
            text_lines.append(line[1][0]) # line[1][0] 是识别出的文字字符串
            
    return {"text": "\n".join(text_lines)}