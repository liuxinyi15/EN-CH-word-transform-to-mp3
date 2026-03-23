from flask import Flask, render_template, request, send_file, after_this_request
import pandas as pd
import edge_tts
import asyncio
import os
import uuid

app = Flask(__name__)

UPLOAD_FOLDER = 'temp_audio'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

async def tts(text, output_path, rate, voice):
    """
    调用 edge-tts 生成音频
    """
    communicate = edge_tts.Communicate(text, voice=voice, rate=rate)
    await communicate.save(output_path)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    try:

        file = request.files.get('file')
        filename = request.form.get('filename', 'output')
        repeat = int(request.form.get('repeat', 1))
        rate = request.form.get('rate', '-20%')
        voice = request.form.get('voice', 'zh-CN-XiaoxiaoNeural')

        if not file:
            return "请上传 Excel 文件", 400


        df = pd.read_excel(file)
        
        if 'English' not in df.columns:
            return "Excel 文件必须包含 'English' 列", 400
        
        has_chinese = 'Chinese' in df.columns
        text_segments = []


        for _, row in df.iterrows():
            eng = str(row['English']).strip()
            if not eng or eng == 'nan': continue
            
            if has_chinese:
                chn = str(row['Chinese']).strip()
                segment = f"{eng}, {chn}. "
            else:
                segment = f"{eng}. "
            
            text_segments.extend([segment] * repeat)

        full_text = "".join(text_segments)

        if not full_text:
            return "文件中没有有效的词汇内容", 400


        unique_filename = f"{uuid.uuid4()}_{filename}.mp3"
        output_path = os.path.join(UPLOAD_FOLDER, unique_filename)


        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(tts(full_text, output_path, rate, voice))
        finally:
            loop.close()

        @after_this_request
        def remove_file(response):
            try:
                if os.path.exists(output_path):
                    os.remove(output_path)
            except Exception as e:
                app.logger.error(f"Error deleting temporary file: {e}")
            return response

        return send_file(
            output_path, 
            as_attachment=True, 
            download_name=f"{filename}.mp3"
        )

    except Exception as e:
        return f"发生错误: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)