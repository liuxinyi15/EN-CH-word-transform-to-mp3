from flask import Flask, render_template, request, send_file
import pandas as pd
import edge_tts
import asyncio
import os

app = Flask(__name__)

VOICE = "zh-CN-XiaoxiaoNeural"

async def tts(text, output_path, rate, voice):
    communicate = edge_tts.Communicate(text, voice=voice, rate=rate)
    await communicate.save(output_path)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    file = request.files['file']
    filename = request.form['filename']
    repeat = int(request.form['repeat'])
    rate = request.form['rate']
    voice = request.form['voice']

    df = pd.read_excel(file)
    text = ""

    has_chinese = 'Chinese' in df.columns

    for _, row in df.iterrows():
        eng = str(row['English'])
        if has_chinese:
            chn = str(row['Chinese'])
            for _ in range(repeat):
                text += f"{eng}, {chn}. "
        else:
            for _ in range(repeat):
                text += f"{eng}. "

    output_file = f"{filename}.mp3"

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    loop.run_until_complete(tts(text, output_file, rate, voice))
    loop.close()

    return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
