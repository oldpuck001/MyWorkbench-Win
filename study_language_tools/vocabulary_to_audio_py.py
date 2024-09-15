# python.py

import os
import shutil
import pandas as pd
from pydub import AudioSegment
import azure.cognitiveservices.speech as speechsdk

def import_vocabulary(request):

    file_path = request.get("data", {}).get("filePath", "")

    vocabulary_file = pd.ExcelFile(file_path)                                   # 使用pandas讀取Excel文件
    sheetnames = vocabulary_file.sheet_names                                    # 獲取所有工作表名稱

    return ['import', sheetnames]

def generate_audio(request):

    speechKey = request.get("data", {}).get("speechKey", "")
    speechRegion = request.get("data", {}).get("speechRegion", "")
    maleVoice = request.get("data", {}).get("maleVoice", "")
    maleRepeats = request.get("data", {}).get("maleRepeats", "")
    maleTranslate = request.get("data", {}).get("maleTranslate", "")
    maleExample = request.get("data", {}).get("maleExample", "")
    femaleVoice = request.get("data", {}).get("femaleVoice", "")
    femaleRepeats = request.get("data", {}).get("femaleRepeats", "")
    femaleTranslate = request.get("data", {}).get("femaleTranslate", "")
    femaleExample = request.get("data", {}).get("femaleExample", "")
    language = request.get("data", {}).get("language", "")
    sheet = request.get("data", {}).get("sheet", "")
    singleAudio = request.get("data", {}).get("singleAudio", "")
    file_path = request.get("data", {}).get("filePath", "")
    save_path = request.get("data", {}).get("savePath", "")
    
    # 獲取文件夾路徑與文件名
    folder_path, file_name = os.path.split(save_path)

    # 去除文件擴展名，生成同名文件夾路徑
    folder_name = os.path.splitext(file_name)[0]

    temporary_folder = os.path.join(folder_path, 'temporary_folder')
    if not os.path.exists(temporary_folder):
        os.makedirs(temporary_folder)

    if singleAudio == True:
        save_folder = os.path.join(folder_path, folder_name)
        if not os.path.exists(save_folder):
            os.makedirs(save_folder)

    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == '.xlsx':
        vocabulary_df = pd.read_excel(file_path, sheet_name=sheet, engine='openpyxl')
    elif file_extension == '.xls':
        vocabulary_df = pd.read_excel(file_path, sheet_name=sheet, engine='xlrd')

    column_word = vocabulary_df.iloc[:, 0]

    # 初始化 Speech SDK 設定
    speech_config = speechsdk.SpeechConfig(subscription=speechKey, region=speechRegion)
    speech_config.set_speech_synthesis_output_format(speechsdk.SpeechSynthesisOutputFormat.Audio48Khz96KBitRateMonoMp3)

    # 准備空的音頻流列表
    audio_segments = []
    num = 0

    for word in column_word:
        if singleAudio == True:
            audio_segment = []

        if maleVoice == True:
            if language == '1':
                voice_name = "en-US-GuyNeural"
            elif language == '2':
                voice_name = "ja-JP-KeitaNeural"
            elif language == '3':
                voice_name = "th-TH-NiwatNeural"

            if maleRepeats == '1':
                word_text = word
            elif maleRepeats == '2':
                word_text = word + '\n' + word
            elif maleRepeats == '3':
                word_text = word + '\n' + word + '\n' + word

            save_file_male = os.path.join(temporary_folder, f'{word}_male.mp3')
            audio_male = synthesize_speech_to_file(word_text, voice_name, save_file_male, speech_config)
            audio_segments.append(audio_male)
            if singleAudio == True:
                audio_segment.append(audio_male)

        if femaleVoice == True:
            if language == '1':
                voice_name = "en-US-JennyNeural"
            elif language == '2':
                voice_name = "ja-JP-NanamiNeural"
            elif language == '3':
                voice_name = "th-TH-PremwadeeNeural"

            if femaleRepeats == '1':
                word_text = word
            elif femaleRepeats == '2':
                word_text = word + '\n' + word
            elif femaleRepeats == '3':
                word_text = word + '\n' + word + '\n' + word

            save_file_female = os.path.join(temporary_folder, f'{word}_female.mp3')
            audio_female = synthesize_speech_to_file(word_text, voice_name, save_file_female, speech_config)
            audio_segments.append(audio_female)
            if singleAudio == True:
                audio_segment.append(audio_female)

        if maleTranslate == True:
            word_text = vocabulary_df.iloc[num, 1]
            save_file_male = os.path.join(temporary_folder, f'{word}_male_translate.mp3')
            audio_male_translate = synthesize_speech_to_file(word_text, 'zh-TW-YunJheNeural', save_file_male, speech_config)
            audio_segments.append(audio_male_translate)
            if singleAudio == True:
                audio_segment.append(audio_male_translate)

        if femaleTranslate == True:
            word_text = vocabulary_df.iloc[num, 1]
            save_file_female = os.path.join(temporary_folder, f'{word}_female_translate.mp3')
            audio_female_translate = synthesize_speech_to_file(word_text, 'zh-TW-HsiaoYuNeural', save_file_female, speech_config)
            audio_segments.append(audio_female_translate)
            if singleAudio == True:
                audio_segment.append(audio_female_translate)

        if maleExample == True:
            word_text = vocabulary_df.iloc[num, 2]
            save_file_male = os.path.join(temporary_folder, f'{word}_male_example.mp3')
            if language == '1':
                voice_name = "en-US-GuyNeural"
            elif language == '2':
                voice_name = "ja-JP-KeitaNeural"
            elif language == '3':
                voice_name = "th-TH-NiwatNeural"
            audio_male_example = synthesize_speech_to_file(word_text, voice_name, save_file_male, speech_config)
            audio_segments.append(audio_male_example)
            if singleAudio == True:
                audio_segment.append(audio_male_example)

        if femaleExample == True:
            word_text = vocabulary_df.iloc[num, 2]
            save_file_female = os.path.join(temporary_folder, f'{word}_female_example.mp3')
            if language == '1':
                voice_name = "en-US-JennyNeural"
            elif language == '2':
                voice_name = "ja-JP-NanamiNeural"
            elif language == '3':
                voice_name = "th-TH-PremwadeeNeural"
            audio_female_example = synthesize_speech_to_file(word_text, voice_name, save_file_female, speech_config)
            audio_segments.append(audio_female_example)
            if singleAudio == True:
                audio_segment.append(audio_female_example)

        if singleAudio == True:
            save_file = os.path.join(save_folder, f'{word}.mp3')
            # 合併音頻段落
            combined_audio_word = sum(audio_segment)
            # 保存為單一 MP3 文件
            combined_audio_word.export(save_file, format="mp3")

        num += 1

    # 合併音頻段落
    combined_audio = sum(audio_segments)

    # 保存為單一 MP3 文件
    combined_audio.export(save_path, format="mp3")
    
    shutil.rmtree(temporary_folder)

    return ['generate']

# 創建 Speech Synthesizer，將音頻保存到文件中
def synthesize_speech_to_file(word_text, voice_name, output_filename, speech_config):
    word_text = word_text
    speech_config.speech_synthesis_voice_name = voice_name
    audio_config = speechsdk.audio.AudioOutputConfig(filename=output_filename)  # 將音頻保存為 MP3 文件
    
    speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=audio_config)
    result = speech_synthesizer.speak_text_async(word_text).get()

    if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
        return AudioSegment.from_mp3(output_filename)  # 讀取剛保存的 MP3 文件
    else:
        return None