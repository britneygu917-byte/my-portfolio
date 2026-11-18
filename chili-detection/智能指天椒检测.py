import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from ultralytics import YOLO
import threading
import time
import json
import pyaudio
from vosk import Model, KaldiRecognizer

import wave
import pygame
import os

piper = None
piper_voice = None

def init_piper():
    """
    初始化Piper语音合成引擎
    """
    global piper, piper_voice
    try:
        from piper import PiperVoice
        model_path = r"D:\指天椒检测\vosk\piper-models\piper-models\zh_CN-huayan-medium.onnx"
        if os.path.exists(model_path):
            piper_voice = PiperVoice.load(model_path)
            print("Piper语音合成引擎初始化成功")
            return True
        else:
            print(f"模型文件不存在: {model_path}")
            return False
    except Exception as e:
        print(f"Piper初始化失败: {e}")
        return False

piper_loaded = init_piper()

def speak_with_piper(text):
    """
    使用Piper库进行语音合成并播放
    """
    global piper_voice
    
    if not piper_voice:
        return False
    
    try:
        audio_stream = piper_voice.synthesize(text)
        audio_data = []
        sample_rate = 22050
        
        for chunk in audio_stream:
            sample_rate = chunk.sample_rate
            audio_int16 = (chunk.audio_float_array * 32767).astype(np.int16)
            audio_data.append(audio_int16.tobytes())
        
        audio_bytes = b''.join(audio_data)
        
        if len(audio_bytes) > 0:
            output_file = "temp_piper.wav"
            with wave.open(output_file, 'wb') as wav_file:
                wav_file.setnchannels(1)
                wav_file.setsampwidth(2)
                wav_file.setframerate(sample_rate)
                wav_file.writeframes(audio_bytes)
            
            pygame.mixer.init()
            pygame.mixer.music.load(output_file)
            pygame.mixer.music.play()
            
            while pygame.mixer.music.get_busy():
                time.sleep(0.1)
            
            pygame.mixer.quit()
            os.remove(output_file)
            return True
    except Exception as e:
        print(f"Piper播放错误: {e}")
        return False
    
    return False

# ==========================================
# 第一部分：语音助手模块（小白）
# ==========================================

class VoiceAssistant:
    """
    语音助手"小白"类
    负责语音识别、语音合成、以及回答用户问题
    """
    
    def __init__(self):
        # 唤醒状态标志：0=未唤醒，1=已唤醒
        self.is_awake = False
        
        # 当前检测到的数据（从主程序获取）
        self.current_data = {
            'mode': 1,  # 当前模式
            'disease_count': 0,  # 病虫害数量
            'disease_types': [],  # 病虫害类型列表
            'maturity_count': 0,  # 成熟度检测数量
            'chili_count': 0,  # 辣椒数量
            'dried_chili_count': 0,  # 干辣椒数量
            'green_leaf_count': 0,  # 绿叶数量
        }
        
        # 唤醒词
        self.wake_word = "语音助理小白"
        
        # 初始化语音合成（使用pyttsx3，离线可用）
        self.init_tts()
        
        # 初始化语音识别
        self.init_vosk()
        
        print("=== 语音助手'小白'初始化完成 ===")
    
    def init_tts(self):
        """
        初始化语音合成引擎（文字转语音）使用Piper
        """
        print("语音合成引擎初始化完成（使用Piper）")
        self.tts_engine = None
    
    def speak(self, text):
        """
        语音合成：将文字转换为语音播放
        
        参数:
            text: 要播放的文字
        """
        print(f"[小白说]: {text}")
        
        # 优先使用Piper语音合成（更自然的中文语音）
        if piper_loaded:
            try:
                result = speak_with_piper(text)
                if result:
                    print("Piper语音播报成功")
                    return
            except Exception as e:
                print(f"Piper失败: {e}")
        
        # 备用方法：使用Windows SAPI.SpVoice
        try:
            import win32com.client
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            speaker.Speak(text)
            print("SAPI.SpVoice语音播报成功")
            return
        except Exception as e:
            print(f"SAPI.SpVoice失败: {e}")
        
        # 所有方法都失败，只显示文字
        print(f"[小白说-文字]: {text}")
    
    def init_vosk(self):
        """
        初始化Vosk语音识别引擎（语音转文字）
        Vosk是离线的，不需要网络连接
        """
        try:
            from vosk import SetLogLevel
            SetLogLevel(-1)  # 关闭日志输出，保持界面整洁
            
            # 加载中文语音识别模型
            # 模型路径需要根据你的实际路径修改
            import os
            # 使用相对路径，避免中文字符编码问题
            model_path = "../../vosk/hello_world/vosk-model-small-cn-0.22"
            print(f"使用相对路径: {model_path}")
            # 检查路径是否存在
            abs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "vosk", "hello_world", "vosk-model-small-cn-0.22")
            abs_path = os.path.normpath(abs_path)
            print(f"绝对路径: {abs_path}")
            if os.path.exists(abs_path):
                print("路径存在，尝试加载模型...")
                self.vosk_model = Model(model_path)
            else:
                print(f"路径不存在: {abs_path}")
                self.vosk_model = None
            
            # 初始化音频流参数
            self.RATE = 16000  # 采样率，必须和模型一致
            self.CHUNK = 4000  # 每次读取的音频数据大小
            
            print("语音识别引擎初始化成功")
        except Exception as e:
            print(f"语音识别初始化失败: {e}")
            self.vosk_model = None
    
    def update_detection_data(self, data):
        """
        更新当前检测数据
        主程序会不断调用这个函数来更新最新的检测数据
        
        参数:
            data: 包含检测数据的字典
        """
        self.current_data.update(data)
    
    def listen_and_respond(self):
        """
        语音助手的核心功能：监听语音并做出回应
        这个方法会在单独的线程中一直运行
        """
        if not self.vosk_model:
            print("语音识别未初始化，语音助手无法启动")
            return
        
        # 初始化音频输入
        audio = pyaudio.PyAudio()
        stream = audio.open(
            format=pyaudio.paInt16,  # 音频格式：16位整数
            channels=1,               # 单声道
            rate=self.RATE,          # 采样率
            input=True,              # 输入模式（录音）
            frames_per_buffer=self.CHUNK
        )
        
        # 初始化识别器
        rec = KaldiRecognizer(self.vosk_model, self.RATE)
        rec.SetWords(True)
        
        print("\n=== 语音助手'小白'已启动 ===")
        print("请说'语音助理小白'来唤醒我")
        print("============================\n")
        
        try:
            while True:
                # 读取音频数据
                data = stream.read(self.CHUNK, exception_on_overflow=False)
                
                # 如果有识别结果
                if rec.AcceptWaveform(data):
                    result = json.loads(rec.Result())
                    
                    # 如果识别到了文字
                    if 'text' in result and result['text']:
                        text = result['text']
                        print(f"[识别到]: {text}")
                        
                        # 处理识别到的文字
                        self.process_voice_command(text)
                        
        except Exception as e:
            print(f"语音识别出错: {e}")
        finally:
            # 清理资源
            stream.stop_stream()
            stream.close()
            audio.terminate()
    
    def process_voice_command(self, text):
        """
        处理识别到的语音命令
        
        参数:
            text: 识别到的文字
        """
        # ========== 第一步：检查唤醒词 ==========
        # 支持多种唤醒词：语音助理小白、小白、助手、助理
        # 但如果包含退出关键词，优先处理退出
        exit_keywords = ["再见", "拜拜", "没事了", "不用了", "结束"]
        if any(keyword in text for keyword in exit_keywords):
            self.speak("好的，有需要随时叫我")
            self.is_awake = False
            return
        
        # 检查唤醒词
        if self.wake_word in text or "小白" in text or "助手" in text or "助理" in text:
            # 如果包含唤醒词，进入唤醒状态
            self.is_awake = True
            self.speak("我在。请问有什么可以帮您？")
            return
        
        # 如果没有被唤醒，不处理其他命令
        if not self.is_awake:
            return
        
        # ========== 第二步：处理各种命令 ==========
        
        # 命令1：询问当前画面显示
        if "画面" in text or "显示" in text or "什么样" in text or "情况" in text or "怎么" in text or "怎样" in text or "如何" in text:
            self.report_current_status()
            return
        
        # 命令2：询问病虫害
        if "病虫害" in text or "病害" in text or "生病" in text or "虫子" in text or "病" in text:
            self.report_disease_status()
            return
        
        # 命令3：询问辣椒成熟度
        if "辣椒" in text or "成熟" in text or "成熟度" in text or "干" in text or "新鲜" in text:
            self.report_maturity_status()
            return
        
        # 未知命令
        self.speak("抱歉。我没有听懂。你可以问其他问题。")
    
    def report_current_status(self):
        """
        报告当前画面状态
        根据不同的检测模式，回答不同的内容
        """
        mode = self.current_data['mode']
        
        if mode == 1:
            # 模式1：病虫害检测
            self.report_disease_status()
        elif mode == 2:
            # 模式2：成熟度检测
            self.report_maturity_status()
        elif mode == 3:
            # 模式3：同时检测
            self.report_combined_status()
    
    def report_disease_status(self):
        """
        报告病虫害检测状态
        """
        disease_count = self.current_data['disease_count']
        disease_types = self.current_data['disease_types']
        
        # 过滤掉"健康"类别，只保留病虫害
        disease_only_types = [d for d in disease_types if d != "健康"]
        
        # 获取唯一的病虫害类型数量（不包括健康）
        unique_diseases = len(set(disease_only_types))
        
        if disease_count == 0:
            self.speak("目前画面中没有检测到病虫害。作物生长状况良好")
        else:
            # 统计每种病虫害的数量
            from collections import Counter
            disease_counter = Counter(disease_only_types)
            
            # 构建回答
            response = f"目前检测到{disease_count}处异常"
            
            # 添加具体类型
            if disease_counter:
                details = []
                for disease, count in disease_counter.items():
                    details.append(f"{disease}{count}处")
                if details:
                    response += "，包括" + "、".join(details)
            
            # 如果病虫害类型超过2种，给出警告
            if unique_diseases >= 1:
                response += "。注意：目前的病虫害种类较多，可能是土壤的酸碱度值过高或过低导致的，建议检测土壤的酸碱度值"
            
            self.speak(response)
    
    def report_maturity_status(self):
        """
        报告成熟度检测状态
        """
        chili_count = self.current_data['chili_count']
        dried_count = self.current_data['dried_chili_count']
        green_count = self.current_data['green_leaf_count']
        total = chili_count + dried_count + green_count
        
        if total == 0:
            self.speak("目前画面中没有检测到辣椒或绿叶")
        else:
            response = f"目前画面检测到"
            
            if chili_count > 0:
                response += f" 新鲜辣椒 {chili_count} 个"
            
            if dried_count > 0:
                if chili_count > 0:
                    response += "、"
                response += f" 干辣椒 {dried_count} 个"
            
            if green_count > 0:
                if chili_count > 0 or dried_count > 0:
                    response += "、"
                response += f"绿叶{green_count}片"
            
            # 检测是否有干辣椒
            if dried_count >= 2:
                # 先播报基本信息
                self.speak(response)
                # 停顿0.2秒
                time.sleep(0.2)
                # 再播报建议
                suggestion = "注意：目前检测的干辣椒过多，可能是土壤的酸碱度，过高或过低导致的建议，检测土壤的酸碱度"
                self.speak(suggestion)
                # 直接返回，避免重复播报
                return
            
            self.speak(response)
    
    def report_combined_status(self):
        """
        报告同时检测模式的状态（病虫害+成熟度）
        """
        disease_count = self.current_data['disease_count']
        chili_count = self.current_data['chili_count']
        dried_count = self.current_data['dried_chili_count']
        green_count = self.current_data['green_leaf_count']
        
        response = "目前画面"
        
        # 病虫害部分
        if disease_count > 0:
            response += f"检测到{disease_count}处病虫害"
        else:
            response += "没有检测到病虫害"
        
        # 成熟度部分
        total_chili = chili_count + dried_count
        if total_chili > 0:
            response += f"，检测到辣椒共{total_chili}个"
            if dried_count > 0:
                response += f"，其中干辣椒{dried_count}个"
        
        if green_count > 0:
            response += f"，绿叶{green_count}片"
        
        self.speak(response)


# ==========================================
# 第二部分：主程序（辣椒检测系统）
# ==========================================

print("=== 辣椒检测系统 ===")
print("正在初始化...")

# 加载中文字体
try:
    font = ImageFont.truetype("simhei.ttf", 24)
    print("中文字体加载成功!")
except:
    try:
        font = ImageFont.truetype("C:/Windows/Fonts/simhei.ttf", 24)
        print("中文字体加载成功!")
    except:
        font = ImageFont.load_default()
        print("中文字体加载失败，使用默认字体")

# 模型：病害检测
print("正在加载病害检测模型...")
# 使用训练目录中的最佳权重
disease_model = YOLO('D:/指天椒检测/指天椒检测/训练好的模型/病虫害训练权重/weights/best.pt')
# 按照data.yaml中的类别顺序
DISEASE_NAMES = ['炭疽病', '蚜虫', '粘虫', '健康', '叶斑病', '白粉虱', '黄化病']
DISEASE_COLORS = [(0,255,0), (255,0,0), (0,255,255), (0,255,255), (255,0,255), (255,255,0), (0,128,255)]
print("病害检测模型加载成功!")

# 模型：成熟度检测
print("正在加载成熟度检测模型...")
maturity_model = YOLO('D:/指天椒检测/指天椒检测/训练好的模型/实地训练权重/train/weights/best.pt')
# xz_dataset 的类别：0: 辣椒, 1: 干辣椒, 2: 绿叶, 3: 病害
MATURITY_NAMES = {
    0: "辣椒",
    1: "干辣椒",
    2: "绿叶",
    3: "病害"
}
MATURITY_COLORS = {
    0: (0, 0, 255),      # 红色 - 辣椒
    1: (148, 0, 211),  # 紫红色 - 干辣椒
    2: (0, 255, 0),      # 绿色 - 绿叶
    3: (255, 0, 0)       # 蓝色 - 病害
}
print("成熟度检测模型加载成功!")

# 摄像头
RTSP_URL = "rtsp://admin:admin123@192.168.1.108:554/cam/realmonitor?channel=1&subtype=0"
cap = cv2.VideoCapture(RTSP_URL)

if not cap.isOpened():
    print("无法连接到摄像头!")
    exit()

print("开始检测...")
print("按 '1' 病害检测")
print("按 '2' 成熟度检测 (YOLO)")
print("按 '3' 同时检测")
print("按 's' 保存, 'q' 退出")
print("按 '+'/'-' 调整病害阈值")
print("按 'm'/'n' 调整成熟度阈值")
print(" === 语音助手功能 === ")
print("说'语音助理小白'唤醒语音助手")
print("询问'目前画面显示是什么样'获取实时数据")
print("===================\n")

# 初始化变量
current_mode = 1
disease_conf = 0.08  # 病害检测阈值
maturity_conf = 0.30  # 成熟度检测阈值
detection_count = 0
show_healthy = True

# ==========================================
# 第三部分：启动语音助手线程
# ==========================================

# 创建语音助手实例
voice_assistant = VoiceAssistant()

# 创建语音助手线程
# daemon=True 表示当主程序结束时，这个线程也会自动结束
voice_thread = threading.Thread(target=voice_assistant.listen_and_respond, daemon=True)
voice_thread.start()

# ==========================================
# 第四部分：主循环（图像检测）
# ==========================================

while True:
    ret, frame = cap.read()
    if not ret:
        print("无法读取帧")
        continue
    
    # 将OpenCV的BGR格式转换为PIL的RGB格式（用于绘制中文）
    frame_pil = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
    draw = ImageDraw.Draw(frame_pil)
    
    # 初始化计数器
    disease_count = 0
    maturity_count = 0
    chili_count = 0
    dried_chili_count = 0
    green_leaf_count = 0
    disease_types = []  # 记录检测到的病虫害类型
    
    # ==================== 检测病害 ====================
    if current_mode == 1 or current_mode == 3:
        # 使用病害检测模型进行预测
        disease_results = disease_model.predict(frame, conf=disease_conf, verbose=False, max_det=100)
        for result in disease_results:
            for box in result.boxes:
                # 获取检测框坐标
                x1, y1, x2, y2 = box.xyxy[0].cpu().numpy()
                # 获取置信度
                conf = box.conf[0].cpu().numpy()
                # 获取类别
                cls = int(box.cls[0].cpu().numpy())
                
                # 记录病虫害类型（用于语音报告）
                disease_types.append(DISEASE_NAMES[cls])
                
                # 根据设置决定是否显示健康叶片
                if cls == 0 and not show_healthy:  # 0是健康类别
                    continue
                
                # 构建标签文字
                label = f"{DISEASE_NAMES[cls]}: {conf:.2f}"
                color = DISEASE_COLORS[cls]
                
                # 健康叶片用细线，病害用粗线
                width = 2 if cls == 0 else 4
                
                # 绘制检测框
                draw.rectangle([x1, y1, x2, y2], outline=color, width=width)
                
                # 绘制标签背景
                bbox = draw.textbbox((x1, y1-28), label, font=font)
                draw.rectangle([bbox[0]-3, bbox[1]-3, bbox[2]+3, bbox[3]+3], fill=(0,0,0))
                # 绘制标签文字
                draw.text((x1, y1-28), label, font=font, fill=color)
                
                disease_count += 1
    
    # ==================== 检测成熟度 (YOLO) ====================
    if current_mode == 2 or current_mode == 3:
        # 使用成熟度检测模型进行预测
        maturity_results = maturity_model.predict(frame, conf=maturity_conf, verbose=False, max_det=100)
        for result in maturity_results:
            for box in result.boxes:
                # 获取检测框坐标
                x1, y1, x2, y2 = box.xyxy[0].cpu().numpy()
                # 获取置信度
                conf = box.conf[0].cpu().numpy()
                # 获取类别
                cls = int(box.cls[0].cpu().numpy())
                
                # 跳过病害类（由病害模型处理）
                if cls == 3:
                    continue
                
                # 构建标签文字
                label = f"{MATURITY_NAMES[cls]}: {conf:.2f}"
                color = MATURITY_COLORS[cls]
                
                # 绘制检测框
                draw.rectangle([x1, y1, x2, y2], outline=color, width=3)
                
                # 绘制标签背景
                bbox = draw.textbbox((x1, y1-28), label, font=font)
                draw.rectangle([bbox[0]-3, bbox[1]-3, bbox[2]+3, bbox[3]+3], fill=(0,0,0))
                # 绘制标签文字
                draw.text((x1, y1-28), label, font=font, fill=color)
                
                # 统计各类别的数量
                maturity_count += 1
                if cls == 0:
                    chili_count += 1
                elif cls == 1:
                    dried_chili_count += 1
                elif cls == 2:
                    green_leaf_count += 1
    
    # ==================== 更新语音助手数据 ====================
    # 将当前检测数据发送给语音助手
    voice_assistant.update_detection_data({
        'mode': current_mode,
        'disease_count': disease_count,
        'disease_types': disease_types,
        'maturity_count': maturity_count,
        'chili_count': chili_count,
        'dried_chili_count': dried_chili_count,
        'green_leaf_count': green_leaf_count,
    })
    
    # ==================== 显示信息 ====================
    # 显示当前模式
    mode_text = {1: "病害检测", 2: "成熟度检测 (YOLO)", 3: "同时检测"}
    healthy_status = "显示" if show_healthy else "隐藏"
    info = f"模式: {mode_text[current_mode]} | 病害: {disease_count} | 成熟度: {maturity_count}"
    
    # 绘制信息栏背景
    bbox = draw.textbbox((10, 10), info, font=font)
    draw.rectangle([bbox[0]-3, bbox[1]-3, bbox[2]+3, bbox[3]+3], fill=(0,0,0))
    draw.text((10, 10), info, font=font, fill=(255,255,255))
    
    # 显示阈值信息
    info2 = f"病害阈值: {disease_conf:.2f} | 成熟度阈值: {maturity_conf:.2f} | 健康叶片: {healthy_status}"
    bbox2 = draw.textbbox((10, 45), info2, font=font)
    draw.rectangle([bbox2[0]-3, bbox2[1]-3, bbox2[2]+3, bbox2[3]+3], fill=(0,0,0))
    draw.text((10, 45), info2, font=font, fill=(255,255,255))
    
    # 显示成熟度详细信息
    if current_mode == 2 or current_mode == 3:
        info3 = f"辣椒: {chili_count} | 干辣椒: {dried_chili_count} | 绿叶: {green_leaf_count}"
        bbox3 = draw.textbbox((10, 80), info3, font=font)
        draw.rectangle([bbox3[0]-3, bbox3[1]-3, bbox3[2]+3, bbox3[3]+3], fill=(0,0,0))
        draw.text((10, 80), info3, font=font, fill=(255,255,255))
    
    # 显示语音助手状态
    voice_status = "已唤醒" if voice_assistant.is_awake else "未唤醒"
    info4 = f"语音助手: {voice_status} | 说'语音助理小白'唤醒"
    bbox4 = draw.textbbox((10, 115), info4, font=font)
    draw.rectangle([bbox4[0]-3, bbox4[1]-3, bbox4[2]+3, bbox4[3]+3], fill=(0,0,0))
    draw.text((10, 115), info4, font=font, fill=(0,255,0) if voice_assistant.is_awake else (255,255,255))
    
    # 将PIL图像转换回OpenCV格式用于显示
    frame = cv2.cvtColor(np.array(frame_pil), cv2.COLOR_RGB2BGR)
    cv2.imshow("辣椒检测", frame)
    
    # ==================== 键盘控制 ====================
    key = cv2.waitKey(1) & 0xFF
    
    if key == ord('s'):
        # 保存当前画面
        detection_count += 1
        cv2.imwrite(f"detection_{detection_count}.jpg", frame)
        print(f"已保存: detection_{detection_count}.jpg")
    elif key == ord('q'):
        # 退出程序
        break
    elif key == ord('1'):
        # 切换到病害检测模式
        current_mode = 1
        print("切换到病害检测模式")
    elif key == ord('2'):
        # 切换到成熟度检测模式
        current_mode = 2
        print("切换到成熟度检测模式 (YOLO)")
    elif key == ord('3'):
        # 切换到同时检测模式
        current_mode = 3
        print("切换到同时检测模式")
    elif key == ord('h'):
        # 切换健康叶片显示
        show_healthy = not show_healthy
        print(f"健康叶片: {'显示' if show_healthy else '隐藏'}")
    
    # 病害阈值控制
    elif key == ord('+'):
        disease_conf = min(0.9, disease_conf + 0.02)
        print(f"病害阈值: {disease_conf:.2f}")
    elif key == ord('-'):
        disease_conf = max(0.02, disease_conf - 0.02)
        print(f"病害阈值: {disease_conf:.2f}")
    
    # 成熟度阈值控制
    elif key == ord('m'):  # increase maturity threshold
        maturity_conf = min(0.9, maturity_conf + 0.05)
        print(f"成熟度阈值: {maturity_conf:.2f}")
    elif key == ord('n'):  # decrease maturity threshold
        maturity_conf = max(0.1, maturity_conf - 0.05)
        print(f"成熟度阈值: {maturity_conf:.2f}")

# ==========================================
# 第五部分：清理资源
# ==========================================

cap.release()
cv2.destroyAllWindows()
print("检测结束")
