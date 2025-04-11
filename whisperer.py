#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
WhisperTyper - 음성 인식 및 텍스트 변환 프로그램
Copyright (c) 2024 Yeogiaen
"""

import sys, os
import threading
import time
import datetime
import json
import tkinter as tk
from tkinter import messagebox
import logging
import socket
import winsound  # winsound import 추가 확인 (이미 상단에 있을 수 있음)
import re # 정규식 모듈 임포트

# 메시지 모듈 가져오기
from messages import messages, get_message

# 언어 설정 (기본값: 한국어)
current_language = "ko"  # "ko" 또는 "en"
auto_language_detection = False  # 언어 자동 감지 사용 여부
# API 키 단축키 변수
api_key_shortcut = False

# 전역 변수
recording = False
audio_data = []
stream = None
force_clipboard = False
ctrl_pressed = False
shift_pressed = False
alt_pressed = False
recording_started_with_combo = False
selected_device = None  # 선택된 마이크 장치
default_device = None   # 기본 마이크 장치
root = None             # Tk 창
status_label = None     # 상태 표시 레이블
tray_icon = None        # 트레이 아이콘
console_log_file = None # 콘솔 로그 파일
api_key = None          # OpenAI API 키
keyboard_listener = None # 키보드 리스너
pynput_initialized = False # Pynput 초기화 여부

# 중요 모듈들은 비동기적으로 나중에 로드
openai = None
pyperclip = None
sd = None
np = None
soundfile = None
winsound = None
has_tray_support = False
has_pyaudio = False
pystray = None
Image = None
ImageDraw = None
Listener = None
Controller = None
Key = None
KeyCode = None

# 단축키 전역 변수 추가 (파일 상단 전역 변수 섹션에 추가)
hotkey_modifiers = {"ctrl": True, "shift": True, "alt": True}  # 기본 단축키: Ctrl+Shift+Alt
hotkey_key = None  # 추가 키 없음

# 프로그램 다중 실행 방지
def prevent_multiple_instances():
    """프로그램의 다중 실행을 방지합니다"""
    try:
        # 소켓을 생성하여 특정 포트에 바인딩 시도
        global single_instance_socket
        single_instance_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        single_instance_socket.bind(('localhost', 51888))
        logging.info("프로그램 실행: 첫 번째 인스턴스")
        return True  # 첫 번째 인스턴스
    except socket.error:
        logging.warning("프로그램이 이미 실행 중입니다. 중복 실행을 방지합니다.")
        # 이미 실행 중인 프로그램이 있으면 메시지 표시 후 종료
        if not hasattr(sys, 'frozen'):  # 개발 모드에서는 경고만 표시
            print("프로그램이 이미 실행 중입니다.")
        else:  # 배포 버전에서는 GUI 메시지 표시
            try:
                root = tk.Tk()
                root.withdraw()
                messagebox.showwarning(
                    "Yeogiaen WhisperTyper",
                    "프로그램이 이미 실행 중입니다.\n시스템 트레이에서 프로그램 아이콘을 확인하세요."
                )
                root.destroy()
            except:
                pass
        return False  # 이미 다른 인스턴스가 실행 중

# 나중에 필요할 때 모듈 로딩
def load_modules_async():
    """필요한 모듈을 비동기적으로 로딩"""
    global openai, pyperclip, sd, np, soundfile, winsound, pystray, Image, ImageDraw
    global Listener, Controller, Key, KeyCode, has_tray_support, has_pyaudio

    try:
        import openai
        logging.info("OpenAI 모듈 로딩 완료")
    except ImportError as e:
        logging.error(f"OpenAI 모듈 로딩 실패: {str(e)}")

    try:
        import pyperclip
        logging.info("Pyperclip 모듈 로딩 완료")
    except ImportError as e:
        logging.error(f"Pyperclip 모듈 로딩 실패: {str(e)}")

    try:
        import sounddevice as sd
        import numpy as np
        logging.info("Sounddevice 및 Numpy 모듈 로딩 완료")
    except ImportError as e:
        logging.error(f"Sounddevice 또는 Numpy 모듈 로딩 실패: {str(e)}")

    try:
        import soundfile
        logging.info("Soundfile 모듈 로딩 완료")
    except ImportError as e:
        logging.error(f"Soundfile 모듈 로딩 실패: {str(e)}")

    try:
        import winsound
        logging.info("Winsound 모듈 로딩 완료")
    except ImportError as e:
        logging.error(f"Winsound 모듈 로딩 실패: {str(e)}")

    try:
        import pystray
        from PIL import Image, ImageDraw
        has_tray_support = True
        logging.info("Pystray 및 PIL 모듈 로딩 완료")
    except ImportError as e:
        has_tray_support = False
        logging.error(f"Pystray 또는 PIL 모듈 로딩 실패: {str(e)}")

    try:
        from pynput.keyboard import Listener, Controller, Key, KeyCode
        logging.info("Pynput 모듈 로딩 완료")
    except ImportError as e:
        logging.error(f"Pynput 모듈 로딩 실패: {str(e)}")
        log_to_console("오류: 키보드 감지 모듈을 로드할 수 없습니다.")
        return False

    try:
        import pyaudio
        has_pyaudio = True
        logging.info("PyAudio 모듈 로딩 완료")
    except ImportError:
        has_pyaudio = False
        logging.info("PyAudio 모듈 없음, 기본 오디오 시스템 사용")

# Tkinter 라벨 업데이트를 위한 안전한 함수
def set_status(text):
    """상태 라벨 업데이트를 스레드 안전하게 처리"""
    global status_label
    if status_label and status_label.winfo_exists():
        try:
            status_label.config(text=text)
            status_label.update()
        except Exception as e:
            logging.error(f"상태 업데이트 오류: {str(e)}")

# 마이크 초기화를 비동기적으로 수행
def init_microphone_async():
    """마이크 초기화를 비동기적으로 수행"""
    global sd, selected_device, default_device

    if not sd:
        logging.warning("Sounddevice 모듈이 로드되지 않아 마이크 초기화를 건너뜁니다.")
        return

    try:
        # 현재 자동 감지된 기본 장치
        default_device_idx = sd.default.device[0]  # 기본 입력 장치 인덱스

        # 사용 가능한 오디오 장치 목록 가져오기
        devices = sd.query_devices()
        logging.info(f"사용 가능한 오디오 장치: {len(devices)}개 감지됨")

        # 입력 장치 필터링 (마이크만)
        input_devices = []
        for idx, device in enumerate(devices):
            try:
                if device['max_input_channels'] > 0:
                    logging.info(f"입력 장치 #{idx}: {device['name']} (채널: {device['max_input_channels']})")
                    input_devices.append((idx, device))
            except KeyError:
                continue

        if not input_devices:
            logging.error("사용 가능한 입력 장치(마이크)가 없습니다.")
            return

        # 기본 입력 장치 정보 로깅
        try:
            default_device_info = sd.query_devices(kind='input')
            default_device = default_device_idx
            logging.info(f"기본 입력 장치: {default_device_info['name']} (장치 번호: {default_device})")

            # 첫 번째 실행 시 기본 장치를, 아니면 이전에 사용한 장치 계속 사용
            if selected_device is None:
                selected_device = default_device
                logging.info(f"기본 마이크로 선택됨: {default_device_info['name']}")

        except Exception as e:
            logging.error(f"기본 입력 장치 정보 가져오기 오류: {str(e)}")

            # 기본 장치가 없으면 첫 번째 이용 가능한 입력 장치 사용
            if input_devices:
                idx, device = input_devices[0]
                selected_device = idx
                logging.info(f"첫 번째 가용 마이크로 선택됨: {device['name']} (장치 번호: {idx})")

        # 마이크 작동 테스트 (짧은 스트림 생성해보기)
        try:
            if selected_device is not None:
                logging.info(f"마이크 연결 테스트 중 (장치 #{selected_device})...")
                test_stream = sd.InputStream(device=selected_device, channels=1, samplerate=16000)
                test_stream.start()
                time.sleep(0.1)  # 짧게 테스트
                test_stream.stop()
                test_stream.close()
                logging.info("마이크 연결 테스트 성공")
        except Exception as e:
            logging.error(f"마이크 연결 테스트 실패: {str(e)}")
            # 실패해도 계속 진행 (실제 녹음 시 다시 시도)

    except Exception as e:
        logging.error(f"마이크 초기화 오류: {str(e)}")
        log_to_console(f"마이크 초기화 오류: {str(e)}")
        # 실패해도 계속 진행

# 기본 로깅 설정
def setup_logging():
    """로깅 설정"""
    global console_log_file

    # 로그 파일 이름 설정
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')

    # 로그 디렉토리 확인
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # 로그 파일 경로
    log_file = os.path.join(log_dir, f"whisperer_{timestamp}.log")
    console_log_file = "whisperer_console.log"

    # 로그 포맷 설정
    log_format = '%(asctime)s - %(levelname)s - %(message)s'

    # 기본 로거 설정
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )

    # 콘솔 로그 초기화
    try:
        with open(console_log_file, 'w', encoding='utf-8') as f:
            f.write(f"=== Whisperer 로그 시작: {timestamp} ===\n\n")

        # 시작 로그 기록
        log_to_console("=== Yeogiaen WhisperTyper 콘솔 ===")
        log_to_console("이 창을 닫아도 프로그램은 계속 실행됩니다.")
        log_to_console("\n로그 출력을 시작합니다...\n")

        logging.info("로깅 시스템 초기화 완료")
    except Exception as e:
        logging.error(f"로그 파일 초기화 오류: {str(e)}")

# 콘솔 로그 기록 함수
def log_to_console(message):
    """콘솔 로그 파일에 메시지 기록"""
    global console_log_file

    # 콘솔에 출력
    print(message)

    # 파일에 기록
    if console_log_file:
        try:
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            with open(console_log_file, 'a', encoding='utf-8') as f:
                f.write(f"[{timestamp}] {message}\n")
        except Exception as e:
            print(f"로그 기록 오류: {str(e)}")

def main():
    # README 파일 추출
    extract_readme_files()

    # 현재 언어 설정 표시
    lang_name = "English" if current_language == 'en' else "한국어"
    print(f"Current language: {lang_name}")

    try:
        print("프로그램 시작 중...")
        print("============================")
        print("= Yeogiaen WhisperTyper =")
        print("============================")

        # 로깅 설정
        setup_logging()
        logging.info("애플리케이션 시작")
        log_to_console("로깅 시스템 초기화 완료")

        # Tkinter 루트 창 생성 (숨김)
        global root
        root = tk.Tk()
        root.withdraw()  # 창 숨기기
        root.title("Yeogiaen WhisperTyper")

        # 모듈 로딩
        print("주요 모듈 로딩 중...")
        load_modules()
        log_to_console("모듈 로딩 완료")

        # 마이크 초기화 (명시적 호출)
        print("마이크 초기화 중...")
        log_to_console("마이크 초기화 중...")
        init_microphone_async()

        # 설정 로드 (언어, 단축키 등)
        load_settings()
        logging.info(f"현재 언어: {current_language}")
        # 현재 언어 설정을 콘솔에 표시
        language_name = "한국어" if current_language == "ko" else "English"
        log_to_console(f"현재 언어 설정: {language_name} ({current_language})")

        # 현재 단축키 설정 로그
        hotkey_str = []
        if hotkey_modifiers.get("ctrl", False):
            hotkey_str.append("Ctrl")
        if hotkey_modifiers.get("shift", False):
            hotkey_str.append("Shift")
        if hotkey_modifiers.get("alt", False):
            hotkey_str.append("Alt")
        if hotkey_key:
            hotkey_str.append(hotkey_key)

        hotkey_display = "+".join(hotkey_str)
        log_to_console(f"현재 단축키 설정: {hotkey_display}")

        # 녹음 폴더 생성
        recordings_dir = "recordings"
        if not os.path.exists(recordings_dir):
            os.makedirs(recordings_dir)
            log_to_console(f"녹음 폴더 생성됨: {recordings_dir}")

        # API 키 확인 및 설정
        print("API 키 확인 중...")
        log_to_console("API 키 확인 중...")
        if os.path.exists("openai_api_key.txt"):
            try:
                with open("openai_api_key.txt", "r") as f:
                    global api_key
                    api_key = f.read().strip()
                    if openai:
                        openai.api_key = api_key
                    logging.info("API 키 로드 완료")
                    log_to_console("API 키 로드 완료")
            except Exception as e:
                logging.error(f"API 키 로드 오류: {str(e)}")
                log_to_console(f"API 키 로드 오류: {str(e)}")
                log_to_console("API 키 설정 창이 열립니다...")
                show_api_key_dialog(required=True)
        else:
            logging.warning("API 키 파일이 없습니다.")
            log_to_console("API 키 파일이 없습니다. 설정 창을 표시합니다.")
            show_api_key_dialog(required=True)

        # 키보드 리스너 설정
        print("키보드 리스너 설정 중...")
        log_to_console("키보드 리스너 설정 중...")
        setup_keyboard_listener()

        # 시스템 트레이 아이콘 설정
        print("시스템 트레이 아이콘 설정 중...")
        log_to_console("시스템 트레이 아이콘 설정 중...")
        setup_tray_icon()

        # 초기화 완료
        logging.info("초기화 완료")
        print("\n프로그램이 시스템 트레이에서 실행 중입니다.")
        log_to_console("===================================")
        log_to_console("프로그램이 실행 중입니다.")
        log_to_console("시스템 트레이에서 아이콘을 확인하세요.")
        log_to_console("단축키: Ctrl+Shift+Alt (눌렀다 떼면 녹음 시작/종료)")
        log_to_console("===================================")

        # 메인 이벤트 루프 실행
        root.mainloop()

    except Exception as e:
        logging.error(f"메인 함수 오류: {str(e)}")
        print(f"심각한 오류 발생: {str(e)}")
        log_to_console(f"심각한 오류 발생: {str(e)}")
        try:
            import tkinter.messagebox as mbox
            mbox.showerror("오류", f"프로그램 실행 중 오류가 발생했습니다: {str(e)}")
        except:
            pass

# 모든 모듈을 즉시 로드하는 함수
def load_modules():
    """필요한 모듈을 동기적으로 로딩"""
    global openai, pyperclip, sd, np, soundfile, winsound, pystray, Image, ImageDraw
    global Listener, Controller, Key, KeyCode, has_tray_support, has_pyaudio

    try:
        import openai
        print("OpenAI 모듈 로딩 완료")
    except ImportError as e:
        print(f"OpenAI 모듈 로딩 실패: {str(e)}")

    try:
        import pyperclip
        print("Pyperclip 모듈 로딩 완료")
    except ImportError as e:
        print(f"Pyperclip 모듈 로딩 실패: {str(e)}")

    try:
        import sounddevice as sd
        import numpy as np
        print("Sounddevice 및 Numpy 모듈 로딩 완료")
    except ImportError as e:
        print(f"Sounddevice 또는 Numpy 모듈 로딩 실패: {str(e)}")

    try:
        import soundfile
        print("Soundfile 모듈 로딩 완료")
    except ImportError as e:
        print(f"Soundfile 모듈 로딩 실패: {str(e)}")

    try:
        import winsound
        print("Winsound 모듈 로딩 완료")
    except ImportError as e:
        print(f"Winsound 모듈 로딩 실패: {str(e)}")

    try:
        import pystray
        from PIL import Image, ImageDraw
        has_tray_support = True
        print("Pystray 및 PIL 모듈 로딩 완료")
    except ImportError as e:
        has_tray_support = False
        print(f"Pystray 또는 PIL 모듈 로딩 실패: {str(e)}")

    try:
        from pynput.keyboard import Listener, Controller, Key, KeyCode
        print("Pynput 모듈 로딩 완료")
    except ImportError as e:
        print(f"Pynput 모듈 로딩 실패: {str(e)}")

    try:
        import pyaudio
        has_pyaudio = True
        print("PyAudio 모듈 로딩 완료")
    except ImportError:
        has_pyaudio = False
        print("PyAudio 모듈 없음, 기본 오디오 시스템 사용")

# 키보드 리스너 설정
def setup_keyboard_listener():
    """키보드 리스너 설정"""
    global Listener, Key, Controller, KeyCode, pynput_initialized
    load_settings()  # added to ensure hotkey settings from whisperer_settings.json are reloaded

    if Listener is None:
        try:
            from pynput.keyboard import Listener, Key, Controller, KeyCode
            pynput_initialized = True
            logging.info("Pynput 모듈 로드됨")
        except ImportError as e:
            logging.error(f"Pynput 모듈 로드 실패: {str(e)}")
            log_to_console("오류: 키보드 감지 모듈을 로드할 수 없습니다.")
            return False

    def on_press(key):
        """키 누름 이벤트 핸들러"""
        global ctrl_pressed, shift_pressed, alt_pressed, recording, recording_started_with_combo

        try:
            logging.debug(f"키 누름: {key}")

            # 키 상태 업데이트
            if key == Key.ctrl_l or key == Key.ctrl_r:
                ctrl_pressed = True
                logging.debug("Ctrl 키 눌림")
            elif key == Key.shift_l or key == Key.shift_r:
                shift_pressed = True
                logging.debug("Shift 키 눌림")
            elif key == Key.alt_l or key == Key.alt_r:
                alt_pressed = True
                logging.debug("Alt 키 눌림")

            logging.debug(f"현재 키 상태: Ctrl={ctrl_pressed}, Shift={shift_pressed}, Alt={alt_pressed}")
            logging.debug(f"설정된 단축키: 수정자={hotkey_modifiers}, 키={hotkey_key}")

            if not recording:
                # If no 사용자 정의 단축키가 설정되어 있으면 기본 단축키 (Ctrl+Shift+Alt) 사용
                if hotkey_key is None and ctrl_pressed and shift_pressed and alt_pressed:
                    logging.info("기본 녹음 단축키 감지됨 (Ctrl+Shift+Alt)")
                    log_to_console("녹음 시작 단축키 감지...")
                    if root:
                        root.after(10, start_recording)
                    else:
                        start_recording()
                    recording_started_with_combo = True
                    return True

                # 사용자 정의 단축키 조합 확인
                modifier_match = (
                    (not hotkey_modifiers.get("ctrl", False) or ctrl_pressed) and
                    (not hotkey_modifiers.get("shift", False) or shift_pressed) and
                    (not hotkey_modifiers.get("alt", False) or alt_pressed)
                )

                key_match = False
                if hotkey_key:
                    if isinstance(key, KeyCode) and key.char:
                        key_char = key.char
                        # If ctrl is pressed and key_char is a control character, convert it to its corresponding letter
                        if ctrl_pressed and len(key_char) == 1 and ord(key_char) < 32:
                            key_char = chr(ord(key_char) + 64)
                        key_match = key_char.upper() == hotkey_key.upper()
                        logging.debug("키 비교: {} vs {} = {}".format(key_char.upper(), hotkey_key.upper(), key_match))
                    elif key:
                        key_str = str(key).replace('Key.', '')
                        key_match = key_str.upper() == hotkey_key.upper()
                        logging.debug("특수키 비교: {} vs {} = {}".format(key_str.upper(), hotkey_key.upper(), key_match))
                else:
                    key_match = True

                if modifier_match and key_match:
                    logging.info(f"사용자 정의 녹음 단축키 감지됨: 수정자={hotkey_modifiers}, 키={hotkey_key}")
                    log_to_console("사용자 정의 녹음 단축키 감지...")
                    if root:
                        root.after(10, start_recording)
                    else:
                        start_recording()
                    recording_started_with_combo = True
                    return True

        except Exception as e:
            logging.error(f"키 누름 처리 중 오류: {str(e)}")
            return False

    def on_release(key):
        """키 뗌 이벤트 핸들러"""
        global ctrl_pressed, shift_pressed, alt_pressed, recording, recording_started_with_combo

        try:
            # 현재 키 로깅
            logging.debug(f"키 뗌: {key}")

            # 키 상태 업데이트
            if key == Key.ctrl_l or key == Key.ctrl_r:
                ctrl_pressed = False
                logging.debug("Ctrl 키 뗌")
            elif key == Key.shift_l or key == Key.shift_r:
                shift_pressed = False
                logging.debug("Shift 키 뗌")
            elif key == Key.alt_l or key == Key.alt_r:
                alt_pressed = False
                logging.debug("Alt 키 뗌")

            # 녹음 중이고 단축키로 시작했을 때만 처리
            if recording and recording_started_with_combo:
                # 수정자 키(Ctrl, Shift, Alt)가 떼어졌는지 확인
                if (key == Key.ctrl_l or key == Key.ctrl_r or
                    key == Key.shift_l or key == Key.shift_r or
                    key == Key.alt_l or key == Key.alt_r):

                    logging.info("녹음 종료 단축키 감지 (수정자 키 뗌)")
                    log_to_console("녹음 종료 단축키 감지...")

                    # 녹음 종료 (메인 스레드에서 실행)
                    if root:
                        root.after(10, stop_recording)
                    else:
                        stop_recording()
                    recording_started_with_combo = False
                    return True

        except Exception as e:
            logging.error(f"키 뗌 처리 중 오류: {str(e)}")
            return False

        # 기존에 리스너 재시작하는 코드 블록 제거
        return True

    # 리스너 시작
    try:
        # 기존 리스너가 있으면 중지
        global keyboard_listener
        if 'keyboard_listener' in globals() and keyboard_listener is not None:
            try:
                keyboard_listener.stop()
                logging.info("기존 키보드 리스너 중지됨")
            except:
                pass

        # 새 리스너 생성 및 시작
        keyboard_listener = Listener(on_press=on_press, on_release=on_release)
        keyboard_listener.daemon = True  # 데몬 스레드로 설정
        keyboard_listener.start()
        logging.info("키보드 리스너 시작됨")
        log_to_console("키보드 리스너가 시작되었습니다. (단축키 설정: 녹음)")
        return True

    except Exception as e:
        logging.error(f"키보드 리스너 시작 실패: {str(e)}")
        log_to_console(f"오류: 키보드 리스너 시작 실패: {str(e)}")
        return False

# 언어 설정 저장/로드 함수 수정
def save_settings():
    try:
        settings = {
            "language": current_language,
            "auto_detection": auto_language_detection,
            "hotkey": {
                "modifiers": hotkey_modifiers,
                "key": hotkey_key
            }
        }
        print(f"저장할 설정: {settings}")
        with open('whisperer_settings.json', 'w', encoding='utf-8') as f:
            json.dump(settings, f)
            # 파일 강제 쓰기
            f.flush()
            os.fsync(f.fileno())
        logging.info("설정 저장 완료")
        print("설정 파일 저장됨: whisperer_settings.json")
    except Exception as e:
        logging.error(f"설정 저장 오류: {str(e)}")
        print(f"설정 저장 오류: {str(e)}")

def load_settings():
    global current_language, hotkey_modifiers, hotkey_key, auto_language_detection
    try:
        if os.path.exists('whisperer_settings.json'):
            with open('whisperer_settings.json', 'r', encoding='utf-8') as f:
                settings = json.load(f)
                print(f"로드된 설정: {settings}")
                if "language" in settings:
                    current_language = settings["language"]
                if "auto_detection" in settings:
                    auto_language_detection = settings["auto_detection"]
                if "hotkey" in settings:
                    if "modifiers" in settings["hotkey"]:
                        hotkey_modifiers = settings["hotkey"]["modifiers"]
                    if "key" in settings["hotkey"]:
                        hotkey_key = settings["hotkey"]["key"]
            logging.info(f"설정 로드 완료: 언어={current_language}, 단축키 수정자={hotkey_modifiers}, 단축키={hotkey_key}")
            print(f"설정 로드 완료: 언어={current_language}, 단축키 수정자={hotkey_modifiers}, 단축키={hotkey_key}")
        else:
            print("설정 파일이 없습니다. 기본 설정을 사용합니다.")
    except Exception as e:
        logging.error(f"설정 로드 오류: {str(e)}")
        print(f"설정 로드 오류: {str(e)}")

# 단일 설정 함수들 (이전 코드와의 호환성)
def save_language_setting():
    save_settings()

def load_language_setting():
    load_settings()

# 메시지 가져오기 함수 래퍼
def get_msg(key, *args):
    global current_language
    return get_message(key, *args, language=current_language)

# 시스템 트레이 아이콘 이미지 생성 및 설정 함수
def create_image():
    """시스템 트레이 아이콘 이미지 생성"""
    global Image, ImageDraw

    try:
        # 모듈이 로드되었는지 확인
        if Image is None or ImageDraw is None:
            from PIL import Image, ImageDraw
            logging.info("PIL 모듈 로딩 완료")

        # 리소스 파일 경로 계산 (패키지 내부 또는 현재 디렉토리)
        favicon_paths = []

        # 실행 파일 경로 기준 (PyInstaller로 패키징된 경우)
        base_path = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        favicon_paths.append(os.path.join(base_path, "favicon.ico"))

        # 현재 작업 디렉토리 기준
        favicon_paths.append(os.path.abspath("favicon.ico"))

        # 현재 스크립트 디렉토리 기준
        script_dir = os.path.dirname(os.path.abspath(__file__))
        favicon_paths.append(os.path.join(script_dir, "favicon.ico"))

        # 모든 가능한 경로 로깅
        logging.info(f"favicon.ico 파일 가능한 경로들: {favicon_paths}")

        # 각 경로 시도
        favicon_loaded = False
        for favicon_path in favicon_paths:
            logging.info(f"favicon.ico 파일 경로 시도: {favicon_path}")

            if os.path.exists(favicon_path):
                try:
                    # 이미지 로드
                    img = Image.open(favicon_path)
                    logging.info(f"이미지 로드됨: {favicon_path}, 크기 {img.size}, 포맷 {img.format}, 모드 {img.mode}")

                    # 16x16 크기로 조정 (트레이 아이콘에 최적)
                    if img.size != (16, 16):
                        img = img.resize((16, 16), Image.LANCZOS)
                        logging.info("이미지 크기 16x16으로 조정됨")

                    favicon_loaded = True
                    return img
                except Exception as e:
                    logging.error(f"이미지 로드 오류: {favicon_path}: {str(e)}")
                    import traceback
                    logging.error(traceback.format_exc())

        if not favicon_loaded:
            logging.error("모든 favicon.ico 파일 경로 시도 실패")
            log_to_console("favicon.ico 파일을 찾을 수 없습니다.")

            # 배포 버전이 아닌 경우에만 기본 아이콘 생성 (개발 모드)
            if not hasattr(sys, 'frozen'):
                try:
                    logging.info("기본 아이콘 생성 중...")
                    image = Image.new('RGB', (16, 16), color=(73, 109, 137))
                    d = ImageDraw.Draw(image)
                    d.text((4, 4), "W", fill=(255, 255, 0))
                    logging.info("기본 아이콘 생성됨")
                    return image
                except Exception as e:
                    logging.error(f"기본 아이콘 생성 오류: {str(e)}")

    except Exception as e:
        logging.error(f"아이콘 생성 준비 중 오류: {str(e)}")
        log_to_console(f"아이콘 생성 오류: {str(e)}")

    return None

def setup_tray_icon():
    """시스템 트레이 아이콘 설정"""
    global tray_icon, pystray, Image, ImageDraw

    try:
        # 필요한 모듈이 로드되었는지 확인
        if pystray is None:
            logging.info("pystray 모듈 로딩 중...")
            import pystray
            logging.info("pystray 모듈 로딩 완료")

        if Image is None or ImageDraw is None:
            logging.info("PIL 모듈 로딩 중...")
            from PIL import Image, ImageDraw
            logging.info("PIL 모듈 로딩 완료")
    except ImportError as e:
        logging.error(f"트레이 아이콘 설정 실패 - 모듈 로딩 오류: {str(e)}")
        return None
    except Exception as e:
        logging.error(f"트레이 아이콘 설정 중 예상치 못한 오류: {str(e)}")
        return None

    try:
        # 아이콘 이미지 생성
        logging.info("트레이 아이콘 이미지 생성 중...")
        icon_image = create_image()
        logging.info("트레이 아이콘 이미지 생성 완료")

        if icon_image is None:
            logging.error("트레이 아이콘 이미지 생성 실패")
            if root:
                root.after(10, lambda: messagebox.showerror("아이콘 오류", "트레이 아이콘 이미지를 생성할 수 없습니다.\nfavicon.ico 파일이 존재하는지 확인하세요."))
            return None

        # 메뉴 항목 정의
        def exit_action(icon, item):
            icon.stop()
            os._exit(0)

        # 녹음 파일 폴더 열기 함수
        def open_recordings_folder(icon, item):
            try:
                recordings_dir = os.path.abspath("recordings")
                if not os.path.exists(recordings_dir):
                    os.makedirs(recordings_dir)
                # Windows에서 폴더 열기
                os.startfile(recordings_dir)
                log_to_console(f"녹음 파일 폴더 열기: {recordings_dir}")
            except Exception as e:
                log_to_console(f"녹음 파일 폴더 열기 오류: {str(e)}")

        # README 파일 열기 함수
        def open_readme_file(icon, item):
            try:
                # 현재 언어 설정에 따라 README 파일 선택
                readme_file = "README.KR.md" if current_language == "ko" else "README.md"
                readme_path = os.path.abspath(readme_file)

                if os.path.exists(readme_path):
                    # Windows에서 기본 앱으로 파일 열기
                    os.startfile(readme_path)
                    log_to_console(f"README 파일 열기: {readme_path}")
                else:
                    log_to_console(f"README 파일을 찾을 수 없습니다: {readme_path}")
            except Exception as e:
                log_to_console(f"README 파일 열기 오류: {str(e)}")

        # OpenAI API 키 설정 함수
        def set_openai_api_key(icon, item):
            show_api_key_dialog(required=False)
            log_to_console(get_msg("api_key_saved", "openai_api_key.txt"))
            if openai is not None:
                openai.api_key = api_key

        # 마이크 정보 표시 및 재설정 함수
        def check_microphone(icon, item):
            global selected_device, default_device, sd

            try:
                if sd:
                    # 현재 선택된 마이크 정보
                    current_info = ""
                    try:
                        if selected_device is not None:
                            device_info = sd.query_devices(selected_device)
                            current_info = f"현재 마이크: {device_info['name']} (장치 #{selected_device})"
                            log_to_console(current_info)
                        else:
                            current_info = "선택된 마이크 없음"
                            log_to_console(current_info)
                    except Exception as e:
                        log_to_console(f"마이크 정보 확인 오류: {str(e)}")

                    # 사용 가능한 마이크 다시 검색
                    log_to_console("마이크 초기화 중...")
                    init_microphone_async()
                    log_to_console("마이크 초기화 완료")

                    # 새로 설정된 마이크 정보
                    if selected_device is not None:
                        try:
                            device_info = sd.query_devices(selected_device)
                            log_to_console(f"설정된 마이크: {device_info['name']} (장치 #{selected_device})")
                        except Exception as e:
                            log_to_console(f"마이크 정보 확인 오류: {str(e)}")
                else:
                    log_to_console("오디오 모듈이 로드되지 않았습니다.")
            except Exception as e:
                log_to_console(f"마이크 확인 오류: {str(e)}")

        # 콘솔 창 열기 함수
        def open_console(icon, item):
            try:
                import subprocess

                # 콘솔 창 열기 배치 파일 생성
                with open('open_console.bat', 'w', encoding='utf-8') as f:
                    f.write('@echo off\n')
                    f.write(f'title {get_msg("whisper_console")}\n')
                    f.write('chcp 65001\n')  # UTF-8 인코딩 설정
                    f.write(f'echo {get_msg("console_title")}\n')
                    f.write(f'echo {get_msg("console_info")}\n')
                    f.write('echo.\n')
                    f.write(f'echo {get_msg("log_start")}\n')
                    f.write('echo.\n')
                    # 로그 파일 생성 및 실시간 모니터링
                    f.write('echo > whisperer_console.log\n')
                    f.write(f'echo {get_msg("key_monitoring")}\n')
                    f.write('echo.\n')
                    # 로그 파일 실시간 모니터링 (PowerShell 사용)
                    f.write('powershell -command "Get-Content -Path whisperer_console.log -Wait -Encoding UTF8"\n')

                # 새 콘솔 창에서 배치 파일 실행
                subprocess.Popen(['start', 'open_console.bat'],
                               shell=True,
                               creationflags=subprocess.CREATE_NEW_CONSOLE)

                log_to_console("콘솔 창이 열렸습니다.")

                # 콘솔 로그 파일 경로 설정
                global console_log_file
                console_log_file = os.path.abspath("whisperer_console.log")

                # 콘솔 창이 열렸음을 로그에 기록
                log_to_console(get_msg("console_opened"))
                log_to_console(get_msg("program_status"))
                log_to_console(get_msg("recordings_folder", os.path.abspath("recordings")))

            except Exception as e:
                log_to_console(get_msg("console_open_error", str(e)))

        # 한국어로 변경 함수
        def set_korean_language(icon, item):
            global current_language, auto_language_detection
            current_language = "ko"
            auto_language_detection = False
            save_settings()
            log_to_console(get_msg("language_changed", "한국어"))
            # 트레이 아이콘 메뉴 업데이트
            update_tray_menu()
            
        # 영어로 변경 함수
        def set_english_language(icon, item):
            global current_language, auto_language_detection
            current_language = "en"
            auto_language_detection = False
            save_settings()
            log_to_console(get_msg("language_changed", "English"))
            # 트레이 아이콘 메뉴 업데이트
            update_tray_menu()
            
        # 자동 감지 토글 함수
        def toggle_auto_detection(icon, item):
            global auto_language_detection
            auto_language_detection = not auto_language_detection
            save_settings()
            if auto_language_detection:
                log_to_console(get_msg("auto_detection_enabled"))
            else:
                log_to_console(get_msg("auto_detection_disabled"))
            # 트레이 아이콘 메뉴 업데이트
            update_tray_menu()
            
        # 언어 변경 메뉴 항목 (기존 호환성 유지용)
        def change_language(icon, item):
            global current_language
            # 언어 전환 (한국어 <-> 영어)
            current_language = "en" if current_language == "ko" else "ko"
            auto_language_detection = False
            save_settings()
            log_to_console(get_msg("language_changed", current_language))
            # 트레이 아이콘 메뉴 업데이트
            update_tray_menu()

        # 단축키 설정 함수 추가 (언어 변경 함수 아래에 추가)
        def set_hotkey(icon, item):
            show_hotkey_dialog()
            log_to_console(get_msg("hotkey_updated", "단축키가 업데이트되었습니다."))

        def update_tray_menu():
            # 트레이 아이콘 메뉴 업데이트
            tray_icon.menu = pystray.Menu(
                pystray.MenuItem(get_msg("open_recordings_folder"), open_recordings_folder),
                pystray.MenuItem(get_msg("open_readme"), open_readme_file),
                pystray.MenuItem(get_msg("open_console"), open_console),
                pystray.MenuItem(get_msg("set_openai_api_key"), set_openai_api_key),
                pystray.MenuItem(get_msg("set_hotkey", "단축키 설정"), set_hotkey),
                # 언어 설정 하위 메뉴 추가
                pystray.MenuItem(
                    get_msg("language_menu"),
                    pystray.Menu(
                        pystray.MenuItem(
                            get_msg("korean_language"), 
                            set_korean_language,
                            checked=lambda item: current_language == "ko" and not auto_language_detection
                        ),
                        pystray.MenuItem(
                            get_msg("english_language"), 
                            set_english_language,
                            checked=lambda item: current_language == "en" and not auto_language_detection
                        ),
                        pystray.MenuItem(
                            get_msg("auto_detection"), 
                            toggle_auto_detection,
                            checked=lambda item: auto_language_detection
                        ),
                    )
                ),
                # 기존 언어 변경 메뉴 삭제
                pystray.MenuItem(get_msg("exit"), exit_action)
            )

        # 트레이 아이콘 생성
        logging.info("트레이 아이콘 객체 생성 중...")
        tray_icon = pystray.Icon("whisperer")
        tray_icon.icon = icon_image
        tray_icon.title = "Yeogiaen WhisperTyper"

        # 초기 메뉴 설정
        tray_icon.menu = pystray.Menu(
            pystray.MenuItem(get_msg("open_recordings_folder"), open_recordings_folder),
            pystray.MenuItem(get_msg("open_readme"), open_readme_file),
            pystray.MenuItem(get_msg("open_console"), open_console),
            pystray.MenuItem(get_msg("set_openai_api_key"), set_openai_api_key),
            # 언어 설정 하위 메뉴 추가
            pystray.MenuItem(
                get_msg("language_menu"),
                pystray.Menu(
                    pystray.MenuItem(
                        get_msg("korean_language"), 
                        set_korean_language,
                        checked=lambda item: current_language == "ko" and not auto_language_detection
                    ),
                    pystray.MenuItem(
                        get_msg("english_language"), 
                        set_english_language,
                        checked=lambda item: current_language == "en" and not auto_language_detection
                    ),
                    pystray.MenuItem(
                        get_msg("auto_detection"), 
                        toggle_auto_detection,
                        checked=lambda item: auto_language_detection
                    ),
                )
            ),
            pystray.MenuItem(get_msg("set_hotkey", "단축키 설정"), set_hotkey),
            pystray.MenuItem(get_msg("exit"), exit_action)
        )

        # 백그라운드 스레드에서 트레이 아이콘 실행
        logging.info("트레이 아이콘 실행 준비 완료")
        # 트레이 아이콘을 별도 스레드로 실행
        icon_thread = threading.Thread(target=tray_icon.run, daemon=True)
        icon_thread.start()
        logging.info("트레이 아이콘 스레드 시작됨")

        return tray_icon
    except Exception as e:
        logging.error(f"트레이 아이콘 설정 중 오류 발생: {str(e)}")
        return None

# API 키 설정 대화 상자 표시 함수
def show_api_key_dialog(required=False):
    """API 키 입력 대화 상자를 표시합니다."""
    global api_key, openai

    # 현재 API 키 로드
    current_key = ""
    if os.path.exists("openai_api_key.txt"):
        try:
            with open("openai_api_key.txt", "r") as f:
                current_key = f.read().strip()
        except Exception as e:
            logging.error(f"API 키 파일 읽기 오류: {str(e)}")

    # 완전히 독립적인 모달 대화 상자 생성
    dialog = tk.Toplevel()
    dialog.title(get_msg("openai_api_key_title"))
    dialog.geometry("500x250")  # 창 크기 증가
    dialog.resizable(False, False)

    # 모달 설정 (부모 창 비활성화)
    dialog.transient()
    dialog.grab_set()
    dialog.focus_set()

    # 항상 위에 표시
    dialog.attributes("-topmost", True)

    # 메인 프레임
    frame = tk.Frame(dialog, padx=25, pady=20)  # 여백 증가
    frame.pack(fill=tk.BOTH, expand=True)

    # 텍스트 서식 변경 - 개행을 추가한 라벨
    label_text = ""
    if current_language == "ko":
        label_text = "OpenAI API Key를 입력해 주세요\nAPI Key는 아래 URL에서 확인해 주세요\nhttps://platform.openai.com/api-keys"
    else:
        label_text = "Enter your OpenAI API Key\nYou can find your API key at the URL below\nhttps://platform.openai.com/api-keys"

    tk.Label(frame, text=label_text, font=("Segoe UI", 11), justify=tk.LEFT).pack(anchor="w", pady=(0, 15))

    # 엔트리 위젯
    entry = tk.Entry(frame, font=("Courier New", 11), width=45, bd=2)  # 글꼴 크기 및 테두리 증가
    entry.insert(0, current_key)
    entry.pack(fill=tk.X, pady=8)  # 여백 증가
    entry.selection_range(0, tk.END)  # 전체 텍스트 선택

    # 결과 메시지 라벨
    result_label = tk.Label(frame, text="", font=("Segoe UI", 10))
    result_label.pack(pady=10)  # 여백 증가

    # 버튼 프레임
    btn_frame = tk.Frame(frame)
    btn_frame.pack(pady=10)  # 여백 증가

    # 결과 변수
    result = [None]  # 리스트로 감싸서 참조로 전달

    # 저장 함수
    def save_key():
        key = entry.get().strip()

        if required and not key:
            result_label.config(text=get_msg("no_api_key_warning"), fg="red")
            entry.focus_set()
            return

        try:
            # 파일에 저장
            with open("openai_api_key.txt", "w") as f:
                f.write(key)

            # 전역 변수에 설정
            global api_key
            api_key = key

            # OpenAI 설정
            if openai:
                openai.api_key = key

            # 성공 메시지
            result_label.config(text=get_msg("api_key_saved", "openai_api_key.txt"), fg="green")
            logging.info("API 키가 저장되었습니다.")
            log_to_console(get_msg("api_key_saved", "openai_api_key.txt"))

            # 결과 설정
            result[0] = key

            # 성공 후 창 닫기
            dialog.after(1000, dialog.destroy)  # 대기 시간 감소

        except Exception as e:
            result_label.config(text=f"{get_msg('error')}: {str(e)}", fg="red")
            logging.error(f"API 키 저장 오류: {str(e)}")

    # 취소 함수
    def cancel():
        if required and not api_key:
            result_label.config(text=get_msg("no_api_key_warning"), fg="red")
            return
        dialog.destroy()

    # 저장 버튼
    save_btn = tk.Button(btn_frame, text=get_msg("save"), command=save_key,
                         width=12, height=1, bg="#4CAF50", fg="white", font=("Segoe UI", 10, "bold"))  # 크기 증가
    save_btn.pack(side=tk.LEFT, padx=10)  # 간격 증가

    # 취소 버튼
    cancel_btn = tk.Button(btn_frame, text=get_msg("cancel"), command=cancel,
                          width=12, height=1, bg="#f44336", fg="white", font=("Segoe UI", 10, "bold"))  # 크기 증가
    cancel_btn.pack(side=tk.LEFT, padx=10)  # 간격 증가

    # 이벤트 바인딩 (창에 직접 바인딩)
    dialog.bind("<Return>", lambda event: save_key())
    dialog.bind("<Escape>", lambda event: cancel())

    # 창 닫기 처리
    dialog.protocol("WM_DELETE_WINDOW", cancel)

    # 포커스 설정 - 다양한 방법으로 포커스 강제 설정
    dialog.after(100, lambda: entry.focus_set())
    dialog.after(200, lambda: entry.focus_force())
    dialog.after(300, lambda: entry.icursor(tk.END))  # 커서를 텍스트 끝으로

    entry.select_range(0, tk.END)  # 전체 텍스트 선택

    # 창 위치 설정 (화면 중앙)
    dialog.update_idletasks()
    width = dialog.winfo_width()
    height = dialog.winfo_height()
    x = (dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (dialog.winfo_screenheight() // 2) - (height // 2)
    dialog.geometry(f"+{x}+{y}")

    # 추가적인 포커스 처리
    dialog.after(200, lambda: entry.focus_force())

    # 모달 대화 상자 실행 (기다림)
    dialog.wait_window()

    return result[0] or api_key

# API 키 오류 대화 상자
def show_api_key_error_dialog(error_message):
    """API 키 오류 대화 상자를 표시합니다."""
    dialog = tk.Toplevel()
    dialog.title(get_msg("error"))
    dialog.geometry("450x200")
    dialog.resizable(False, False)

    # 모달 설정
    dialog.transient()
    dialog.grab_set()
    dialog.focus_set()
    dialog.attributes("-topmost", True)

    # 메인 프레임
    frame = tk.Frame(dialog, padx=20, pady=15)
    frame.pack(fill=tk.BOTH, expand=True)

    # 오류 아이콘
    try:
        error_label = tk.Label(frame, text="⚠️", font=("Segoe UI", 24), fg="red")
        error_label.pack(pady=(0, 10))
    except:
        pass  # 이모지 표시 오류가 나면 무시

    # 오류 메시지
    message_label = tk.Label(frame, text=error_message, font=("Segoe UI", 10), wraplength=400)
    message_label.pack(pady=10)

    # 버튼 프레임
    btn_frame = tk.Frame(frame)
    btn_frame.pack(pady=10)

    # API 키 설정 버튼
    def open_api_key_dialog():
        dialog.destroy()
        show_api_key_dialog(required=False)

    setup_btn = tk.Button(btn_frame, text=get_msg("set_openai_api_key"), command=open_api_key_dialog,
                         width=15, bg="#4CAF50", fg="white", font=("Segoe UI", 9, "bold"))
    setup_btn.pack(side=tk.LEFT, padx=10)

    # 닫기 버튼
    close_btn = tk.Button(btn_frame, text=get_msg("cancel"), command=dialog.destroy,
                         width=10, bg="#f44336", fg="white", font=("Segoe UI", 9, "bold"))
    close_btn.pack(side=tk.LEFT, padx=10)

    # 키보드 이벤트
    dialog.bind("<Return>", lambda event: open_api_key_dialog())
    dialog.bind("<Escape>", lambda event: dialog.destroy())

    # 창 위치 설정 (화면 중앙)
    dialog.update_idletasks()
    width = dialog.winfo_width()
    height = dialog.winfo_height()
    x = (dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (dialog.winfo_screenheight() // 2) - (height // 2)
    dialog.geometry(f"+{x}+{y}")

    # 대화 상자 표시
    dialog.wait_window()

# 단축키 설정 대화상자 함수 추가 (API 키 대화 상자 다음에 추가)
def show_hotkey_dialog():
    """단축키 변경 대화 상자를 표시합니다."""
    global hotkey_modifiers, hotkey_key, Listener, Key, KeyCode

    # 현재 단축키 문자열 생성
    def get_hotkey_string():
        parts = []
        if hotkey_modifiers.get("ctrl", False):
            parts.append("Ctrl")
        if hotkey_modifiers.get("shift", False):
            parts.append("Shift")
        if hotkey_modifiers.get("alt", False):
            parts.append("Alt")
        if hotkey_key:
            parts.append(str(hotkey_key).upper())
        return "+".join(parts)

    # 완전히 독립적인 모달 대화 상자 생성
    dialog = tk.Toplevel()
    dialog.title(get_msg("set_hotkey_title", "단축키 설정"))
    dialog.geometry("450x300")
    dialog.resizable(False, False)

    # 모달 설정 (부모 창 비활성화)
    dialog.transient()
    dialog.grab_set()
    dialog.focus_set()

    # 항상 위에 표시
    dialog.attributes("-topmost", True)

    # 메인 프레임
    frame = tk.Frame(dialog, padx=25, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)

    # 설명 라벨
    tk.Label(frame, text=get_msg("hotkey_instruction", "단축키를 설정하세요. 체크박스로 수정자 키를 선택하세요."),
             font=("Segoe UI", 11), justify=tk.LEFT).pack(anchor="w", pady=(0, 15))

    # 수정자 키 체크박스
    modifiers_frame = tk.Frame(frame)
    modifiers_frame.pack(fill=tk.X, pady=5)

    # Ctrl 체크박스
    ctrl_var = tk.BooleanVar(value=hotkey_modifiers.get("ctrl", True))
    ctrl_cb = tk.Checkbutton(modifiers_frame, text="Ctrl", variable=ctrl_var, font=("Segoe UI", 10))
    ctrl_cb.pack(side=tk.LEFT, padx=10)

    # Shift 체크박스
    shift_var = tk.BooleanVar(value=hotkey_modifiers.get("shift", True))
    shift_cb = tk.Checkbutton(modifiers_frame, text="Shift", variable=shift_var, font=("Segoe UI", 10))
    shift_cb.pack(side=tk.LEFT, padx=10)

    # Alt 체크박스
    alt_var = tk.BooleanVar(value=hotkey_modifiers.get("alt", True))
    alt_cb = tk.Checkbutton(modifiers_frame, text="Alt", variable=alt_var, font=("Segoe UI", 10))
    alt_cb.pack(side=tk.LEFT, padx=10)

    # 키 추가 프레임
    key_frame = tk.Frame(frame)
    key_frame.pack(fill=tk.X, pady=10)

    # 추가 키 라벨
    tk.Label(key_frame, text=get_msg("additional_key", "추가 키(선택사항):"),
             font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(0, 10))

    # 추가 키 입력 필드 (읽기 전용)
    key_var = tk.StringVar(value=str(hotkey_key) if hotkey_key else "")
    key_entry = tk.Entry(key_frame, textvariable=key_var, width=10, font=("Segoe UI", 10),
                         state="readonly", bg="white")
    key_entry.pack(side=tk.LEFT)

    # 키 리스닝 상태
    listening = [False]

    # 리스닝 시작 버튼
    def start_listening():
        if listening[0]:
            return

        listening[0] = True
        listen_btn.config(text=get_msg("press_key", "키를 누르세요..."), bg="red")
        key_var.set("")

        # 입력된 키 값을 저장할 변수
        key_value = None

        # 키 눌림 이벤트 핸들러
        def on_key_press(key):
            if not listening[0]:
                return True

            try:
                nonlocal key_value
                # 수정자 키는 무시
                if key == Key.ctrl_l or key == Key.ctrl_r or \
                    key == Key.shift_l or key == Key.shift_r or \
                    key == Key.alt_l or key == Key.alt_r:
                    return True

                # 다른 키는 저장
                if isinstance(key, KeyCode):
                    key_name = key.char
                    if key_name:
                        key_var.set(key_name.upper())
                        key_value = key_name.upper()
                else:
                    # 특수 키는 이름 사용
                    key_name = str(key).replace('Key.', '')
                    key_var.set(key_name.upper())
                    key_value = key_name.upper()

                # 리스닝 종료
                listening[0] = False
                listen_btn.config(text=get_msg("set_key", "키 설정"), bg="#4CAF50")

                # 키 입력이 끝나면 리스너 중지
                return False
            except Exception as e:
                logging.error(f"키 리스닝 오류: {str(e)}")
                return False

        # 임시 리스너 설정
        temp_listener = Listener(on_press=on_key_press)
        temp_listener.start()

        # 5초 후 자동 타임아웃
        def timeout():
            nonlocal key_value
            if listening[0]:
                listening[0] = False
                listen_btn.config(text=get_msg("set_key", "키 설정"), bg="#4CAF50")
                temp_listener.stop()

            # 키 값이 입력되었다면 저장
            if key_value:
                key_var.set(key_value)

        dialog.after(5000, timeout)

    # 리스닝 버튼
    listen_btn = tk.Button(key_frame, text=get_msg("set_key", "키 설정"),
                         command=start_listening, bg="#4CAF50", fg="white",
                         font=("Segoe UI", 9))
    listen_btn.pack(side=tk.LEFT, padx=10)

    # 키 지우기 버튼
    def clear_key():
        key_var.set("")

    clear_btn = tk.Button(key_frame, text=get_msg("clear_key", "지우기"),
                         command=clear_key, bg="#f44336", fg="white",
                         font=("Segoe UI", 9))
    clear_btn.pack(side=tk.LEFT)

    # 현재 단축키 표시
    current_hotkey_frame = tk.Frame(frame)
    current_hotkey_frame.pack(fill=tk.X, pady=15)

    tk.Label(current_hotkey_frame, text=get_msg("current_hotkey", "현재 단축키:"),
             font=("Segoe UI", 10)).pack(side=tk.LEFT)

    current_hotkey_label = tk.Label(current_hotkey_frame, text=get_hotkey_string(),
                                   font=("Segoe UI", 10, "bold"))
    current_hotkey_label.pack(side=tk.LEFT, padx=10)

    # 결과 메시지 라벨
    result_label = tk.Label(frame, text="", font=("Segoe UI", 10))
    result_label.pack(pady=10)

    # 버튼 프레임
    btn_frame = tk.Frame(frame)
    btn_frame.pack(pady=10)

    # 저장 함수
    def save_hotkey():
        global hotkey_modifiers, hotkey_key

        # 최소한 하나의 수정자 키 필요
        if not (ctrl_var.get() or shift_var.get() or alt_var.get()):
            result_label.config(text=get_msg("need_modifier", "최소 하나의 수정자 키(Ctrl, Shift, Alt)가 필요합니다."),
                               fg="red")
            return

        # 수정자 키 저장
        hotkey_modifiers = {
            "ctrl": ctrl_var.get(),
            "shift": shift_var.get(),
            "alt": alt_var.get()
        }

        # 추가 키 저장
        key_text = key_var.get().strip()
        hotkey_key = key_text if key_text else None

        # 저장된 값 로깅
        logging.info(f"저장되는 단축키 값: 수정자={hotkey_modifiers}, 키={hotkey_key}")
        print(f"=== 단축키 저장 시작 ===")
        print(f"저장할 단축키: 수정자={hotkey_modifiers}, 키={hotkey_key}")

        # 설정 파일에 저장
        try:
            save_settings()
            print("설정 파일에 단축키 저장 완료")
            logging.info("단축키 설정 저장 완료")
        except Exception as e:
            err_msg = f"설정 파일 저장 오류: {str(e)}"
            print(err_msg)
            logging.error(err_msg)
            result_label.config(text=err_msg, fg="red")
            return

        # 성공 메시지
        result_label.config(text=get_msg("hotkey_saved", "단축키가 저장되었습니다."), fg="green")

        # 현재 단축키 표시 업데이트
        current_hotkey_label.config(text=get_hotkey_string())

        # 단축키 변경 로그 출력
        hotkey_str = get_hotkey_string()
        log_msg = f"단축키가 변경되었습니다: {hotkey_str}"
        logging.info(log_msg)
        log_to_console(log_msg)
        print(log_msg)

        # 키보드 리스너 재설정 - 새로운 단축키 적용을 위해
        print("키보드 리스너 재설정 중...")
        setup_keyboard_listener()
        print("키보드 리스너 재설정 완료")

        # 1초 후 창 닫기
        dialog.after(1000, dialog.destroy)

    # 취소 함수
    def cancel():
        dialog.destroy()

    # 저장 버튼
    save_btn = tk.Button(btn_frame, text=get_msg("save", "저장"), command=save_hotkey,
                         width=12, height=1, bg="#4CAF50", fg="white", font=("Segoe UI", 10, "bold"))
    save_btn.pack(side=tk.LEFT, padx=10)

    # 취소 버튼
    cancel_btn = tk.Button(btn_frame, text=get_msg("cancel", "취소"), command=cancel,
                          width=12, height=1, bg="#f44336", fg="white", font=("Segoe UI", 10, "bold"))
    cancel_btn.pack(side=tk.LEFT, padx=10)

    # 이벤트 바인딩
    dialog.bind("<Escape>", lambda event: cancel())

    # 창 위치 설정 (화면 중앙)
    dialog.update_idletasks()
    width = dialog.winfo_width()
    height = dialog.winfo_height()
    x = (dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (dialog.winfo_screenheight() // 2) - (height // 2)
    dialog.geometry(f"+{x}+{y}")

    # 대화 상자 표시
    dialog.wait_window()

# 녹음 관련 함수
def start_recording():
    """녹음 시작 함수"""
    global recording, audio_data, sd, np, stream, selected_device, winsound # winsound 전역 변수 사용 명시

    if recording:
        logging.info("이미 녹음 중입니다.")
        return

    # sounddevice 모듈이 로드되었는지 확인
    if sd is None or np is None:
        error_msg = "녹음에 필요한 모듈이 로드되지 않았습니다."
        logging.error(error_msg)
        log_to_console(f"오류: {error_msg}")
        # GUI 오류 메시지 표시 (메인 스레드에서)
        if root:
            root.after(10, lambda: messagebox.showerror("녹음 오류", error_msg))
        return

    try:
        # 녹음 시작
        logging.info("녹음 시작")
        log_to_console("녹음 시작 중...")

        # 마이크 정보 확인
        try:
            # 사용 가능한 장치 확인
            devices = sd.query_devices()
            logging.info(f"사용 가능한 오디오 장치: {len(devices)}개")

            # 입력 장치 선택 (선택된 장치가 없으면 기본 장치 사용)
            device_info = None
            if selected_device is not None:
                device_info = sd.query_devices(selected_device)
                logging.info(f"선택된 마이크 사용: {device_info['name']}")
            else:
                device_info = sd.query_devices(kind='input')
                logging.info(f"기본 마이크 사용: {device_info['name']}")

            # 로그에 디바이스 정보 기록
            log_to_console(f"마이크: {device_info['name']}")
        except Exception as e:
            logging.error(f"마이크 정보 확인 오류: {str(e)}")
            log_to_console(f"마이크 정보 확인 오류: {str(e)}")
            # 계속 진행 (기본 설정으로 시도)

        # 샘플링 설정
        samplerate = 16000  # OpenAI Whisper에 적합한 샘플링 레이트
        channels = 1        # 모노 녹음

        # 오디오 데이터 초기화
        audio_data = []
        recording = True

        # 녹음 콜백 함수
        def audio_callback(indata, frames, time, status):
            if status:
                logging.warning(f"녹음 상태 문제: {status}")
            if recording:
                try:
                    if indata.shape[1] == channels:  # 채널 수 확인
                        audio_data.append(indata.copy())
                    else:
                        # 채널 수가 맞지 않으면 로그만 남기고 계속 진행
                        logging.warning(f"채널 수 불일치: 예상 {channels}, 실제 {indata.shape[1]}")
                except Exception as cb_e:
                    logging.error(f"오디오 콜백 오류: {str(cb_e)}")

        # 스트림 시작
        try:
            device_id = selected_device if selected_device is not None else None
            stream = sd.InputStream(
                device=device_id,
                samplerate=samplerate,
                channels=channels,
                callback=audio_callback
            )
            stream.start()
            logging.info("오디오 스트림 시작됨")

            # 녹음 시작 비프음 추가 (Windows 환경)
            if winsound:
                try:
                    winsound.Beep(600, 200) # 600Hz, 200ms 비프음
                except Exception as beep_e:
                    logging.warning(f"녹음 시작 비프음 재생 오류: {str(beep_e)}")

        except Exception as stream_e:
            error_msg = f"오디오 스트림 시작 오류: {str(stream_e)}"
            logging.error(error_msg)
            log_to_console(error_msg)
            recording = False
            # GUI 오류 메시지 표시 (메인 스레드에서)
            if root:
                root.after(10, lambda: messagebox.showerror("녹음 오류", f"마이크를 시작할 수 없습니다: {str(stream_e)}"))
            return

        # 녹음 시작 알림
        log_to_console("녹음 중... (단축키를 떼면 종료)") # 메시지 수정

    except Exception as e:
        error_msg = f"녹음 시작 오류: {str(e)}"
        logging.error(error_msg)
        log_to_console(error_msg)
        recording = False
        # GUI 오류 메시지 표시 (메인 스레드에서)
        if root:
            root.after(10, lambda: messagebox.showerror("녹음 오류", error_msg))

def stop_recording():
    """녹음 중지 및 오디오 처리 함수"""
    global recording, audio_data, stream, openai, api_key, pyperclip

    if not recording:
        logging.info("녹음 중이 아닙니다.")
        return

    try:
        # 녹음 중지
        recording = False
        log_to_console("녹음 종료 중...")

        # 스트림 종료
        if stream:
            stream.stop()
            stream.close()

        # 소리로 녹음 종료 알림 (Windows 환경)
        if winsound:
            try:
                winsound.Beep(800, 200)  # 800Hz, 200ms
            except:
                pass

        # 녹음된 데이터가 없으면 종료
        if not audio_data:
            logging.warning("녹음된 데이터가 없습니다.")
            log_to_console(get_msg("no_audio_data"))
            return

        log_to_console(get_msg("processing_audio"))

        # 오디오 데이터 합치기
        try:
            audio = np.concatenate(audio_data, axis=0)

            # 녹음 폴더 확인
            recordings_dir = "recordings"
            if not os.path.exists(recordings_dir):
                os.makedirs(recordings_dir)

            # 현재 시간을 파일명으로 사용
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.join(recordings_dir, f"recording_{timestamp}.flac")

            # FLAC 파일로 저장 (Whisper API에 최적)
            log_to_console(get_msg("saving_audio_file", filename))
            soundfile.write(filename, audio, 16000)

            # API 키가 없는 경우 즉시 API 키 설정 창 표시
            if not api_key:
                log_to_console(get_msg("no_api_key_set"))
                # 메인 스레드에서 API 키 설정 창 표시
                if root:
                    root.after(100, lambda: show_api_key_dialog(required=True))
                return

            # Whisper API로 음성 인식
            if openai and api_key:
                log_to_console(get_msg("sending_to_whisper"))
                log_to_console(get_msg("auto_language_detection"))

                try:
                    # 파일을 FLAC 형식으로 Whisper API에 전송
                    openai.api_key = api_key

                    # API 호출 시간 측정 시작
                    api_start_time = time.time()

                    # 언어 설정에 따른 파라미터 처리
                    api_params = {
                        "model": "whisper-1",
                    }
                    
                    # 자동 감지 사용 여부
                    if auto_language_detection:
                        # 자동 감지 사용 - language 파라미터 제외
                        log_to_console(get_msg("auto_language_detection"))
                    else:
                        # 트레이 아이콘 언어 설정에 따라 language 파라미터 추가
                        if current_language == "ko":
                            api_params["language"] = "ko"
                            log_to_console(get_msg("using_language", "한국어"))
                        elif current_language == "en":
                            api_params["language"] = "en"
                            log_to_console(get_msg("using_language", "English"))
                        else:
                            # 다른 언어나 미정의 경우 자동 감지 사용 (현재와 동일)
                            log_to_console(get_msg("auto_language_detection"))
                    
                    with open(filename, "rb") as audio_file:
                        # API 호출 (언어 파라미터 적용)
                        api_params["file"] = audio_file
                        transcript = openai.Audio.transcribe(**api_params)

                    # API 호출 시간 측정 종료 및 레이턴시 계산
                    api_end_time = time.time()
                    api_latency = (api_end_time - api_start_time) * 1000  # 초 단위를 밀리초 단위로 변환

                    text = transcript["text"]

                    # 텍스트 처리
                    if text:
                        # 원본 텍스트 로깅 (이전과 동일)
                        log_to_console("====== " + get_msg("recognition_result") + " ======")
                        log_to_console(get_msg("api_response_time", api_latency))
                        log_to_console(get_msg("recognized_text", text))
                        log_to_console(get_msg("text_length", len(text)))
                        log_to_console("=========================")
                        logging.info(f"[인식결과] 원본 텍스트: {text}")

                        # 1. "개행" 또는 "엔터" 및 주변 공백을 줄바꿈 문자로 변환 (정규식 사용)
                        original_text_before_newline = text
                        text = re.sub(r'\s*(개행|엔터)\s*', '\\n', text) # Regex replace

                        # 변환 로그 (이전과 유사)
                        if text != original_text_before_newline:
                            log_to_console("특정 단어('개행', '엔터') 및 주변 공백을 줄바꿈 문자로 변환했습니다.")
                            logging.info(f"[텍스트 변환] 완료. 변환된 텍스트: {text}")
                        else:
                            logging.info("[텍스트 변환] 변환할 단어('개행', '엔터') 없음")

                        # 2. 전체 텍스트 앞뒤 공백/개행 제거 및 시작/끝 마침표 제거
                        cleaned_text = text.strip()
                        initial_cleaned_text = cleaned_text # 비교용
                        if cleaned_text.startswith('.'):
                            cleaned_text = cleaned_text[1:]
                        if cleaned_text.endswith('.'):
                            cleaned_text = cleaned_text[:-1]
                        cleaned_text = cleaned_text.strip() # 마침표 제거 후 남을 수 있는 공백 재제거

                        # 3. 줄 단위 추가 정리: 각 줄 시작의 '.' 및 공백 제거
                        lines = cleaned_text.split('\\n')
                        cleaned_lines = []
                        text_changed_in_loop = False # 줄 단위 변경 추적

                        for line in lines:
                            processed_line = line.strip() # 각 줄의 앞뒤 공백 제거
                            original_processed_line = processed_line
                            if processed_line.startswith('.'):
                                # 맨 앞 '.' 제거 및 뒤따르는 공백 제거
                                processed_line = processed_line[1:].lstrip()

                            cleaned_lines.append(processed_line)
                            if processed_line != original_processed_line:
                                text_changed_in_loop = True

                        # 4. 최종 텍스트 재구성
                        final_cleaned_text = "\\n".join(cleaned_lines)

                        # 최종 변경 로그 및 결과 할당
                        if final_cleaned_text != initial_cleaned_text or text_changed_in_loop:
                             log_to_console("텍스트 앞/뒤 및 각 줄 시작의 불필요한 공백/마침표를 최종 정리했습니다.")
                             log_to_console(f"최종 정리된 텍스트: {final_cleaned_text}")
                             logging.info(f"[텍스트 최종 정리] 완료. 최종 텍스트: {final_cleaned_text}")
                             text = final_cleaned_text # 최종 결과 사용
                        elif cleaned_text != text.strip(): # Regex 변환 후 첫 strip에서만 변경된 경우
                             log_to_console("텍스트 앞/뒤 공백/개행을 제거했습니다.")
                             log_to_console(f"정리된 텍스트: {cleaned_text}")
                             logging.info(f"[텍스트 정리] 앞/뒤 공백 제거 완료. 텍스트: {cleaned_text}")
                             text = cleaned_text # 중간 결과 사용
                        # else: 로그 불필요

                        # 클립보드/붙여넣기 (최종 text 사용)
                        clipboard_success = False
                        if pyperclip:
                            try:
                                pyperclip.copy(text) # 최종 정리된 text 사용
                                clipboard_success = True
                                log_to_console(get_msg("text_copied"))
                            except Exception as clip_e:
                                logging.error(f"클립보드 복사 오류: {str(clip_e)}")
                                log_to_console(get_msg("copy_error", str(clip_e)))

                        # 붙여넣기 방법 1: 컨트롤러로 붙여넣기
                        paste_success = False
                        if Controller:
                            try:
                                time.sleep(0.2)
                                keyboard = Controller()
                                log_to_console(get_msg("attempting_paste"))

                                keyboard.press(Key.ctrl)
                                keyboard.press('v')
                                keyboard.release('v')
                                keyboard.release(Key.ctrl)
                                paste_success = True
                                log_to_console(get_msg("paste_complete"))

                            except Exception as paste_e:
                                logging.error(f"붙여넣기 오류: {str(paste_e)}")
                                log_to_console(get_msg("paste_error", str(paste_e)))

                                # 붙여넣기 실패 시 대체 방법으로 직접 텍스트 입력 시도
                                try:
                                    log_to_console(get_msg("attempting_direct_input"))
                                    time.sleep(0.2)

                                    # 각 문자를 하나씩 입력 시도
                                    for char in text: # 최종 정리된 text 사용
                                        try:
                                            keyboard.type(char)
                                            time.sleep(0.01)  # 약간의 지연
                                        except:
                                            pass

                                    log_to_console(get_msg("direct_input_complete"))
                                except Exception as direct_e:
                                    logging.error(f"직접 텍스트 입력 오류: {str(direct_e)}")
                                    log_to_console(get_msg("direct_input_error", str(direct_e)))

                        # 최종 상태 보고
                        if clipboard_success:
                            log_to_console(get_msg("clipboard_success"))

                        if not paste_success and not clipboard_success:
                            log_to_console(get_msg("all_input_failed"))
                            log_to_console(get_msg("manual_copy"))
                            log_to_console(text)
                    else:
                        log_to_console(get_msg("no_text_recognized"))

                except Exception as e:
                    logging.error(f"Whisper API 오류: {str(e)}")
                    log_to_console(get_msg("recognition_error", str(e)))

                    # API 키 오류인지 확인
                    error_msg = str(e).lower()
                    if "api key" in error_msg or "apikey" in error_msg or "authentication" in error_msg or "인증" in error_msg:
                        error_message = get_msg("api_key_invalid_long")
                        log_to_console(error_message)

                        # 바로 API 키 설정 창 표시 (메인 스레드에서 실행)
                        if root:
                            def show_error_and_settings():
                                messagebox.showerror(get_msg("api_key_error"), error_message)
                                show_api_key_dialog(required=True)

                            root.after(100, show_error_and_settings)

            else:
                log_to_console(get_msg("openai_api_not_set"))
                # API 키가 설정되지 않은 경우 즉시 API 키 설정 창 표시
                if not api_key and root:
                    root.after(100, lambda: show_api_key_dialog(required=True))

        except Exception as e:
            logging.error(f"오디오 처리 오류: {str(e)}")
            log_to_console(get_msg("audio_processing_error", str(e)))

    except Exception as e:
        logging.error(f"녹음 종료 오류: {str(e)}")
        log_to_console(get_msg("recording_stop_error", str(e)))

    finally:
        # 상태 초기화
        recording = False
        audio_data = []
        stream = None

def extract_readme_files():
    """README 파일을 실행 파일이 있는 디렉토리에 추출합니다."""
    try:
        import os
        import sys

        # 실행 파일 경로 찾기
        if getattr(sys, 'frozen', False):
            # PyInstaller로 패키징된 경우
            base_path = sys._MEIPASS
            exe_dir = os.path.dirname(sys.executable)

            # README 파일 경로
            readme_files = ["README.md", "README.KR.md"]

            for readme_file in readme_files:
                src_path = os.path.join(base_path, readme_file)
                dst_path = os.path.join(exe_dir, readme_file)

                # 파일이 존재하고 대상 경로에 없는 경우에만 복사
                if os.path.exists(src_path) and not os.path.exists(dst_path):
                    import shutil
                    shutil.copy2(src_path, dst_path)
                    print(f"Extracted {readme_file} to {exe_dir}")
    except Exception as e:
        print(f"Error extracting README files: {e}")

# 메인 함수 실행
if __name__ == "__main__":
    try:
        # 프로그램 다중 실행 체크 - 실행 중이면 종료
        if not prevent_multiple_instances():
            sys.exit(0)

        main()
        # 메인 스레드가 여기에 도달하면 루트 윈도우 메인 루프 실행
        if root:
            root.mainloop()
    except KeyboardInterrupt:
        print("\n사용자에 의해 프로그램이 종료되었습니다.")
    except Exception as e:
        logging.error(f"예상치 못한 오류로 프로그램이 종료됩니다: {str(e)}")
        print(f"오류: {str(e)}")
    finally:
        # 종료 시 clean-up
        if tray_icon:
            try:
                tray_icon.stop()
            except:
                pass
        print("프로그램이 종료됩니다.")
        sys.exit(0)
