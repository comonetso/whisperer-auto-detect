**This application was developed on March 17, 2025, by developers at LMTY Yeogiaen Service Co., Ltd. to make Cursor AI more convenient to use.**

# Yeogiaen WhisperTyper

A lightweight desktop application that converts speech to text in real-time using OpenAI's Whisper API. Perfect for quick voice memos, dictation, and accessibility needs.

## Key Features

- **Easy Voice Recording**: Press default shortcut (Ctrl+Shift+Alt) to start recording, release any key to stop
- **Automatic Text Conversion**: Accurate speech recognition using OpenAI's Whisper API
- **Automatic Language Detection**: Automatically detects and converts speech language
- **Clipboard Integration**: Automatically copies and pastes converted text
- **Multi-language Support**: Korean and English interface available
- **Customizable Shortcuts**: Change shortcuts as needed (excluding Windows system reserved shortcuts)
- **System Tray Integration**: Runs in the background with system tray icon access
- **Automatic Recording Save**: All recordings are saved with timestamps in FLAC format (16kHz)

## Requirements

- Windows Operating System
- OpenAI API Key (Sign up: [OpenAI Platform](https://platform.openai.com/))
- Internet connection for API access
- Microphone device (default or selectable)

## Installation

1. Download the latest version from the releases page
2. Extract the ZIP file to your desired location
3. Run `whisperer.exe` (duplicate execution is automatically prevented)
4. Enter your OpenAI API key when prompted (only required on first run)

## How to Use

1. The application runs in the background with a system tray icon
2. Press and hold the set shortcut (default: Ctrl+Shift+Alt) to start recording
3. Speak clearly into your microphone (language is automatically detected)
4. Release any key to stop recording and automatically convert to text
5. The converted text is automatically copied to clipboard and pasted at the current cursor position
6. Recorded files are automatically saved in the 'recordings' folder

## System Tray Options

Right-click the system tray icon to access the following options:

- **Open Recordings Folder**: Opens the folder containing all recorded audio files
- **Open README**: Opens help documentation in the current language
- **Open Console Window**: Opens console window for debugging and log viewing
- **Set OpenAI API Key**: Change your API key
- **Set Shortcuts**: Modify recording start/stop shortcuts
- **Change Language**: Switch between Korean and English interface
- **Exit**: Close the application

## Notes

- The application requires an internet connection to use the OpenAI API
- Each conversion using the Whisper API may incur costs based on API usage
- Recordings are saved in FLAC format in the `recordings` folder with timestamped filenames
- OpenAI API key is securely stored in 'openai_api_key.txt'
- The program automatically prevents duplicate execution
- All logs are stored in the 'logs' folder to help with troubleshooting

## Troubleshooting

- Program won't start: Check if an instance is already running
- Microphone issues: Re-check and initialize microphone from the system tray menu
- Automatic paste not working: Try pressing Ctrl+V manually
- Slow or failed conversion: Check your internet connection
- API key errors: Try resetting the API key from the system tray menu
- Missing tray icon: Verify favicon.ico file is in the same folder as the program
- Corrupted text: Verify UTF-8 encoding is properly set

## Credits

This application uses:
- OpenAI's Whisper API (Speech Recognition)
- Python Libraries:
  - sounddevice (Audio Recording)
  - soundfile (Audio File Processing)
  - pynput (Keyboard Event Handling)
  - pystray (System Tray Icon)
  - tkinter (GUI Elements)
  - Other standard Python libraries

## License

This project is provided under the MIT License.
