# Windows Clipboard Manager

A modern, lightweight clipboard manager for Windows with image support, built using AutoHotkey v2.

![Main Interface](Screenshots/Screenshot%202024-01-04%20221926.png)

## Features

- Multi-item clipboard history (up to 25 items)
- Support for both text and images
- Automatic text logging with timestamps
- Automatic image saving to Pictures folder
- Single-click paste functionality
- Modern UI with hover effects
- Windows + V hotkey activation

## Screenshots

<details>
<summary>Click to view all screenshots</summary>

### Main Interface
![Main Interface](Screenshots/Screenshot%202024-01-04%20221926.png)

### Text and Image Support
![Text Support](Screenshots/Screenshot%202024-01-04%20222007.png)
![Image Support](Screenshots/Screenshot%202024-01-04%20222024.png)

</details>

## Requirements

- Windows 10/11
- AutoHotkey v2.0 (for running source code)

## Installation

### Option 1: Run Compiled Version
1. Download the latest release from the [Releases](../../releases) page
2. Run `clipboard_manager.exe`
3. Press Win+V to access your clipboard history

### Option 2: Run Source Code
1. Install [AutoHotkey v2](https://www.autohotkey.com/)
2. Clone this repository
3. Run `clipboard_manager.ahk`

## Usage

1. Copy any text or image (Ctrl+C)
2. Press Win+V to open the clipboard manager
3. Click any item to paste it
4. Text entries are automatically logged to `Documents\ClipboardLog.txt`
5. Images are saved to `Pictures\ClipboardImages`

## Features

- **Text Handling**
  - Stores formatted and plain text
  - Automatic deduplication
  - Persistent logging with timestamps

- **Image Support**
  - Saves images in PNG format
  - Automatic thumbnail generation
  - Permanent storage in Pictures folder

- **User Interface**
  - Clean, modern design
  - Grid layout with hover effects
  - Status messages for operations
  - Single-click paste functionality

## Development

This project is built with AutoHotkey v2 and uses:
- Windows API for clipboard operations
- GDI+ for image processing
- Shell API for system folder access

### Building from Source

1. Install AutoHotkey v2
2. Clone the repository:git clone https://github.com/royal-crisis/clipboard-manager.git
3. Compile using Ahk2Exe:
- Right-click `clipboard_manager.ahk`
- Select "Compile Script"
- Choose AutoHotkey64.exe as base file

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

Distributed under the MIT License. See `LICENSE` for more information.

## Acknowledgments

- AutoHotkey community for documentation and support
- Windows API documentation
- Contributors and users of this project

## Support

If you encounter any issues or have suggestions:
1. Check the [Issues](../../issues) page
2. Create a new issue if needed
3. Join our [Discussions](../../discussions) for questions and ideas

---
Made with ❤️ by [royal-crisis](https://github.com/royal-crisis)
