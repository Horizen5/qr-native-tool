# QR Code Fragment Scan & Reconstruct Tool

Languages: [简体中文](../README.md) · [English](./README_en.md)

A cross-device file transfer tool powered by **Word VBS + Android native APP**. The core idea is: convert any file into Base64, split it into multiple QR code images using Word's "text to QR code" feature, then scan the codes continuously with a phone and reconstruct the original file.

## Project Origin

This project was born from a very specific real-world scenario:

> In a training institution's computer lab, I needed to export a semiconductor circuit design PDK (Process Design Kit) and tech library. However, the lab computers were heavily restricted—connecting a phone via USB only allowed charging, not file transfer; USB flash drives were also disabled. I tried entering BIOS to change the boot order and boot into a WinPE system, but the BIOS was password-locked.

So this tool came into being: leveraging the pre-installed **Office 2016** on the lab computers, I used Word's **DISPLAYBARCODE field** to encode files into QR codes, then scanned them continuously with a native Android APP, effectively "moving" the tech library files out of the restricted lab environment and into a normal Windows system outside the virtual machine.

## ✨ Features

### VBS QR Code Generator

- ✅ **One-click generation inside Word**: Generate QR code images directly in a Word document, no extra software needed
- ✅ **Base64 fragmentation**: Automatically splits files into data chunks suitable for QR code capacity
- ✅ **Unified checksum**: Every QR code page contains the same checksum header to prevent mixing in unrelated codes
- ✅ **Original filename restoration**: Filenames are transmitted via Base64 encoding, and the APP restores the original name and extension
- ✅ **UTF-8 Chinese support**: Chinese filenames are Base64-encoded to avoid Word field issues with non-ASCII characters
- ✅ **Auto-pagination**: Automatically calculates the number of pages; page 1 carries the filename, subsequent pages carry only data

### Android Native APP

- ✅ **Native development**: Built with Kotlin + CameraX, not a WebView wrapper
- ✅ **Continuous scanning**: Automatically recognizes fragment indexes, supports out-of-order scanning and rescanning missing pages
- ✅ **Checksum filtering**: Only accepts QR codes matching the current task's checksum
- ✅ **Draft box resume**: Pressing back saves current progress to a draft box, allowing you to continue scanning later
- ✅ **Original filename saving**: Automatically restores the filename and extension passed from the VBS side
- ✅ **Save to Download/QR**: Files are saved to the `QR` subfolder of the phone's Downloads directory
- ✅ **WeChat-style path toast**: After saving, shows `文件已保存到 /sdcard/Download/QR/filename`
- ✅ **File size display**: Result page shows reconstructed file size; values under 1KB are shown in bytes
- ✅ **Multi-camera switching**: Supports switching between different rear cameras
- ✅ **Large QR code support**: Optimized memory and parsing logic to handle larger QR codes

## 📦 Files

| File | Description |
|---|---|
| `QR合成-原生-v7.apk` | Android native APP (latest v7) |
| `QR生成-v6.vbs` | Word VBS QR code generation script |

## 🚀 How to Use

### 1. Generate QR Codes on PC

1. Open Word 2016 (or any version supporting the `DISPLAYBARCODE` field)
2. Press `Alt + F11` to open the VBA editor
3. Insert a module and paste the code from `QR生成-v6.vbs`
4. Run the `QR` macro and select the file you want to transfer
5. Word will automatically generate a multi-page QR code document
6. Select all, copy or export as images, and display the QR codes on the restricted computer

### 2. Scan and Reconstruct on Phone

1. Install `QR合成-原生-v7.apk` on your Android phone
2. Open the APP and point the camera at the QR codes on the screen
3. The APP automatically recognizes the checksum, index, and total count
4. After scanning is complete, tap "保存文件" (Save File)
5. The file will be saved to `/sdcard/Download/QR/`

## 🛠 Technical Details

### QR Code Protocol Format

```
TZ|checksum|index|total|fileNameB64|data
```

| Field | Description |
|---|---|
| `TZ` | Fixed protocol header |
| `checksum` | Checksum composed of the first 6 chars of filename Base64 + total Base64 data length |
| `index` | Current page index (starting from 1) |
| `total` | Total number of pages |
| `fileNameB64` | Base64-encoded filename, only carried on page 1 |
| `data` | Base64 data chunk for the current page |

### Fragmentation Strategy

- Subsequent pages carry a fixed 1000 characters of data
- Page 1 data length is dynamically calculated and aligned to a multiple of 4, ensuring the overall Base64 string is decodable
- Filenames are transmitted via Base64, supporting Chinese and special characters

### Android Saving Mechanism

- Android 10+ uses MediaStore to write to `Download/QR`
- Returns the real file path `/storage/emulated/0/Download/QR/filename`
- Toast display replaces it with `/sdcard/Download/QR/filename`
- File extension is identified via file magic numbers (JPEG/PNG/GIF/PDF/ZIP/BMP, etc.)

## 📋 Version History

### v7 (2026-07-15)

- Fixed file size display: shows bytes when under 1KB instead of 0 KB
- Fixed save path: returns real file path instead of content URI numeric ID
- Toast prompt changed to WeChat-style path `/sdcard/Download/QR/filename`

### v6 (2026-07-15)

- Save prompt changed to WeChat-style path

### v5 (2026-07-15)

- Added unified checksum support
- Added original filename and extension restoration
- Added draft box resume scanning

## 📁 File Structure

```
qr-native-tool/
├── QR合成-原生-v7.apk      # Android native APP
├── QR生成-v6.vbs            # Word VBS QR code generation script
├── README.md                # Chinese documentation
├── docs/
│   └── README_en.md         # English documentation
└── .gitignore               # Git ignore rules
```

## ⚠️ Disclaimer

This tool is intended for personal use only, for transferring your own files in legitimate scenarios. Do not use it to infringe on others' privacy, steal data, or violate laws and regulations. The user assumes full responsibility for any consequences arising from the use of this tool.

## 📄 License

MIT License

⭐ If this tool helps you, please give it a Star!
