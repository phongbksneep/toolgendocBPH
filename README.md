# BPHH DocGen – Electron App

Phiên bản mới sử dụng Electron + React thay cho Tkinter.

## Yêu cầu

- Node.js ≥ 18
- npm ≥ 9

## Chạy Dev (Electron + React hot-reload)

```bash
cd electron-app
npm install
npm run dev
```

Lệnh trên khởi động Vite (port 5173) và Electron cùng lúc.

## Build phân phối

### macOS (.dmg)

```bash
cd electron-app
npm run build
```

Output: `electron-app/release/BPHH-DocGen-*.dmg`

### Windows (NSIS installer)

Cần chạy trên máy Windows hoặc CI Windows:

```bat
cd electron-app
npm install
npm run build
```

Output: `electron-app/release/BPHH-DocGen Setup *.exe`

## Cấu trúc

```
electron-app/
├── electron/
│   ├── main.cjs       # Electron main process (IPC, file I/O, DOCX gen)
│   └── preload.cjs    # contextBridge → window.api
├── src/
│   ├── App.jsx        # Root React component
│   ├── components/
│   │   └── FieldForm.jsx   # Scrollable labeled form
│   ├── main.jsx
│   └── index.css
├── index.html
├── vite.config.js
└── package.json
```

## Tính năng

- Mở / Lưu / Lưu thành JSON (v1, v2, v3 format)
- Global Fields + Job Fields editor
- Tìm kiếm trường theo label / key
- Generate DOCX (dùng `docxtemplater` – không cần Python)
- Tự động xóa highlight vàng sau khi fill placeholder
- Hỗ trợ macOS và Windows
