# 📝 MyNote - 极简桌面便笺

<div align="center">

Claude Code 和 Google Gemini 辅助创建的：一个精致的、无干扰的 Windows 桌面便笺应用，支持毛玻璃特效、桌面嵌入、待办事项管理，让您的想法触手可及。

![License](https://img.shields.io/badge/license-MIT-blue)
![Python](https://img.shields.io/badge/Python-3.8+-yellow)
![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue)

[功能特性](#-功能特性) • [快速开始](#-快速开始) • [使用说明](#-使用说明) • [下载](#-下载)

</div>

---

## ✨ 功能特性

### 🎨 极致视觉
- **毛玻璃特效** - 采用 Windows Acrylic 材质，完美融入系统 UI。
- **无边框设计** - 告别笨重的标题栏，内容即是一切。
- **现代化圆角** - 适配 Windows 11 设计语言。
- **高 DPI 适配** - 在 2K/4K 屏幕上字迹清晰锐利。

### 📝 高效记录
- **待办事项** - 支持复选框，自动标记完成状态。
- **智能排版** - 自动换行，中文标点避头尾处理。
- **数据安全** - 实时自动保存，无需手动操作，不怕断电丢失。
- **撤销/重做** - 支持操作回溯，不再担心误删。

### 🎮 灵活交互
- **双重模式**：
    - 📌 **置顶模式** - 始终悬浮在所有窗口之上，记录待办不遗忘。
    - 🖥️ **嵌入桌面** - 像壁纸的一部分一样贴在桌面，清爽不打扰。
- **拖拽排序** - 长按拖动即可调整事项顺序。
- **右键菜单** - 深度定制的右键菜单，提供丰富的功能入口。

---

## 🚀 快速开始

### 方式一：直接运行（推荐普通用户）

1. 在 [Releases](../../releases) 页面下载最新的 `MyNote.exe`。
2. 双击运行即可，无需安装 Python 环境。
3. 如果被杀毒软件拦截，请添加信任（因未签名所致）。

### 方式二：源码运行（推荐开发者）

#### 环境要求
- Windows 10 或 Windows 11
- Python 3.8+

#### 启动步骤

1. **克隆仓库**
   ```bash
   git clone https://github.com/lecrix/my-note.git
   cd my-note
   ```

2. **安装依赖**
   本项目主要依赖 Python 标准库，仅需安装打包工具（如果需要打包）：
   ```bash
   pip install -r requirements.txt
   ```

3. **运行程序**
   ```bash
   python note.py
   ```

---

## 📖 使用说明

### ⌨️ 快捷键指南

| 按键组合       | 功能     | 说明                                        |
| -------------- | -------- | ------------------------------------------- |
| `Enter`        | 新建事项 | 在当前项下方插入新待办                      |
| `Ctrl + Enter` | 切换状态 | 标记当前项为完成/未完成                     |
| `Ctrl + D`     | 删除     | 删除当前选中的事项                          |
| `Ctrl + ↑`     | 上移     | 将当前项上移一位                            |
| `Ctrl + ↓`     | 下移     | 将当前项下移一位                            |
| `Ctrl + Z`     | 撤销     | 撤销上一步操作（仅限删除/排序等应用级操作） |
| `Key Up/Down`  | 焦点移动 | 在不同事项间切换光标                        |

### 🖱️ 鼠标操作

- **拖动窗口**：按住窗口空白处或顶部隐藏的标题栏区域即可拖动。
- **调整大小**：鼠标移动到窗口边缘（右侧、底部、右下角），光标变化后拖动。
- **拖拽排序**：鼠标**长按**待办事项前面的复选框区域，即可进入拖拽模式。

---

## ⚙️ 数据存储

MyNote 的所有数据都保存在程序同级目录下的 `notes_data.json` 文件中。

### 数据结构示例
您可以参考项目中的 `notes_data.example.json`：

<details>
<summary>点击查看 JSON 结构</summary>

```json
{
  "items": [
    {
      "id": 1,
      "text": "示例待办事项",
      "completed": false,
      "created_at": "2025-01-01T00:00:00"
    }
  ],
  "window": {
    "x": 100, "y": 100,
    "width": 300, "height": 400
  },
  "settings": {
    "mode": "topmost",
    "visibility_mode": "always_visible",
    "font_size": 13,
    "opacity_focused": 1.0,
    "opacity_unfocused": 0.7
  }
}
```
</details>

---

## 📦 打包构建

如果您想生成自己的 `.exe` 文件，可以使用我们提供的配置脚本：

```bash
# 确保已安装 pyinstaller
pip install pyinstaller

# 执行打包
pyinstaller build_config.py
```

打包完成后，可执行文件将生成在 `dist/MyNote.exe`。

---

## 📁 项目结构

```
my-note/
├── note.py                 # 核心代码（UI逻辑、WinAPI调用、事件处理）
├── build_config.py         # PyInstaller 打包配置文件
├── notes_data.json         # 用户数据（自动生成，已在 .gitignore 中）
├── notes_data.example.json # 数据结构模板
├── version_info.txt        # Windows 版本信息资源
├── requirements.txt        # 依赖列表
├── LICENSE                 # MIT 许可证
├── README.md               # 项目文档
└── assets/                 # 静态资源
    └── icon.ico            # 程序图标
```

---

## ❓ 常见问题

<details>
<summary><b>Q1: 为什么运行后窗口一片黑/不透明？</b></summary>
A: 毛玻璃特效依赖 Windows 的 DWM 服务。请确保您使用的是 Windows 10/11，并在系统设置中开启了透明效果。如果是 Windows 7，可能无法支持此特效。
</details>

<details>
<summary><b>Q2: 被杀毒软件报毒怎么办？</b></summary>
A: 这是使用 PyInstaller 打包 Python 程序的常见误报。因为程序没有数字签名，且包含修改窗口样式的底层 API 调用。您可以将文件加入白名单，或直接运行 Python 源码。
</details>

<details>
<summary><b>Q3: 如何重置所有设置？</b></summary>
A: 关闭程序，删除目录下的 `notes_data.json` 文件，重新启动即可。
</details>

---

## 📄 许可证

本项目基于 [MIT License](LICENSE) 开源。您可以免费使用、修改和分发，但请保留版权声明。

---

<div align="center">

### ⭐ 如果您喜欢这个小工具，请点亮 Star 支持一下！

Made with ❤️ by Lecris using Python & Tkinter

</div>
