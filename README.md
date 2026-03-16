媒体素材一键打包 PPT (Media-to-PPT Smart Packer)

```markdown
# 🎬 媒体素材一键自适应打包 PPT (Smart Media Packer)

## 📖 项目简介
针对展厅现场屏幕尺寸多变（超宽拼接屏、竖屏等）的痛点，手动排版多比例素材极易拉伸变形。本项目旨在实现“任意比例媒体素材 -> 自动测算并完美适配 PPT 画布”的全自动转化。

## ✨ 核心功能
- **画布自适应探针机制**：自动提取选中的【首个媒体文件】的真实比例，并锁定整个 PPT 画布尺寸，彻底打破 16:9 的死板限制。
- **防扭曲等比最大化**：后续导入的所有图片和视频，将严格按自身比例最大化缩放并绝对居中，保证不拉伸变形。
- **多媒体多文件支持**：完美支持图文及视频（自动设为全屏循环播放），并集成至 Windows “发送到 (Send To)” 菜单，支持一键多选。

## 🛠️ 环境配置
1. 确保本机已安装 Python 3.x 及 Microsoft PowerPoint。
2. 安装所需依赖库：
   ```bash
   pip install pywin32
🚀 安装与右键菜单注入
由于 Windows 普通右键菜单不支持同时传入多个文件，本项目采用更高阶的 “发送到 (Send To)” 菜单注入：

将本工具放置在固定文件夹。

新建 install_sendto.py 并运行以下代码：

Python
import os, sys, win32com.client
# 自动获取 Python 环境和当前工具路径
pythonw_exe = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
tool_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "打包PPT.pyw")

sendto_dir = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'SendTo')
shortcut_path = os.path.join(sendto_dir, "一键打包PPT.lnk")

shell = win32com.client.Dispatch("WScript.Shell")
shortcut = shell.CreateShortCut(shortcut_path)
shortcut.TargetPath = pythonw_exe
shortcut.Arguments = f'"{tool_path}"'
shortcut.IconLocation = pythonw_exe + ",0" 
shortcut.Save()
print("✅ '发送到'快捷方式创建成功！完美支持多选文件！")
💡 使用说明
按住 Ctrl 框选需要打包的多个图片或视频文件（注意：你点击的第一个文件将决定最终 PPT 的比例）。

在选中的文件上右击 -> 发送到 (Send To) -> 选择 一键打包PPT。

后台将静默排版，完成后在同目录生成完美的 output.pptx。
