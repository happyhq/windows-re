# Windows还原点管理工具

## 项目简介
本项目是一个基于Python的Windows系统还原点图形化管理工具，支持创建、查看、还原多个还原点，适用于教育等需要定期恢复系统环境的场景。

## 主要功能
- 创建系统还原点（自定义名称）
- 列出所有还原点
- 选择还原点一键还原（自动重启）
- 多还原点管理

## 依赖环境
- Python 3.x
- Tkinter（Python自带）
- 仅支持Windows系统
- 需以管理员身份运行

## 使用方法
1. 安装Python 3.x，并确保已添加到环境变量。
2. 以管理员身份运行：
   ```bash
   python restore_manager.py
   ```
3. 如需打包为exe：
   ```bash
   pip install pyinstaller
   pyinstaller --noconsole --onefile restore_manager.py
   ```
   打包后在dist目录下生成exe文件。

## 注意事项
- 需以管理员身份运行，否则无法操作还原点。
- 仅支持Windows系统。
- 若创建还原点失败，可能是系统策略或空间不足导致。 