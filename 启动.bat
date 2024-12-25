@echo off
rem 工作目录切换为当前目录
cd /d %~dp0
rem 用谷歌最大化打开某个窗口
rem start "chrome.exe" /max "https://www.baidu.com/"
rem 启动python脚本
python .\SearchAndClick4.py
pause
