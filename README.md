### fetchWeibo_release (For CENC only)

本软件只适用于运行 Windows 10 或更高版本的系统。

此外，需要 Windows PowerShell Package: `BurntToast` 以发送通知；

需要允许执行 PowerShell 脚本，并同时将 `.ps1` 文件关联到直接执行而非打开记事本；

需要 Python `requests` 和 `pypiwin32` 库。

 

此软件应当正常运行，以上所提及的 Package 中，只有 `BurntToast` 需要额外下载。此程序不再提供下载功能，但您可以使用手动安装脚本进行安装。

软件正常运行后，会转后台运行。如果想退出，请打开任务管理器，结束对应的进程即可，或者以管理员身份运行`kill.bat`。

#### 软件运行时会产生文件并存放至运行目录，因此请将软件解压至文件夹内再运行。

#### 本软件包含TTS文字转语音功能，不建议外放，同时也请注意音量。

===Changelog===

`v0.3.6` 1. 修复PowerShell脚本的换行问题
2. `json.load()`改为使用`json.loads()`
3. 输出日志更为详细
4. 修复通知图标显示问题
5. 移除日志中不应存在的`AM/PM`12小时制时间标识

`v0.3.6-pre` 新增自动报深度信息支持

`v0.3.5`  修复多开检测功能

`v0.3.4`  新增备用更新服务器

`v0.3.3`  日志会在程序每次启动时重写

`v0.3.2`  移除自动下载功能

`v0.3.1`  修复多开检测功能，更新时将会直接显示新版本的更新内容。

`v0.3`  新增多开检测功能和运行日志。新增文件`kill.bat`，便于快速关闭程序。

`v0.2.1`  新增手动安装脚本，用户可在下载失败的情况下通过脚本`install.bat`手动进行安装。

`v0.2`  新增音量调节功能，可通过修改`v.txt`中的数值来控制输出的音量大小。范围为0-100，若输入无效值则重置为100。程序将始终读取第一行的内容作为输入值并忽略其他行的内容。

`v0.1a`  修复同一场地震重复播报的问题

`v0.1`  软件发布
