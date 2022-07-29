#!/usr/bin/python
# -*- coding: utf-8 -*-

import requests, time, os, gc, win32com.client, linecache, ctypes, subprocess, sys, json, logging

windirPath = os.environ['windir']
tempPath = os.environ['temp']
firstRun = True
speaker = win32com.client.Dispatch('SAPI.SpVoice')
updateUrl = 'https://239252.xyz/version/fetchWeibo/version.json'
releaseTime = 1659085726
version = '0.3.2'
url = 'https://weibo.com/ceic'
headers = {'User-Agent': 'Mozilla/5.0 (compatible; bingbot/2.0; +http://www.bing.com/bingbot.htm)'}
baseName = os.path.basename(__file__).split('.')[0]+'.exe'
LOG_FORMAT = "[%(asctime)s/%(levelname)s] %(message)s"
DATE_FORMAT = "%Y/%m/%d %H:%M:%S %p"
logging.basicConfig(filename='latest.log', level=logging.DEBUG, format=LOG_FORMAT, datefmt=DATE_FORMAT)

def isAdmin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if isAdmin():
    #subprocess.call('"{0}\\System32\\WindowsPowerShell\\v1.0\\powershell.exe" if(Get-InstalledModule BurntToast) {{Write-Host "已安装BurntToast模块"}} Else {{Write-Host "未安装BurntToast模块，现在将开始安装该模块，请在弹出提示时始终允许操作，并请耐心等待。如果下载进度1分钟后仍未发生变化，请重启此程序或使用代理下载。您也可以尝试使用手动安装脚本install.bat来进行安装。如果您已经通过脚本完成安装，请按N拒绝下载并正常启动程序。"; Install-Module -Name BurntToast}}'.format(windirPath))
    subprocess.call('"{0}\\System32\\WindowsPowerShell\\v1.0\\powershell.exe" Set-ExecutionPolicy -ExecutionPolicy Bypass'.format(windirPath))
    os.system('ftype Microsoft.PowerShellScript.1="{0}\\system32\\WindowsPowerShell\\v1.0\\powershell.exe" "%1"'.format(windirPath))
else:
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    sys.exit(0)
time.sleep(1)

console = ctypes.windll.kernel32.GetConsoleWindow()
if console != 0:
    ctypes.windll.user32.ShowWindow(console, 0)
    ctypes.windll.kernel32.CloseHandle(console) #隐藏窗口


def init_volume():
    vFileName = 'v.txt'
    if os.path.isfile(vFileName) == False:
        new_file(vFileName, 'w', 'utf-8', '100') #default to 100 if not defined
        logging.info('TTS语音合成 - 配置文件不存在，生成新文件')
    else:
        f = open(vFileName, 'r', encoding='utf-8')
        logging.info('TTS语音合成 - 配置文件存在，读取配置')
    linecache.updatecache(vFileName)
    set_volume = linecache.getline(vFileName, 1)
    try:
        int(set_volume)
        if (int(set_volume) >= 0) and (int(set_volume) <= 100):
            print('input is vaild, setting volume to the given value')
            logging.info('TTS语音合成 - 给定的音量值有效，音量设置为：{}'.format(set_volume))
            pass
        else:
            set_volume = 100
            logging.warning('TTS语音合成 - 给定的音量值超界，重置为100')
            print('vaild input but the value is out of range, setting to 100')
    except ValueError as e:
        set_volume = 100
        logging.warning('TTS语音合成 - 给定的音量值无效，重置为100')
        print('invaild input, setting volume to 100')
    speaker.volume = set_volume
    print('set volume to:', set_volume)

def push_notification():
    c = 0
    f = open('result.txt', 'r', encoding='utf-8')
    lines = f.readlines()
    first = lines[0]

    print('first:',first)
    if lines[0].count('自动') == 1:
        data = lines[0].split('：', 1)
    elif lines[0].count('正式') == 1:
        data = lines[0].split('：', 1)
    #data[0] = data[0].lstrip('据')
    print('data:',data)
    line = data[0]
    message = data[1]
    print(line)

    f = open('{0}\\cencNotify.ps1'.format(tempPath), 'w', encoding='gb2312')
    f.write('New-BurntToastNotification -Text \"{0}\",\"{1}\" -AppLogo ".\ico\cenc.ico"'.format(line,message))
    f.close()
    os.system('"{0}\\cencNotify.ps1"'.format(tempPath))
    logging.info('地震信息处理完成，推送通知')

    print('Executing TTS module...')
    init_volume()
    logging.info('TTS语音合成 - 调用模块，朗读文本')
    speaker.Speak(u'{}'.format(first))

def new_file(fn, method, encoding, content):
    f = open(fn, method, encoding=encoding)
    f.write(content)
    f.close()

def update_check():
    try:
        update = requests.get(url=updateUrl, headers=headers)
        logging.info('建立连接 {}'.format(updateUrl))
    except Exception as e:
        print('连接失败：{}'.format(str(e)))
        logging.warning('无法连接到更新服务器 {}'.format(updateUrl))
        pass
    else:
        f = open('version.json', 'w', encoding='utf-8')
        f.write(update.text)
        f.close()

        with open('version.json', 'r', encoding='utf-8') as update_result:
            result = json.load(update_result)
            print(result['time'])
            if releaseTime < (result['time']):
                print('有新版本可用！')
                Mbox('更新检测', '目前版本{}，最新版本{}\n新版本功能：\n{}'.format(version, result['version'], result['desc']), 64)
                logging.info('更新检测 - 检测到新版本：{}，目前版本：{}'.format(result['version'], version))
            else:
                print('目前已是最新版本！')
                logging.info('更新检测 - 目前版本：{}，无需更新'.format(version))

def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)

os.system('tasklist > "{}\\image_list.txt"'.format(tempPath))
f = open('{}\\image_list.txt'.format(tempPath), 'r')
image_list = f.readlines()
image_count = 0

for line in image_list:
    if baseName in line:
        image_count += 1
if image_count > 2:
    Mbox('程序多开检测', '检测到程序多开，请运行kill.bat关闭程序\n然后再运行此程序', 16)
    logging.warning('多开检测 - 未通过检测，程序将会自动退出')
    sys.exit()
else:
    print('正常运行')
    logging.info('多开检测 - 正常，通过检测')
    pass

f.close()

update_check()

while True:
    try:
        data = requests.get(url, headers=headers)
        logging.info('建立连接 {}'.format(url))
    except Exception as e:
        print('连接失败：{}'.format(str(e)))
        print('Retrying in 5 seconds...')
        logging.warning('无法连接到服务器 {}，获取地震信息失败'.format(url))
        # Mbox('首次连接失败','请检查您的网络，程序将继续尝试连接，若仍无法连接则自动退出', 48)
        time.sleep(5)
        continue
    data.encoding = 'utf-8'

    new_file('weibo_ceic.html', 'w', 'utf-8', data.text)

    c = 0
    f = open('weibo_ceic.html', 'r', encoding='utf-8')
    fileIsWritten = True
    for line in f.readlines():
        if '#地震快讯#' in line:
            if '中国地震台网' in line:
                line = line.split('（ <a', 1)[0] #end
                line = line.split('#地震快讯#</a>', 1)[1] #begin
            else:
                logging.info('地震信息获取成功，内容不符合通知条件，等待10秒后重试')
                print('内容不匹配 等待重试')
                time.sleep(10)
                continue
            if fileIsWritten == True:
                fileIsWritten = False
                new_file('result.txt', 'w', 'utf-8', line+'\n')
            else:
                new_file('result.txt', 'a', 'utf-8', line+'\n')
            print(line)
        c += 1
    f.close()
    print()
    fileIsWritten = True

    fileName = 'result.txt'
    linecache.updatecache(fileName)
    fRead = linecache.getline(fileName, 1)
    try:
        line1web
    except NameError:
        line1web = fRead

    try:
        line1file
    except NameError:
        line1file = line1web
        time.sleep(1)
        #f.close()
        continue

    if fRead != line1file:
        print('文件内容不相同')
        firstRun = False
        line1file = fRead
        f.close()
        print('推送通知')
        push_notification()
    else:
        print('文件内容相同')
        if firstRun == True:
            firstRun = False
            time.sleep(1)
            print('首次运行，推送通知')
            push_notification()

    print('Waiting...')
    logging.info('执行完成，150秒后发起下一次请求')
    gc.collect()
    time.sleep(150)
    
