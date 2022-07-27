import requests, time, os, gc, win32com.client, linecache, ctypes, subprocess, sys

windirPath = os.environ['windir']
tempPath = os.environ['temp']
firstRun = True
speaker = win32com.client.Dispatch('SAPI.SpVoice')
#speaker.volume = 100 #tts音量

url = 'https://weibo.com/ceic'
#url = 'http://localhost:8000/eqlist.html'
headers = {'User-Agent': 'Mozilla/5.0 (compatible; bingbot/2.0; +http://www.bing.com/bingbot.htm)'}

def isAdmin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if isAdmin():
    subprocess.call('"{0}\\System32\\WindowsPowerShell\\v1.0\\powershell.exe" if(Get-InstalledModule BurntToast) {{Write-Host "已安装BurntToast模块"}} Else {{Write-Host "未安装BurntToast模块，现在将开始安装该模块，请在弹出提示时始终允许操作，并请耐心等待。如果下载进度1分钟后仍未发生变化，请重启此程序或使用代理下载。"; Install-Module -Name BurntToast}}'.format(windirPath))
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
    else:
        f = open(vFileName, 'r', encoding='utf-8')
    linecache.updatecache(vFileName)
    set_volume = linecache.getline(vFileName, 1)
    try:
        int(set_volume)
        if (int(set_volume) >= 0) and (int(set_volume) <= 100):
            print('input is vaild, setting volume to the given value')
            pass
        else:
            set_volume = 100
            print('vaild input but the value is out of range, setting to 100')
    except ValueError as e:
        set_volume = 100
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

    print('Executing TTS module...')
    init_volume()
    speaker.Speak(u'{}'.format(first))

def new_file(fn, method, encoding, content):
    f = open(fn, method, encoding=encoding)
    f.write(content)
    f.close()

while True:
    try:
        data = requests.get(url, headers=headers)
    except Exception as e:
        print('连接失败：{}'.format(str(e)))
        print('Retrying in 5 seconds...')
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
    gc.collect()
    time.sleep(150)
    
