#!/usr/bin/python
# -*- coding: utf-8 -*-

import requests, time, os, gc, win32com.client, linecache, ctypes, subprocess, sys, json, logging, webbrowser, platform
from win10toast_click import ToastNotifier

toaster = ToastNotifier()
debugMode = True if sys.gettrace() else False
windirPath = os.environ['windir']
tempPath = os.environ['temp']
firstRun = True
speaker = win32com.client.Dispatch('SAPI.SpVoice')
updateUrl = 'http://lh2.hkg-ali.server.sweetcandy233.top/assets/fetchWeibo/version.json'
bakUpdateUrl = 'http://api-production.na-ms.server.sweetcandy233.top/assets/fetchWeibo/version.json'
url = 'https://weibo.com/ceic'
# url = 'http://127.0.0.1:8000/weibo_ceic.html'
# headers = {'User-Agent': 'Mozilla/5.0 (compatible; bingbot/2.0; +http://www.bing.com/bingbot.htm)'}
headers = {'User-Agent': 'Mozilla/5.0(compatible; Baiduspider/2.0; +http://www.baidu.com/search/spider.html)'}
appHeaders = {'User-Agent': 'Mozilla/5.0 (compatible; fetchWeiboApp/0.3.8)'}
# baseName = os.path.basename(__file__).split('.')[0]+'.exe'
baseName = 'fetchWeibo'
LOG_FORMAT = "[%(asctime)s/%(levelname)s] %(message)s"
DATE_FORMAT = "%Y/%m/%d %H:%M:%S"
logging.basicConfig(filename='latest.log', encoding='utf-8', level=logging.INFO, format=LOG_FORMAT, datefmt=DATE_FORMAT)

appInfo = """{
    "name": "fetchWeibo",
    "version": "0.3.8",
    "time": 1688376308
}"""

appInfo = json.loads(appInfo)

def platform_check():
# > 10.0.10240 counts as Windows 10
# > 10.0.22000 counts as Windows 11
    if ('Windows-10' or 'Windows-11') not in platform.platform():
        Mbox('注意', '您的操作系统版本可能不受支持，目前仅支持Windows 10及以上版本，若无法正常使用程序，请确定使用的系统符合条件。若仍有问题，请截图反馈至开发者。\n'+platform.platform()+'\n'+platform.version()+'\nfetchWeiboApp/0.3.8', 48)
platform_check()

def isAdmin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if debugMode:
    print('debug mode is running')
    logging.info('调试模式正在运行')
    pass
else:
    print('debug mode is not running')
    logging.info('调试模式未运行')
    if isAdmin():
        subprocess.call('"{0}\\System32\\WindowsPowerShell\\v1.0\\powershell.exe" Set-ExecutionPolicy -ExecutionPolicy Bypass'.format(windirPath))
        os.system('ftype Microsoft.PowerShellScript.1="{0}\\system32\\WindowsPowerShell\\v1.0\\powershell.exe" "%1"'.format(windirPath))
        logging.info('创建文件关联 - ftype Microsoft.PowerShellScript.1="{0}\\system32\\WindowsPowerShell\\v1.0\\powershell.exe" "%1"'.format(windirPath))
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        sys.exit(0)
time.sleep(1)

console = ctypes.windll.kernel32.GetConsoleWindow()
if console != 0:
    ctypes.windll.user32.ShowWindow(console, 0)
    ctypes.windll.kernel32.CloseHandle(console) # 隐藏窗口

def init_volume():
    # recommended value: 70
    vFileName = 'v.txt'
    if os.path.isfile(vFileName) == False:
        new_file(vFileName, 'w', 'utf-8', '70') #default to 70 if not defined
        logging.info('TTS语音合成 - 配置文件不存在，生成新文件')
    else:
        f = open(vFileName, 'r', encoding='utf-8')
        logging.info('TTS语音合成 - 配置文件存在，读取配置')
    linecache.updatecache(vFileName)
    set_volume = linecache.getline(vFileName, 1) # 到底为什么日志会多出一行空行呢
    try:
        int(set_volume)
        if (int(set_volume) >= 0) and (int(set_volume) <= 100):
            print('input is vaild, setting volume to the given value')
            logging.info('TTS语音合成 - 给定的音量值有效，音量设置为：{}'.format(set_volume))
            pass
        else:
            set_volume = 70
            logging.warning('TTS语音合成 - 给定的音量值超界，重置为70')
            print('vaild input but the value is out of range, setting to 70')
    except ValueError as e:
        set_volume = 70
        logging.warning('TTS语音合成 - 给定的音量值无效，重置为70')
        print('invaild input, setting volume to 70')
    speaker.volume = set_volume
    print('set volume to:', set_volume)

def execute_url():
    try:
        webbrowser.open_new(url)
    except Exception:
        print('Failed to open URL.')

def push_notification():
    c = 0
    f = open('result.txt', 'r', encoding='utf-8')
    lines = f.readlines()
    first = lines[0]

    print('first:', first)
    if '：' in lines[0]: # 正常格式，使用冒号进行截取
        data = lines[0].split('：', 1)
    else: # 人工发布格式，对第一个逗号进行截取，取逗号后面的内容作为正文
        data = lines[0].split('，', 1)
    autoFlag = False
    depthAutoFlag = False
    if '自动' in lines[0]:
        autoFlag = True
        try:
            autoData = requests.get(url='https://www.appfly.cn/api/earthquake/list?page=1', headers=headers)
            depthAutoFlag = True
        except Exception:
            depthAutoFlag = False
            logging.warning('自动报深度获取失败，跳过。')
            pass
        if depthAutoFlag:
            autoData_json = json.loads(autoData.text)
            if (autoData_json['data'][0]['status']) == 'automatic': # 自动报位置不一定为[0]
                depthAuto = autoData_json['data'][0]['depth']
                depthAuto = str(int(float(depthAuto)))
                print(depthAuto)
                # depthAutoFlag = True

    print('data:', data)
    line = data[0]
    try:
        message = data[1].rstrip('\n')
        if '若有震感' in message:
            message = message.split('若有震感')[0]
    except IndexError: # 格式异常时使用a标签作为结束判断
        message = data[0].strip('（<a')[0]
        if '若有震感' in message:
            message = message.split('若有震感')[0]
    if depthAutoFlag:
        message_split = message.split('，')
        msg1 = message_split[0]
        msg2 = message_split[1]
        # msgList 已修改，待检验，message_split[2]始终是“最终结果以正式速报为准。”
        try:
            depthAuto
        except UnboundLocalError:
            print('自动报信息不匹配，忽略自动报深度信息。')
            msgList = [msg1, '，', msg2, '，', message_split[2]]
        else:
            msgList = [msg1, '，', msg2, '，', '震源深度{}千米'.format(depthAuto), '，', message_split[2]]
        final_content = ''.join(msgList)
        print(final_content)
    else:
        pass
    print(line)
    
    # f = open('{0}\\cencNotify.ps1'.format(tempPath), 'w', encoding='gb2312')
    if depthAutoFlag:
        input_content = final_content
    else:
        input_content = message
    # f.write('New-BurntToastNotification -Text \"{0}\",\"{1}\" -AppLogo \"{2}\\ico\\cenc.ico\"'.format(line, input_content, os.getcwd()))
    # f.close()
    # os.system('"{0}\\cencNotify.ps1"'.format(tempPath))
    toaster.show_toast(title=line, msg=input_content, icon_path=r'.\\ico\\cenc.ico', duration=None, threaded=True, callback_on_click=execute_url)
    # subprocess.call('"{3}\\System32\\WindowsPowerShell\\v1.0\\powershell.exe" -command New-BurntToastNotification -Text \"{0}\",\"{1}\" -AppLogo \"{2}\\ico\\cenc.ico\"'.format(line, input_content, os.getcwd(), windirPath))
    logging.info('地震信息处理完成，推送通知')
    logging.info('isAuto: {}, earthquakeInfo: "{}"'.format(autoFlag, input_content))

    print('Executing TTS module...')
    init_volume()
    logging.info('TTS语音合成 - 调用模块，朗读文本')
    if depthAutoFlag:
        tts_content = final_content
    else:
        tts_content = message

    speaker.Speak(u'{}，{}'.format(line, tts_content), 1)

def new_file(fn, method, encoding, content):
    f = open(fn, method, encoding=encoding)
    f.write(content)
    f.close()

def update_check():
    try:
        update = requests.get(url=updateUrl, headers=appHeaders)
    except Exception as e:
        print('连接失败：{}'.format(str(e)))
        logging.warning('无法连接到更新服务器')
        try:
            update = requests.get(url=bakUpdateUrl, headers=appHeaders)
        except Exception as e:
            print('连接失败：{}'.format(str(e)))
            logging.warning('无法连接到备用更新服务器')
            logging.warning('更新检查失败，跳过。')
            pass
    else:
        result = json.loads(update.text)

        if result['status']['toggle'] == 'off':
            Mbox('公告', '{}'.format(result['status']['em_announcement']), 48)
            sys.exit(1)

        print(result['time'])
        if appInfo['time'] < (result['time']):
            print('有新版本可用！')
            Mbox('更新检测', '目前版本{}，最新版本{}\n更新内容：\n{}\n点击确定以继续运行程序'.format(appInfo['version'], result['version'], result['desc']), 64)
            logging.info('更新检测 - 检测到新版本：{}，目前版本：{}'.format(result['version'], appInfo['version']))
        else:
            print('目前已是最新版本！')
            logging.info('更新检测 - 目前版本：{}，无需更新'.format(appInfo['version']))

def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)

# begin

if os.path.exists('latest.log'):
    with open('latest.log', 'w', encoding='utf-8') as f:
        f.write('')
else:
    pass

os.system('tasklist > "{}\\image_list.txt"'.format(tempPath))
f = open('{}\\image_list.txt'.format(tempPath), 'r')
image_list = f.readlines()
image_count = 0

for line in image_list:
    if baseName in line:
        image_count += 1
if image_count > 2:
    Mbox('程序多开检测', '检测到程序多开，请运行kill.bat关闭程序\n然后再运行此程序', 16)
    logging.info('多开检测 - 返回值：{}'.format(image_count))
    logging.warning('多开检测 - 未通过检测，程序将会自动退出')
    sys.exit(1)
else:
    print('正常运行')
    logging.info('多开检测 - 返回值：{}'.format(image_count))
    logging.info('多开检测 - 正常')
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
        # pending: first try failed
        # Mbox('首次连接失败','请检查您的网络，程序将继续尝试连接，若仍无法连接则自动退出', 48)
        time.sleep(5)
        continue
    data.encoding = 'utf-8'

    new_file('weibo_ceic.html', 'w', 'utf-8', data.text)

    # 文件写入内容暂时与tts、通知内容不同（仅限自动报时），待修改
    c = 0
    f = open('weibo_ceic.html', 'r', encoding='utf-8')
    fileIsWritten = True
    for line in f.readlines():
        if '#地震快讯#' in line:
            if '据' in line:
                line = line.split(' 据')[1] #开始判断
                line = line.split('（<a ')[0] #结束判断
            elif '中国地震台网' in line:
                line = line.split('（ <a', 1)[0] #结束判断
                line = line.split('#地震快讯#</a>', 1)[1] #开始判断
                #if '若有震感' in line:
                #    line.split('若有震感')[0] #结束判断
                
            else:
                logging.info('检测到指定tag，但tag信息不符合，等待10秒后重试')
                print('内容不匹配 等待重试')
                time.sleep(10)
                continue
            if fileIsWritten:
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
        continue

    if fRead != line1file:
        print('文件内容不相同')
        firstRun = False
        line1file = fRead
        f.close()
        print('推送通知')
        # try:
        #     pyperclip.copy(line1file)
        # except Exception:
        #     pass
        push_notification()
    else:
        print('文件内容相同')
        if firstRun:
            firstRun = False
            time.sleep(1)
            print('首次运行，推送通知')
            push_notification()

    print('Waiting...')
    logging.info('执行完成，150秒后发起下一次请求')
    gc.collect()
    time.sleep(150)
