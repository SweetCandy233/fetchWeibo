import requests, time, os, gc, win32com.client

windirPath = os.environ['windir']
tempPath = os.environ['temp']
firstRun = True
speaker = win32com.client.Dispatch('SAPI.SpVoice')

url = 'https://weibo.com/ceic'
headers = {'User-Agent': 'Mozilla/5.0 (compatible; bingbot/2.0; +http://www.bing.com/bingbot.htm)'}

def push_notification():
    c = 0
    f = open('result.txt', 'r', encoding='utf-8')
    lines = f.readlines()
    first = lines[0]

    print('first:',first)
    data = lines[0].split('：', 1)
    print('data:',data)
    line = data[0]
    message = data[1]
    print(line)

    tempps1file = open('{0}\\cencNotify.ps1'.format(tempPath), 'w', encoding='gb2312')
    tempps1file.write('New-BurntToastNotification -Text \"{0}\",\"{1}\" -AppLogo ".\ico\cenc.ico"'.format(line,message))
    tempps1file.close()
    os.system('"{0}\\cencNotify.ps1"'.format(tempPath))

    print('Executing TTS module...')
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
                line = line.split('（ <a', 1)[0]
                line = line.split('#地震快讯#</a>', 1)[1]
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
    with open(fileName, 'r', encoding='utf-8') as f:
        fRead = f.read()
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
