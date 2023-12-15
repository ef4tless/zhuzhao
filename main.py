import os


def title():
    title = r'''
        .__                    .__                      
________|  |__   __ __ ________|  |__  _____     ____   
\___   /|  |  \ |  |  \\___   /|  |  \ \__  \   /  _ \  
 /    / |   Y  \|  |  / /    / |   Y  \ / __ \_(  <_> ) 
/_____ \|___|  /|____/ /_____ \|___|  /(____  / \____/  
      \/     \/              \/     \/      \/          
    烛照 by talentsec
    '''
    print(title)

def scanpasswd():
    try:
        required_files = ['WebCrack/url.txt']
        missing_files = []
        for file in required_files:
            if not os.path.exists(file):
                missing_files.append(file)
        if missing_files:
            print("缺少必要的文件:", ", ".join(missing_files))
            return

        while True:
            print("选择要执行的功能:")
            print("1. 执行webcrack扫描")
            print("2. 执行成功结果数据处理")
            print("3. 退出")
            choice = input("请输入您的选择 (1/2/3): ")

            if choice == '1':
                os.chdir('WebCrack/')
                os.system('python webcrack.py')
                break  # 退出循环
            elif choice == '2':
                os.chdir('WebCrack/')
                os.system('python scanpasswd.py')
                # 在这里添加其他操作的代码
                break  # 退出循环
            elif choice == '3':
                print("退出程序")
                break  # 退出循环
            else:
                print("无效的输入，请重新输入")

    except KeyboardInterrupt:
        print("\n操作被用户中断，返回主菜单...")

def scanvlun():
    try:
        required_files = ['scanvuln/result.txt', 'scanvuln/info.xlsx', 'scanvuln/天眼查.xlsx']
        missing_files = []
        for file in required_files:
            if not os.path.exists(file):
                missing_files.append(file)
        if missing_files:
            print("缺少必要的文件:", ", ".join(missing_files))
            return

        os.chdir('scanvuln/')
        os.system('python scanvuln.py')
    except KeyboardInterrupt:
        print("\n操作被用户中断，返回主菜单...")

def scanhighriskport():
    try:
        required_files = ['scanport/scan_results.xlsx', 'scanport/info.xlsx', 'scanvuln/天眼查.xlsx']
        missing_files = []
        for file in required_files:
            if not os.path.exists(file):
                missing_files.append(file)
        if missing_files:
            print("缺少必要的文件:", ", ".join(missing_files))
            return
        os.chdir('scanport/')
        os.system('python scanhighriskport.py')
    except KeyboardInterrupt:
        print("\n操作被用户中断，返回主菜单...")


def main():
    title()
    while True:
        print("请选择功能:")
        print("1 - 扫描弱口令")
        print("2 - 扫描高危漏洞")
        print("3 - 扫描高危端口")
        print("0 - 退出")
        choice = input("请输入选项：")

        if choice == "1":
            scanpasswd()
            os.chdir('../')
        elif choice == "2":
            scanvlun()
            os.chdir('../')
        elif choice == "3":
            scanhighriskport()
            os.chdir('../')
        elif choice == "0":
            print("退出程序")
            break
        else:
            print("无效选项，请重新输入。")


if __name__ == '__main__':
    main()

