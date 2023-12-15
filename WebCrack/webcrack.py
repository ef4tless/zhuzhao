import os
import datetime

from crack.crack_task import CrackTask


# 设置开始编号
num = 1


def single_process_crack(url_list):

    all_num = len(url_list)
    cur_num = num
    print("总任务数: " + str(all_num))
    for url in url_list[num-1:]:
        CrackTask().run(cur_num, url)
        cur_num += 1


if __name__ == '__main__':
    try:
        import conf.config
    except:
        print("加载配置文件失败！")
        exit(0)

    url_file_name = 'url.txt'

    if '://' in url_file_name:
        CrackTask().run(1, url_file_name)
    else:
        url_list = []
        if os.path.exists(url_file_name):
            print(url_file_name, "exists!\n")
            with open(url_file_name, 'r', encoding="UTF-8") as url_file:
                for url in url_file.readlines():
                    url = url.strip()
                    if url.startswith('#') or url == '' or ('.edu.cn' in url) or ('.gov.cn' in url):
                        continue
                    url_list.append(url)
            start = datetime.datetime.now()
            single_process_crack(url_list)
            end = datetime.datetime.now()
            print(f'All processes done! Cost time: {str(end - start)}')
        else:
            print(url_file_name + " not exist!")
            exit(0)
