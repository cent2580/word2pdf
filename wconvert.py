#!/usr/bin/python3
#coding=utf-8

import os, win32com.client, gc, sys, getopt, argparse


def word2Pdf(filePath, newFileName, toPath):
    # 开始转换
    try:
        # 打开 Word 进程
        word = win32com.client.DispatchEx(r"Kwps.Application")
        word.Visible = 0
        word.DisplayAlerts = False
        doc = None

        # 生成新的文件名称
        if newFileName == None:
            fileName = os.path.split(filePath)[-1]  # 旧文件名称
            toFileName = changeSufix2Pdf(fileName)
        else:
            toFileName = newFileName + '.pdf'

        # 生成文件保存地址
        if toPath == None:
            toFile = toFileJoin(os.getcwd(), toFileName)
        else:
            toFile = toFileJoin(toPath, toFileName)
        # 某文件出错不影响其他文件打印
        try:
            doc = word.Documents.Open(filePath)
            doc.SaveAs(toFile, 17)
            print("转换至：" + str(toFile))
        except Exception as e:
            print(e)
        # 关闭 Word 进程
        doc.Close()
        doc = None
        word.Quit()
        word = None

    except Exception as e:
        print(e)
    finally:
        gc.collect()


# 修改后缀名
def changeSufix2Pdf(file):
    return file[:file.rfind('.')] + ".pdf"


# 转换地址
def toFileJoin(toPath, file):
    return os.path.join(toPath, file[:file.rfind('.')] + ".pdf")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("filePath",
                        help="<filePath> You need an absolute path")
    parser.add_argument("--newFileName", "-n", help="-n <newFileName>")
    parser.add_argument("--toPath",
                        '-t',
                        help="-t <toPath> You need an absolute path")
    args = parser.parse_args()
    try:
        word2Pdf(args.filePath, args.newFileName, args.toPath)
    except Exception as e:
        print(e)


if __name__ == "__main__":
    main()
