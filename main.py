#!/usr/bin/env python
# coding:utf-8
# author aibow

"""
遍历目录,查找所有word文档,检查文档是否包含指定关键词

结果日志(result.out):
path:命中次数:命中词列表

错误日志列表(result.err):
time:path:错误说明

工作流程:
- 读取命中词列表和处理目录
- 遍历目录,查找doc,docx文档
- 将文档转换为txt文本
- 检查命中词
- 记录结果
"""

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

from optparse import OptionParser, OptParseError
import os
import os.path
from win32com.client import Dispatch
import shutil
import time
import hashlib

def utf2gbk(s):
    if not s:
        return ''
    try:
        return s.decode('utf-8').encode('gbk', 'ignore')
    except Exception as e:
        return s

def gbk2utf(s):
    if not s:
        return s
    try:
        return s.decode('gbk').encode('utf-8', 'ignore')
    except Exception as e:
        return s


def convert(path, tempPath):
    if not path or not os.path.exists(path) or not os.path.isfile(path):
        raise Exception('Path Not Found Or Path Invalid')
    app = Dispatch('Word.Application')
    app.Visible = 0
    app.DisplayAlerts = 0
    app.Documents.Open(FileName=path)
    app.ActiveDocument.SaveAs(FileName=tempPath, FileFormat=2)
    app.Quit()

def log(msg):
    print '[%s] %s' % (time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()), msg)

def main(dir, wordFile, flush):
    log('Application Starting ...')
    # 检查缓存目录是否存在,不存在则创建
    currPath = os.path.abspath(os.path.curdir)
    cachePath = os.path.join(currPath, 'temp')
    errorPath = os.path.join(currPath, 'error.txt')
    resultPath = os.path.join(currPath, 'result.txt')
    tempPath = os.path.join(currPath, 'doc.tmp')
    # 是否清空缓存
    if flush and os.path.exists(cachePath):
        shutil.rmtree(cachePath)
    if not os.path.exists(cachePath):
        # 尝试创建目录
        os.mkdir(cachePath)
    else:
        if not os.path.isdir(cachePath):
            raise Exception('Cache Path Invalid')
    # 检查目录是否存在
    if not dir or not os.path.exists(dir) or not os.path.isdir(dir):
        raise Exception('Directory Invalid')
    # 检查词库文件是否存在
    if not wordFile or not os.path.exists(wordFile) or not os.path.isfile(wordFile):
        raise Exception('Word File Invalid')
    # 读取词库列表
    log('Loading Word File ...')
    wordList = []
    with open(wordFile, 'r') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            if line[0] == '#':
                continue
            wordList.append(line)
    if len(wordList) == 0:
        raise Exception('Word File Empty')
    log('Word File Load Success, Find %s Word' % len(wordList))
    # 遍历读取word文档
    log('Find Doc File ...')
    docList = []
    for root, dirs, files in os.walk(dir, False):
        for name in files:
            temp = os.path.join(root, name)
            if temp.lower().endswith('.doc') or temp.lower().endswith('.docx'):
                docList.append(temp)
    log('Find Success, Find %s Document' % len(docList))
    # 处理文件
    log('Preprocessor Document ...')
    for docFile in docList:
        log('Convert Document %s' % docFile)
        # 检查是否已经缓存了处理结果
        cacheFile = os.path.join(cachePath, '%s.tmp' % hashlib.md5(docFile).hexdigest())
        # 已经存在缓存
        if os.path.exists(cacheFile):
            continue
        # 处理文档
        try:
            # 移除临时文件
            if os.path.exists(tempPath):
                os.remove(tempPath)
            convert(docFile, tempPath)
            with open(tempPath, 'r') as f:
                body = f.read()
            with open(cacheFile, 'w') as f:
                f.write(gbk2utf('%s\n\n%s' % (docFile, body)))
        except Exception as e:
            with open(errorPath, 'a+') as f:
                f.write('Convert Document Error\n%s\n%s\n' % (docFile, str(e)))
            log('Convert Document Error {%s} %s' % (docFile, str(e)))
        log('Convert Document {%s} Success' % tempPath)
    log('Preprocessor Document Success')
    # 检查结果文件是否存在,如果存在则删除
    if os.path.exists(resultPath):
        os.remove(resultPath)
    # 检测
    for root, dirs, files in os.walk(cachePath, False):
        for name in files:
            cacheFile = os.path.join(root, name)
            # 读取文件内容
            try:
                with open(cacheFile, 'r') as f:
                    path = f.readline().strip()
                    body = f.read()
                hits = []
                log('Check Document {%s} ...' % utf2gbk(path))
                for word in wordList:
                   if body.find(word) != -1:
                       hits.append(word)
                if len(hits) > 0:
                    with open(resultPath, 'a+') as f:
                        f.write('%s:%s:%s\n' % (path, len(hits), ','.join(hits)))
                    log(utf2gbk('%s:%s:%s' % (path, len(hits), ','.join(hits))))
            except Exception as e:
                with open(errorPath, 'a+') as f:
                    f.write('Check Document Error\n%s\n%s\n' % (cacheFile, str(e)))
        log('Application Finished')

if __name__ == '__main__':
    try:
        op = OptionParser()
        op.add_option('-d', '--dir', type='string', dest='dir', default='', help='Directory Path')
        op.add_option('-w', '--word', type='string', dest='word', default='', help='Word File Path')
        op.add_option('-f', '--flush', type='int', dest='flush', default=0, help='Clear Cache File')
        arg, _ = op.parse_args(sys.argv)
        main(arg.dir, arg.word, arg.flush)
    except OptParseError as e:
        log('Option Parser Error')
    except Exception as e:
        log(str(e))