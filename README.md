# 搜索指定目录中包含指定关键词的word文档

> 由于采用`win32com`,所以仅支持`windows`系统
> `windows`系统必须安装`office word`软件
> 需要`python 2.7.x`环境支持
> 需要`pywin32`库支持

## 参数说明

```
python main.py -h
```

- `-h`: 帮助信息
- `-d`: 指定要搜索的目录,支持递归检索,请使用绝对路径
- `-w`: 指定要搜索的关键词文件,每行一个关键词,`#`表示注释,区分大小写
- `-f`: 是否清理上次执行遗留的缓存.如果取值为`0`,那么搜索出错中断,可以继续接上次进度执行.如果取值为`1`,那么将会清理上次执行的数据.

## 演示

命令
```
python main.py -d d:/FindDocs/test -w words.txt -f
```

输出内容
```
[2016-05-25 11:29:48] Application Starting ...
[2016-05-25 11:29:48] Loading Word File ...
[2016-05-25 11:29:48] Word File Load Success, Find 4 Word
[2016-05-25 11:29:48] Find Doc File ...
[2016-05-25 11:29:48] Find Success, Find 8 Document
[2016-05-25 11:29:48] Preprocessor Document ...
[2016-05-25 11:29:48] Convert Document D:\FindDocs\test\folder\folder\empty.docx
[2016-05-25 11:29:48] Convert Document {D:\FindDocs\doc.tmp} Success
[2016-05-25 11:29:48] Convert Document D:\FindDocs\test\folder\folder\文档 - 副本.doc
[2016-05-25 11:29:51] Convert Document {D:\FindDocs\doc.tmp} Success
[2016-05-25 11:29:51] Convert Document D:\FindDocs\test\folder\folder\文档 - 副本.docx
[2016-05-25 11:29:54] Convert Document {D:\FindDocs\doc.tmp} Success
[2016-05-25 11:29:54] Convert Document D:\FindDocs\test\folder\folder\文档.docx
[2016-05-25 11:29:56] Convert Document {D:\FindDocs\doc.tmp} Success
[2016-05-25 11:29:56] Convert Document D:\FindDocs\test\empty.docx
[2016-05-25 11:29:58] Convert Document {D:\FindDocs\doc.tmp} Success
[2016-05-25 11:29:58] Convert Document D:\FindDocs\test\文档 - 副本.doc
[2016-05-25 11:30:02] Convert Document {D:\FindDocs\doc.tmp} Success
[2016-05-25 11:30:02] Convert Document D:\FindDocs\test\文档 - 副本.docx
[2016-05-25 11:30:06] Convert Document {D:\FindDocs\doc.tmp} Success
[2016-05-25 11:30:06] Convert Document D:\FindDocs\test\文档.docx
[2016-05-25 11:30:09] Convert Document {D:\FindDocs\doc.tmp} Success
[2016-05-25 11:30:09] Preprocessor Document Success
[2016-05-25 11:30:09] Check Document {D:\FindDocs\test\folder\folder\empty.docx} ...
[2016-05-25 11:30:09] Check Document {D:\FindDocs\test\folder\folder\文档 - 副本.doc} ...
[2016-05-25 11:30:09] D:\FindDocs\test\folder\folder\文档 - 副本.doc:2:关键词,$#%
[2016-05-25 11:30:09] Check Document {D:\FindDocs\test\folder\folder\文档 - 副本.docx} ...
[2016-05-25 11:30:09] D:\FindDocs\test\folder\folder\文档 - 副本.docx:2:关键词,$#%
[2016-05-25 11:30:09] Check Document {D:\FindDocs\test\文档 - 副本.doc} ...
[2016-05-25 11:30:09] D:\FindDocs\test\文档 - 副本.doc:2:关键词,$#%
[2016-05-25 11:30:09] Check Document {D:\FindDocs\test\folder\folder\文档.docx} ...
[2016-05-25 11:30:09] D:\FindDocs\test\folder\folder\文档.docx:2:关键词,$#%
[2016-05-25 11:30:09] Check Document {D:\FindDocs\test\文档.docx} ...
[2016-05-25 11:30:09] D:\FindDocs\test\文档.docx:2:关键词,$#%
[2016-05-25 11:30:09] Check Document {D:\FindDocs\test\文档 - 副本.docx} ...
[2016-05-25 11:30:09] D:\FindDocs\test\文档 - 副本.docx:2:关键词,$#%
[2016-05-25 11:30:09] Check Document {D:\FindDocs\test\empty.docx} ...
[2016-05-25 11:30:09] Application Finished
```

结果记录文件`result.txt`内容
```
D:\FindDocs\test\folder\folder\文档 - 副本.doc:2:关键词,$#%
D:\test\folder\folder\文档 - 副本.docx:2:关键词,$#%
D:\FindDocs\test\文档 - 副本.doc:2:关键词,$#%
D:\FindDocs\test\folder\folder\文档.docx:2:关键词,$#%
D:\FindDocs\test\文档.docx:2:关键词,$#%
D:\FindDocs\test\文档 - 副本.docx:2:关键词,$#%
```


