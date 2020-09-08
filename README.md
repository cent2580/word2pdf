# word2pdf

lconvert --- linux环境下转换格式脚本
command： python lconvert.py -f "format" filepath
Dependency: pywpsrpc

wconvert --- windows环境下word转pdf脚本
command: python wconvert.py filepath -n newfilename -t topath(save path)
Dependency: pypiwin32

lib/libstd++.so.6.0.26 --- centos环境下缺少的编译文件
usage：
查看命令库：
strings /usr/lib64/libstdc++.so.6 | grep GLIBC
strings /usr/lib64/libstdc++.so.6 | grep CXXABI
1.把libstdc++.so.6拷贝到/usr/lib64目录下，cp libstdc++.so.6 /usr/lib64；
2.删除原来的libstdc++.so.6符号链接，rm -rf libstdc++.so.6；
3.新建新的符号链接， In -s libstdc++.so.6.0.24 libstdc++.so.6。