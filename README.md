# NEUOJ刷题（C语言）实验报告自动生成器
请使用网站:https://oj.neu.edu.cn/

## 要求
Python3.7,BeautifulSoup4,Python-docx,Selenium,Pillow

## 使用方法
Step 0.配置环境<br>
安装Python3.7，使用pip install安装以上所列出的软件包
推荐使用PyCharm IDE

Step 1.添加个人信息<br>
·用自己姓名替换name
·用自己的刷题所用的用户名替换username的单引号内内容
·用学号替换StuID。
<br>
 
Step 2.运行脚本<br>
运行程序之后,会生成一个在template文件夹中以你的学号命名的docx文档。<br>
docx文档第一页为封面，后面会生成自动填写了学号，姓名，题号，代码页的正文，
只需将代码运行截图插入即可。<br>
