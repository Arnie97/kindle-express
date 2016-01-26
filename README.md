What
----
这是一个用VBS写的SMTP客户端，功能类似于Amazon官方的Send To Kindle。

Why
---
我不太了解Linux桌面下有没有类似Windows下“发送到”菜单的机制。
而且开源社区应该不缺SMTP客户端，我就不造轮子了。

而对于Windows而言，考虑到Windows XP还在我家老爷机上超期服役，PowerShell就不考虑了。
那么，据我所知，WSH是唯一一种不用装开发环境的脚本。

How
---
这个脚本从命令行参数得到要发送的文件：

    $ cscript express.vbs file_1.mobi file_2.txt ...

当然，所有用命令行参数实现的地方都可以用，例如：

- 把你要发送的文件拖到程序的图标上
- 把程序本身，或者程序的快捷方式放到`shell:sendto`文件夹中
