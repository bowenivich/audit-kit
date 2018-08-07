# Audit Kit

由于通过VBA爬客编表选定内容的效率较低，导致公司电脑在爬表过程中死机而影响工作，此audit-kit基于Python的应用程序根据`ABASToolkit.xlam`爬表思路编写。慢慢更新，如有Python脚本的需求 :+1: 以及改进意见 :-1:，请联系作者[@alfredbowenfeng](https://github.com/alfredbowenfeng)，邮箱是 alfred.bowenfeng@gmail.com。

若仅爬取个位数量的客编表，使用`ABASToolkit.xlam`即可。若爬取上百份客编表，`ABASToolkit.xlam`的效率较低，此时可使用Python爬取。由于程序中应用了多进程，因此效率会比VBA高许多。

## 下载地址
请在本Repository的Releases，即[此处](https://github.com/alfredbowenfeng/audit-kit/releases)下载所需要的文件。

## 使用

### Breakdown by Value
打开.exe应用程序会首先打开终端，终端会打开一个窗口。在使用的过程中，请勿关闭终端。程序执行结束后，在"Execute"按钮下会出现"Completed"或"Error"以表示此程序是否正常运行完成。
![Breakdown By Value](/sources/breakdown-by-value.png)

在点击"Execute"按钮前，请先输入各项配置。例如：
```
C:\Users\Alfred.Feng\Audit\Engagement\ClientName\PBC\
E1_应收账款
A
7
AJ
1001
C:\Users\Alfred.Feng\Desktop\result.xlsx
```

### Breakdown by Link
打开.exe应用程序会首先打开终端，终端会打开一个窗口。在使用的过程中，请勿关闭终端。程序执行后，在"Execute"按钮下会出现"Completed"或"Error"以表示此程序是否正常运行完成。
![Breakdown By Link](/sources/breakdown-by-link.png)

在点击"Execute"按钮前，请先输入各项配置。例如：
```
C:\Users\Alfred.Feng\Audit\Engagement\ClientName\PBC\
U6_营业外收支
B
61
C
83
C:\Users\Alfred.Feng\Desktop\result.xlsx
```

## 未来更新内容
- 同名文件覆盖提示：目前"Result Path"重名时会直接覆盖旧文件。
- 输入错误提示：目前输入错误的结果为直接报错，不会提示错误原因。
- 大小写支持：目前"Start Column"与"End Column"仅支持大写英文字母。