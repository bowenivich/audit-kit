# Audit Kit
一些Audit小工具，代码可直接下载，具有图形界面的App可在[发表页面](https://github.com/alfredbowenfeng/audit-kit/releases)下载。慢慢更新，如有Python脚本的需求 :+1: 以及改进意见 :-1:，请联系作者[@alfredbowenfeng](https://github.com/alfredbowenfeng)，邮箱是[alfred.bowenfeng@gmail.com](mailto:alfred.bowenfeng@gmail.com)。

## Breakdown 爬取
由于通过VBA爬客编表选定内容的效率较低，导致公司电脑在爬表过程中死机而影响工作，此audit-kit基于Python的App根据`ABASToolkit.xlam`爬表思路编写，在爬取上百份客编表中的部分内容时，效率将远高于VBA。

### Breakdown by Value
- 1. 在程序中已设立对指定的文件夹中所有文件的3个筛选条件。
	- 以`.xlsx`结尾的文件
	- 不以`._`开头的文件（通常为恢复文件）
	- 不以`~$`开头的文件（通常为active workbook的临时文件）

- 2. 请仍确保指定的文件夹中不存在例外情况（例如2个重复文件等）。 

- 3. 以运行文件，可直接双击.exe的应用程序。
![Breakdown By Value](/sources/README/breakdown-value-ui.png)

- 4. 请在文本框输入各项配置。例如：
```
C:\Users\Alfred.Feng\Audit\Engagement\ClientName\PBC\
E1_应收账款
A
7
AJ
1001
C:\Users\Alfred.Feng\Desktop\result.xlsx
```

- 5. 程序执行结束后，在"Execute"按钮下方会出现"Completed"或"Error"以表示此程序是否正常运行完成（通常运行完成可能会耗费些时间）。

### Breakdown by Link
- 1. 在程序中已设立对指定的文件夹中所有文件的3个筛选条件。
	- 以`.xlsx`结尾的文件
	- 不以`._`开头的文件（通常为恢复文件）
	- 不以`~$`开头的文件（通常为active workbook的临时文件）

- 2. 请仍确保指定的文件夹中不存在例外情况（例如2个重复文件等）。 

- 3. 以运行文件，可直接双击.exe的应用程序。
![Breakdown By Link](/sources/README/breakdown-link-ui.png)

- 4. 请在文本框输入各项配置。例如：
```
C:\Users\Alfred.Feng\Audit\Engagement\ClientName\PBC\
U6_营业外收支
B
61
C
83
C:\Users\Alfred.Feng\Desktop\result.xlsx
```

- 5. 程序执行结束后，在"Execute"按钮下方会出现"Completed"或"Error"以表示此程序是否正常运行完成（通常此过程可极快得完成）。

- 6. 在生成的.xlsx文件中，如果出现`#REF`的现象，复制、以Transpose模式粘贴即可生效。若对此步骤有疑问，请直接联系[Alfred Feng](mailto:alfred.bowenfeng@gmail.com)。

### Future Updates
- 同名目标文件覆盖提示：目前"Result Path"重名时会直接覆盖旧文件。
- 输入错误提示：目前输入错误的结果为直接报错，不会提示错误原因。
- 大小写支持：目前"Start Column"与"End Column"仅支持大写英文字母。

## 增值税发票识别信息

### VAT Invoices Scanning
- 1. 在程序中未对指定的文件夹中所有文件设立任何筛选条件，因此在运行前，请确保指定的文件夹中不存在例外情况（例如可能存在非图片格式的文件）。

- 2. 此程序将严格按照文件名来排序，生成的图片信息排序与文件名排序严格一致。iOS的相机会自动对图片编号（例如"IMG_1001.JPG"），可直接全部放入文件夹。

- 3. 以运行文件，可直接双击.exe的应用程序。
![VAT Invoices Scanning](/sources/README/vat-invoices-scanning-ui.png)

- 4. 请在文本框输入各项配置。例如：
```
C:\Users\Alfred.Feng\Audit\Engagement\ClientName\Invoices\
C:\Users\Alfred.Feng\Desktop\result.xlsx
```

- 5. 程序执行结束后，在"Execute"按钮下方会出现"Completed"或"Error"以表示此程序是否正常运行完成。

### Future Updates
- 对指定的文件夹中的所有文件设立筛选条件。
- 可选择输出内容（似乎没用，毕竟新得到的）。
- 对文件实现自动编号以完成cross-reference。

## 未来将开展的项目
- 运用LATEX，根据.xlsx文件中的内容，自动生成律师询证函及相关底稿。
- ......