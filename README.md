# attendance
#考勤信息分析软件
#依赖的包有：os、sys、PyQt5、pandas、datetime
系统要求config文件需要与该系统放在同一目录下。有人机交互UI，在UI总可以选择输入文件和输出目录，并且两个路径都支持中英文混合。
可以多次点击运行，并且每次点击运行生成文件按输入文件名+生成时间+文件内容的格式进行命名。
系统执行的反馈信息输出在交互界面的提示框内。
想要生成exe文件需要pyinstaller工具进行生成，-F：产生单个的可执行文件。-w，指定程序运行时不显示命令行窗口（仅对 Windows 有效）。-c，指定使用命令行窗口运行程序（仅对 Windows 有效）
