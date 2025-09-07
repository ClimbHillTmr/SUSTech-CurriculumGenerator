# SUSTech iCal课表生成工具

这个脚本可以帮你把从[南科大教务系统](https://tis.sustech.edu.cn)下载的xlsx格式的课表转化为可以导入日历软件通用的iCal（.ics）格式的课表。

> **注意**：本项目是基于 [dazhi0619/CurriculumGenerator](https://github.com/dazhi0619/CurriculumGenerator) 的代码库进行开发的，增加了路程时间提醒功能。详细信息请查看 [ATTRIBUTION.md](ATTRIBUTION.md)。

## 为什么要生成iCal格式课表

可以导入设备自带日历软件，从而与系统深度整合，例如基本上不占用系统资源实现提醒功能，或通过Siri等语音助手管理课程安排。

<img src="images/Siri-Integration.png" width = "50%" />

## 使用方式

1. 首先在教务系统主界面左侧点击课表，下载xlsx格式课表。
2. 若电脑存在[Python](https://www.python.org)环境及[openpyxl](https://openpyxl.readthedocs.io/en/stable/)，则下载[CurriculumGenerator](https://github.com/dazhi0619/CurriculumGenerator)目录下的两个.py文件。
   如果您不知道我在说什么，请[在此处](https://github.com/dazhi0619/CurriculumGenerator/releases/)下载打包好的可执行文件。
3. 运行

   ```
   python3 CurriculumGenerator.py <xlsx课表文件名> <学期开始日期> <学期结束日期> [路程时间提醒(分钟)]
   ```

   若您下载的是可执行文件，请在下载文件夹中按住shift右键单击文件夹空白处，选择"在此处打开powershell窗口"，在弹出的窗口中输入

   ```
   .\CurriculumGenerator <xlsx课表文件名> <学期开始日期> <学期结束日期> [路程时间提醒(分钟)]
   ```

   其中，路程时间提醒是可选参数，默认为30分钟，表示课前多少分钟提醒出发。
4. 将生成好的 `课表.ics`导入日历软件。通常情况下直接打开即可。对于iPhone和iPad，请将此文件AirDrop到您的设备上，或设法通过Safari浏览器打开此文件。

## 路程时间提醒功能

本项目增加了路程时间提醒功能，会根据课程类型自动设置不同的提醒时间。默认基准时间为30分钟。

例如，如果您设置基准时间为20分钟：`python3 CurriculumGenerator.py export.xlsx 20250908 20251228 20`，则各类课程的提醒时间将相应调整为40分钟、30分钟、27分钟和20分钟。

导入日历后，系统会在设定的时间自动提醒您出发前往上课地点。

## 自动避开假期功能

本项目新增了自动避开假期的功能，会自动从网络获取中国大陆的法定节假日信息，并在生成课表时自动跳过这些日期的课程安排。具体功能包括：

- 自动从网络API获取最新的中国大陆法定节假日信息
- 如果网络请求失败，会使用内置的2024年和2025年节假日信息作为备用
- 在节假日期间的课程将不会被添加到生成的日历中

这样可以避免在法定节假日期间出现错误的课程安排，使生成的课表更加准确。

## 许可证

本项目继续遵循 MIT 许可证。详情请参阅 [LICENSE](LICENSE) 文件。

## 致谢

感谢原始项目作者 [dazhi0619](https://github.com/dazhi0619) 提供的基础代码。
同时，本项目部分使用了 @zhongbr 的[代码](https://www.cnblogs.com/zhongbr/p/python_calender.html)，在此一并表示感谢。
