# 代码归属与声明

本项目是基于 [dazhi0619/CurriculumGenerator](https://github.com/dazhi0619/CurriculumGenerator) 的代码库进行开发的。原始代码库使用了 MIT 许可证，我们在此保留所有原始作者的版权声明和许可条款。

## 原始项目信息

- 原始项目名称：SUSTech iCal课表生成工具
- 原始作者：[dazhi0619](https://github.com/dazhi0619)
- 原始代码库：https://github.com/dazhi0619/CurriculumGenerator

## 修改说明

本项目在原始代码基础上进行了以下改进：

- 添加了路程时间提醒功能，使用VALARM组件在日历中设置提醒
- 添加了命令行参数支持自定义路程时间提醒的基准值（默认30分钟）
- 添加了自动避开假期功能，通过网络API获取中国大陆法定节假日信息，避免在假期安排课程

## 许可证

本项目继续遵循 MIT 许可证。详情请参阅 [LICENSE](LICENSE) 文件。

## 致谢

感谢原始项目作者 [dazhi0619](https://github.com/dazhi0619) 提供的基础代码。
同时，正如原始项目所述，部分代码使用了 @zhongbr 的[代码](https://www.cnblogs.com/zhongbr/p/python_calender.html)，在此一并表示感谢。