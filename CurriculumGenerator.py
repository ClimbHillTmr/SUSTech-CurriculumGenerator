#!/usr/bin/env python
# -*- coding: UTF-8 -*-
import openpyxl, sys, Curriculum, re, datetime, requests, json

__course_start_time = {
    "1" : datetime.time(8),
    "3" : datetime.time(10, 20),
    "5" : datetime.time(14),
    "7" : datetime.time(16, 20),
    "9" : datetime.time(19)
    }
__course_end_time = {
    "2" : datetime.time(9, 50),
    "4" : datetime.time(12, 10),
    "6" : datetime.time(15, 50),
    "8" : datetime.time(18, 10),
    "10" : datetime.time(20, 50),
    "11" : datetime.time(21, 50)
    }

def usage():
    print("使用方法：CurriculumGenerator.py <xlsx课程表> <学期开始时间> <学期结束时间> [路程时间提醒(分钟)]")
    print("学期开始和结束时间格式：yyyymmdd。例如：20150101。")
    print("路程时间提醒：可选参数，默认为30分钟，表示课前多少分钟提醒出发。")
    print("请注意：学期开始时间为学期第一周的周一，学期结束时间为学期最后一周的周末。")
    print("请自行调整调休相关课程时间安排，以及国庆周课程安排。")
    sys.exit(-1)

# 检查参数数量，允许3-4个参数
if len(sys.argv) < 4 or len(sys.argv) > 5:
    usage()
if len(sys.argv[1]) < 6 or not sys.argv[1].endswith(".xlsx"):
    usage()
if len(sys.argv[2]) != 8 or len(sys.argv[3]) != 8:
    usage()

# 设置默认路程时间提醒（分钟）
default_travel_time = 30
if len(sys.argv) == 5:
    try:
        default_travel_time = int(sys.argv[4])
        if default_travel_time < 0:
            print("错误：路程时间提醒必须是正整数")
            usage()
    except ValueError:
        print("错误：路程时间提醒必须是整数")
        usage()
        
# 获取中国大陆节假日信息
def get_holidays(start_date, end_date):
    print("正在获取节假日信息...")
    holidays = []
    try:
        # 格式化日期为YYYYMMDD
        start_str = start_date.strftime("%Y%m%d")
        end_str = end_date.strftime("%Y%m%d")
        
        # 使用提莫的神秘小站API获取节假日信息
        url = f"https://timor.tech/api/holiday/range/{start_str}/{end_str}"
        response = requests.get(url, timeout=10)
        data = response.json()
        
        if data["code"] == 0:  # 请求成功
            holiday_data = data["holiday"]
            for date_str, info in holiday_data.items():
                if info["holiday"]:
                    # 将YYYY-MM-DD格式转换为datetime.date对象
                    year, month, day = map(int, date_str.split("-"))
                    holidays.append(datetime.date(year, month, day))
            print(f"成功获取到{len(holidays)}个节假日信息")
        else:
            print(f"获取节假日信息失败: {data['msg']}")
            print("将使用默认节假日信息")
            # 使用默认的节假日信息（当前年份）
            current_year = datetime.datetime.now().year
            holidays = get_default_holidays(current_year)
    except Exception as e:
        print(f"获取节假日信息出错: {str(e)}")
        print("将使用默认节假日信息")
        # 使用默认的节假日信息（当前年份）
        current_year = datetime.datetime.now().year
        holidays = get_default_holidays(current_year)
    
    return holidays

# 获取默认节假日信息（当API请求失败时使用）
def get_default_holidays(year):
    # 根据年份返回预设的节假日列表
    if year == 2025:
        return [
            # 元旦
            datetime.date(2025, 1, 1),
            # 春节
            datetime.date(2025, 1, 29), datetime.date(2025, 1, 30), datetime.date(2025, 1, 31),
            datetime.date(2025, 2, 1), datetime.date(2025, 2, 2), datetime.date(2025, 2, 3), datetime.date(2025, 2, 4),
            # 清明节
            datetime.date(2025, 4, 5), datetime.date(2025, 4, 6), datetime.date(2025, 4, 7),
            # 劳动节
            datetime.date(2025, 5, 1), datetime.date(2025, 5, 2), datetime.date(2025, 5, 3), datetime.date(2025, 5, 4), datetime.date(2025, 5, 5),
            # 端午节
            datetime.date(2025, 6, 2), datetime.date(2025, 6, 3), datetime.date(2025, 6, 4),
            # 中秋节
            datetime.date(2025, 9, 13), datetime.date(2025, 9, 14), datetime.date(2025, 9, 15),
            # 国庆节
            datetime.date(2025, 10, 1), datetime.date(2025, 10, 2), datetime.date(2025, 10, 3),
            datetime.date(2025, 10, 4), datetime.date(2025, 10, 5), datetime.date(2025, 10, 6), datetime.date(2025, 10, 7)
        ]
    elif year == 2024:
        return [
            # 元旦
            datetime.date(2024, 1, 1),
            # 春节
            datetime.date(2024, 2, 10), datetime.date(2024, 2, 11), datetime.date(2024, 2, 12),
            datetime.date(2024, 2, 13), datetime.date(2024, 2, 14), datetime.date(2024, 2, 15), datetime.date(2024, 2, 16), datetime.date(2024, 2, 17),
            # 清明节
            datetime.date(2024, 4, 4), datetime.date(2024, 4, 5), datetime.date(2024, 4, 6),
            # 劳动节
            datetime.date(2024, 5, 1), datetime.date(2024, 5, 2), datetime.date(2024, 5, 3), datetime.date(2024, 5, 4), datetime.date(2024, 5, 5),
            # 端午节
            datetime.date(2024, 6, 10),
            # 中秋节
            datetime.date(2024, 9, 15), datetime.date(2024, 9, 16), datetime.date(2024, 9, 17),
            # 国庆节
            datetime.date(2024, 10, 1), datetime.date(2024, 10, 2), datetime.date(2024, 10, 3),
            datetime.date(2024, 10, 4), datetime.date(2024, 10, 5), datetime.date(2024, 10, 6), datetime.date(2024, 10, 7)
        ]
    else:
        print(f"警告：没有{year}年的预设节假日信息，将使用2024年的节假日信息")
        return get_default_holidays(2024)
workbook = openpyxl.load_workbook(sys.argv[1])
sheet = workbook.active
term_start = sys.argv[2]
term_start_date = datetime.date(int(term_start[0:4]), int(term_start[4:6]), int(term_start[6:]))
term_end = sys.argv[3]
term_end_date = datetime.date(int(term_end[0:4]), int(term_end[4:6]), int(term_end[6:]))

# 获取节假日信息
holidays = get_holidays(term_start_date, term_end_date)

curriculum = Curriculum.Curriculum()
day = 0
for col in sheet.iter_cols(min_col=2, max_col=8, min_row=4, max_row=13):
    for course in col:
        if course.value:
            val = course.value.split("\n")
            # 处理特殊格式的课程信息（如实验室安全学）
            if "实验室安全学" in course.value and "[1-8节]" in course.value:
                name = "实验室安全学"
                # 为实验室安全学设置具体地点
                location = "南方科技大学-第一科研楼101"
                num_str = re.findall(r"\d+", sheet["A%d"%course.row].value)
                num = [num_str[0], "8"] if len(num_str) > 0 else ["1", "8"]
                week_num = re.findall(r"\d+", re.findall(r"\[\d+周\]", course.value)[0])[0]
                type = Curriculum.CourseRepetitionType.weekly
                # 计算第12周的日期
                week_start = term_start_date + datetime.timedelta(days=(int(week_num)-1)*7)
                start_time = datetime.datetime.combine(week_start + datetime.timedelta(days=day), __course_start_time.get(num[0]))
                end_time = datetime.datetime.combine(week_start + datetime.timedelta(days=day), __course_end_time.get(num[1]))
                
                # 检查是否为节假日
                course_date = start_time.date()
                
                if course_date not in holidays:
                    term_end_date_dt = datetime.datetime.combine(term_end_date, datetime.time.max)
                    # 为实验室安全学设置较长的路程时间提醒（默认值的2倍）
                    safety_travel_time = default_travel_time * 2
                    Curriculum.add_course(curriculum, name, start_time, end_time, location, type, term_end_date_dt, travel_time_minutes=safety_travel_time)
                continue
                
            # 处理常规格式的课程
            for i in range(len(val)//4):
                name = val[i * 4]
                info = val[i * 4 + 3].split("][")
                num = re.findall(r"\d+", sheet["A%d"%course.row].value)
                if info[0][-2] == '单':
                    type = Curriculum.CourseRepetitionType.biweekly
                    start_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day), __course_start_time.get(num[0]))
                    end_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day), __course_end_time.get(num[1]))
                elif info[0][-2] == '双':
                    type = Curriculum.CourseRepetitionType.biweekly
                    start_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day+7), __course_start_time.get(num[0]))
                    end_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day+7), __course_end_time.get(num[1]))
                else:
                    type = Curriculum.CourseRepetitionType.weekly
                    start_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day), __course_start_time.get(num[0]))
                    end_time = datetime.datetime.combine(term_start_date + datetime.timedelta(days=day), __course_end_time.get(num[1]))
                location = "南方科技大学-" + info[1]
                # 检查是否为节假日
                is_holiday = False
                
                # 使用全局获取的节假日信息
                
                # 将datetime.datetime转换为datetime.date进行比较
                course_date = start_time.date()
                
                # 如果课程日期是节假日，则跳过添加
                if course_date in holidays:
                    continue
                    
                term_end_date_dt = datetime.datetime.combine(term_end_date, datetime.time.max)
                # 根据课程名称设置不同的路程时间提醒
                travel_time = default_travel_time  # 使用命令行参数设置的默认值
                
                # 根据课程名称或地点调整路程时间
                if "实验" in name or "实践" in name:
                    travel_time = int(default_travel_time * 1.5)  # 实验课程需要更多准备时间
                elif "第一科研楼" in location or "荔园" in location:
                    travel_time = int(default_travel_time * 1.33)  # 距离较远的地点
                
                Curriculum.add_course(curriculum, name, start_time, end_time, location, type, term_end_date_dt, travel_time_minutes=travel_time)
    day += 1

curriculum.save_as_ics_file()