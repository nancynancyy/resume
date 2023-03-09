import os
import ttkbootstrap as tk
from docxtpl import DocxTemplate
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame
import convert
def main():
    root = tk.Window(themename='litera')
    root.geometry('600x500+500+500')
    root.title('简历生成器')
    #root.wm_attributes('-topmost', 1)  #保持窗口在最前
    #图标
    root.iconbitmap('./resume.ico')
    #滚动条？
    # sf=ScrolledFrame(root,autohide=True)
    # sf.grid(column=4,sticky=tk.W,padx=10,pady=10)
    #数据库
    name_str_var = tk.StringVar()
    # 0 女 1 男
    gender_str_var = tk.StringVar()
    phone_str_var = tk.StringVar()
    address_str_var = tk.StringVar()
    #data_entry1

    # 教育背景
    school_str_var = tk.StringVar()
    major_str_var=tk.StringVar()
    gpa_str_var=tk.StringVar()
    course_str_var=tk.StringVar()

    # 校园经历
    job1_str_var=tk.StringVar()
    work1_str_var=tk.StringVar()

    # 实习经历
    job2_str_var=tk.StringVar()
    work2_str_var=tk.StringVar()

    # 个人评价
    skill_str_var=tk.StringVar()
    assessment_str_var=tk.StringVar()

    #导出
    # 0 否 1 是
    word_str_var = tk.IntVar()
    pdf_str_var=tk.IntVar()

    # 所获证书
    cert_list = [
        [tk.IntVar(), '国家奖学金'],
        [tk.IntVar(), '国家励志奖学金'],
        [tk.IntVar(), '郑州大学一等奖学金'],
        [tk.IntVar(), '郑州大学二等奖学金'],
        [tk.IntVar(), '郑州大学三等奖学金']
    ]

    # 个人信息
    tk.Label(root, width=5).grid()
    tk.Label(root, text='姓名：').grid(row=1, column=1, sticky=tk.W, pady=10)
    tk.Entry(root, textvariable=name_str_var).grid(row=1, column=2, sticky=tk.W)
    #按钮：上传简历照片
    # button3 = tk.Button(root, text='上传简历照片', width=20)
    # button3.grid(row=1, column=3, sticky=tk.E)
    # button3.config(command=photo.choose_file)

    tk.Label(root, text='联系方式：').grid(row=3, column=2, sticky=tk.E, pady=20)
    tk.Entry(root, textvariable=phone_str_var).grid(row=3, column=3, sticky=tk.W)
    tk.Label(root, text='所在城市：').grid(row=3, column=3, sticky=tk.E, pady=10)
    tk.Entry(root, textvariable=address_str_var).grid(row=3, column=4, sticky=tk.W)

    # 性别 单选框
    tk.Label(root, text='性别：').grid(row=1, column=2, sticky=tk.E, pady=10)
    radio_frame = tk.Frame()
    radio_frame.grid(row=1, column=3, sticky=tk.W)
    tk.Radiobutton(radio_frame, text='男', variable=gender_str_var, value='男').pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(radio_frame, text='女', variable=gender_str_var, value='女').pack(side=tk.LEFT, padx=5)
    # 出生日期
    tk.Label(root, text='出生日期：').grid(row=3, column=1, sticky=tk.W, pady=10)
    data_entry1 = tk.DateEntry()
    data_entry1.grid(row=3, column=2, sticky=tk.W, pady=10)


    # 所获证书
    tk.Label(root, text='所获证书：').grid(row=6, column=1, sticky=tk.W, pady=10)
    check_frame = tk.Frame()
    check_frame.grid(row=6, column=2, sticky=tk.W)
    tk.Checkbutton(check_frame, text=cert_list[0][1], variable=cert_list[0][0]).pack(side=tk.LEFT, padx=5)
    tk.Checkbutton(check_frame, text=cert_list[1][1], variable=cert_list[1][0]).pack(side=tk.LEFT, padx=5)    #第一个括号加入bootstyle="square-toggle"，为方形复选框
    tk.Checkbutton(check_frame, text=cert_list[2][1], variable=cert_list[2][0]).pack(side=tk.LEFT, padx=5)    # bootstyle="round-toggle"为圆形复选框
    tk.Checkbutton(check_frame, text=cert_list[3][1], variable=cert_list[3][0]).pack(side=tk.LEFT, padx=5)
    tk.Checkbutton(check_frame, text=cert_list[4][1], variable=cert_list[4][0]).pack(side=tk.LEFT, padx=5)

    # 教育背景
    tk.Label(root,text='教育背景',background='pink').grid(row=10,column=2,sticky=tk.W,pady=10)  #E靠右,W左,sticky选项去指定对齐方式，可以选择的值有：N/S/E/W，分别代表上对齐/下对齐/左对齐/右对齐,padx=10表示距离左右两边组件的长度都为10
    #教育背景-学校名称
    tk.Label(root, text='学校名称：').grid(row=11, column=1, sticky=tk.W, pady=10)
    tk.Entry(root, textvariable=school_str_var).grid(row=11, column=2, sticky=tk.W)
    #教育背景-专业
    tk.Label(root, text='专业：').grid(row=11, column=2, sticky=tk.E, pady=10)
    tk.Entry(root, textvariable=major_str_var).grid(row=11, column=3, sticky=tk.W)
    #教育背景-开始时间、结束时间
    tk.Label(root, text='开始时间：').grid(row=13, column=1, sticky=tk.W, pady=10)
    data_entry2 = tk.DateEntry()
    data_entry2.grid(row=13, column=2, sticky=tk.W, pady=10)

    tk.Label(root, text='结束时间：').grid(row=13, column=2, sticky=tk.E, pady=10)
    data_entry3 = tk.DateEntry()
    data_entry3.grid(row=13, column=3, sticky=tk.W, pady=10)
    #教育背景-GPA
    tk.Label(root, text='GPA：').grid(row=15, column=1, sticky=tk.W, pady=10)
    tk.Entry(root, textvariable=gpa_str_var).grid(row=15, column=2, sticky=tk.W)
    #教育背景-主修课程
    tk.Label(root, text='主修课程：').grid(row=16, column=1, sticky=tk.W, pady=10)
    tk.Entry(root, textvariable=course_str_var,width=70).grid(row=16, column=2, sticky=tk.W)

    # 校园经历
    tk.Label(root,text='校园经历',background='pink').grid(row=20,column=2,sticky=tk.W,pady=10)
    #校园经历-职务名称
    tk.Label(root, text='职务名称：').grid(row=21, column=1, sticky=tk.W, pady=10)
    tk.Entry(root, textvariable=job1_str_var).grid(row=21, column=2, sticky=tk.W)
    #校园经历-开始时间、结束时间
    tk.Label(root, text='开始时间：').grid(row=22, column=1, sticky=tk.W, pady=10)
    data_entry4 = tk.DateEntry()
    data_entry4.grid(row=22, column=2, sticky=tk.W, pady=10)
    #print(data_entry2.entry.get())
    tk.Label(root, text='结束时间：').grid(row=22, column=2, sticky=tk.E, pady=10)
    data_entry5 = tk.DateEntry()
    data_entry5.grid(row=22, column=3, sticky=tk.W, pady=10)
    #校园经历-职责
    tk.Label(root, text='职责：').grid(row=24, column=1, sticky=tk.W, pady=10)
    tk.Entry(root, textvariable=work1_str_var,width=70).grid(row=24, column=2, sticky=tk.W)

    # 实习经历
    tk.Label(root,text='实习经历',background='pink').grid(row=30,column=2,sticky=tk.W,pady=10)
    #实习经历-职务名称
    tk.Label(root, text='职务名称：').grid(row=31, column=1, sticky=tk.W, pady=10)
    tk.Entry(root, textvariable=job2_str_var).grid(row=31, column=2, sticky=tk.W)
    #实习经历-开始时间、结束时间
    tk.Label(root, text='开始时间：').grid(row=32, column=1, sticky=tk.W, pady=10)
    data_entry6 = tk.DateEntry()
    data_entry6.grid(row=32, column=2, sticky=tk.W, pady=10)

    tk.Label(root, text='结束时间：').grid(row=32, column=2, sticky=tk.E, pady=10)
    data_entry7 = tk.DateEntry()
    data_entry7.grid(row=32, column=3, sticky=tk.W, pady=10)
    #实习经历-职责
    tk.Label(root, text='职责：').grid(row=34, column=1, sticky=tk.W, pady=10)
    tk.Entry(root, textvariable=work2_str_var,width=70).grid(row=34, column=2, sticky=tk.W)


    # 个人评价
    tk.Label(root,text='个人评价',background='pink').grid(row=40,column=2,sticky=tk.W,pady=10)
    #个人评价-技能特长
    tk.Label(root, text='技能特长：').grid(row=41, column=1, sticky=tk.W, pady=10)
    tk.Entry(root, textvariable=skill_str_var,width=70).grid(row=41, column=2, sticky=tk.W)
    #个人评价-自我评价
    tk.Label(root, text='自我评价：').grid(row=41, column=2, sticky=tk.E, pady=10)
    tk.Entry(root, textvariable=assessment_str_var,width=50).grid(row=41, column=3, sticky=tk.W)


    #提交并转换为docx文件
    tk.Label(root, text="").grid(row=50, column=2, sticky=tk.W)
    button = tk.Button(root, text='提交', width=20)
    button.grid(row=50, column=2, sticky=tk.W)
    def toWord():
        if os.path.isfile('output\out.docx'):  #判断是否存在
            os.remove('output\out.docx') # 删除文件
        if os.path.isfile(r'output\out.pdf'):
            os.remove(r'output\out.pdf')
        doc = DocxTemplate("templates/001.docx")
        context = {
            'name': name_str_var.get(),
            'gender': gender_str_var.get(),
            'birthday': data_entry1.entry.get(),
            'phone': phone_str_var.get(),
            'address': address_str_var.get(),
            'certificate': [h for v, h in cert_list if v.get()],
            'university': school_str_var.get(),
            'major': major_str_var.get(),
            'start_time1': data_entry2.entry.get(),
            'end_time1': data_entry3.entry.get(),
            'gpa': gpa_str_var.get(),
            'course': course_str_var.get(),
            'campus_job': job1_str_var.get(),
            'start_time2': data_entry4.entry.get(),
            'end_time2': data_entry5.entry.get(),
            'campus_work': work1_str_var.get(),
            'internship_job': job2_str_var.get(),
            'start_time3': data_entry6.entry.get(),
            'end_time3': data_entry7.entry.get(),
            'internship_work': work2_str_var.get(),
            'skill': skill_str_var.get(),
            'self_assessment': assessment_str_var.get()

        }
        doc.render(context)
        doc.save("output/out.docx")
        print("Successfully generated out.docx!")


    button.config(command=toWord)


    #pdf转换
    button2 = tk.Button(root, text='转换为pdf', width=20)
    button2.grid(row=50, column=3, sticky=tk.W)

    button2.config(command=convert.convert_to_pdf)


    #在运行框显示输入结果
    def get_info():
        data = {
            '姓名': name_str_var.get(),
            '性别': gender_str_var.get(),
            '出生日期': data_entry1.entry.get(),
            '联系方式': phone_str_var.get(),
            '所在城市': address_str_var.get(),
            '所获证书': [h for v, h in cert_list if v.get()],
            '教育经历': [school_str_var.get(), major_str_var.get(), data_entry2.entry.get(), data_entry3.entry.get(),
                     gpa_str_var.get(), course_str_var.get()],
            '校园经历': [job1_str_var.get(), data_entry4.entry.get(), data_entry5.entry.get(), work1_str_var.get()],
            '实习经历': [job2_str_var.get(), data_entry6.entry.get(), data_entry7.entry.get(), work2_str_var.get()],
            '个人评价': [skill_str_var.get(), assessment_str_var.get()]
        }
        print(data)
        with open('1.txt', mode='a') as f:
            f.write('\n')
            f.write(str(data))



    root.mainloop()

if __name__=='__main__':
    main()