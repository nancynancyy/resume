import pathlib

import win32com
from win32com import client


def convert_to_pdf_libre():
    # 调用程序
    oo_service_manager = win32com.client.DispatchEx("com.sun.star.ServiceManager")
    desktop = oo_service_manager.CreateInstance("com.sun.star.frame.Desktop")
    oo_service_manager._FlagAsMethod("Bridge_GetStruct")

    # 生成参数元组
    def create_prop(name, value):
        prop = oo_service_manager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        prop.Name = name
        prop.Value = value
        return prop

    # 读取文档的参数
    loading_properties = [create_prop("ReadOnly", True), create_prop("Hidden", True)]
    path = pathlib.PurePath(pathlib.Path.cwd(), 'output', 'out.docx').as_posix()
    filepath = "file:///%s" % path

    document = desktop.loadComponentFromUrl(filepath, "_blank", 0, tuple(loading_properties))
    document.CurrentController.Frame.ContainerWindow.Visible = False
    # 转换为pdf的参数
    pdf_properties = [create_prop("FilterName", "writer_pdf_Export")]

    newpath = filepath[:-len("docx")] + "pdf"  # pdf 保存路径和名称

    try:
        document.storeToURL(newpath, pdf_properties)  # 转换输出
    except Exception:
        raise
    try:
        document.close(True)  # 关闭
    except Exception:
        raise


def convert_to_pdf_wps():
    path = pathlib.PurePath(pathlib.Path.cwd(), 'output', 'out.docx').as_posix()
    pdf_path = path[:-len("docx")] + "pdf"  # pdf 保存路径和名称
    o = win32com.client.DispatchEx("Kwps.Application")
    o.Visible = False
    doc = o.Documents.Open(path)
    doc.ExportAsFixedFormat(pdf_path, 17)
    o.Quit()


def convert_to_pdf_ms():
    path = pathlib.PurePath(pathlib.Path.cwd(), 'output', 'out.docx').as_posix()
    pdf_path = path[:-len("docx")] + "pdf"  # pdf 保存路径和名称
    word = client.DispatchEx("Word.Application")  # 打开word应用程序
    # for file in files:
    doc = word.Documents.Open(path)  # 打开word文件
    doc.ExportAsFixedFormat(pdf_path, 17)  # 另存为后缀为".pdf"的文件，其中参数17表示为pdf
    doc.Close()  # 关闭原来word文件
    word.Quit()


def convert_to_pdf():
    try:
     convert_to_pdf_ms()
    except Exception:
        print("no ms office")
    else:
        print("Successfully used ms office convert to pdf!")
        return
    try:
        convert_to_pdf_libre()
    except Exception:
        print("no libre office")
    else:
        print("Successfully used libre office convert to pdf!")
        return
    try:
        convert_to_pdf_wps()
    except Exception:
        print("no wps office")
    else:
        print("Successfully used wps office convert to pdf!")
        return

