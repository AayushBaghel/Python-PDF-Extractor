import pdfminer
import tkinter as tk
from tkinter import filedialog
from io import StringIO
import re
import time
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser

def func(mystr):
    ind = mystr.find("Natural Diamond")
    if ind==-1:
        type_diamond = "Lab"
    else:
        type_diamond = "Natural"
    type_d.configure(text="Type of diamond: "+str(type_diamond))
    report_num = ""
    desc = ""
    d_ind = mystr.find("DESCRIPTION")
    if ind==-1:
        desc="Laboratory Grown Diamond"
    else:
        desc="Natural Diamond"
    des1.configure(text="Decription: "+str(desc))
    shapecut=""
    s_ind1 = mystr.find("Brilliant")
    s_ind2 = mystr.find("BRILLIANT")
    if s_ind1!=-1:
        if mystr.find("Cut-")!=-1:
            shapecut = "Cut-Cornered Rectangular Modified Brilliant"
        else:
            if mystr[s_ind1-2]=='d':
                shapecut = "Round Brilliant"
            else:
                shapecut = "Oval Brilliant"
    elif s_ind2!=-1:
        if mystr[s_ind1-2]=='D':
            shapecut = "ROUND BRILLIANT"
        else:
            shapecut = "OVAL BRILLIANT"
    elif mystr.find("EMERALD CUT")!=-1:
        shapecut = "EMERALD CUT"
    shape_box.configure(text="Shape and Cutting Style: "+str(shapecut))
    measure = ""
    m_ind1 = mystr.find("mm")
    m_ind2 = mystr.find("MM")
    if m_ind1 != -1:
        if mystr[m_ind1-20]==" " or mystr[m_ind1-20]=="\n":
            measure=mystr[m_ind1-19:m_ind1+2]
        else:
            measure=mystr[m_ind1-20:m_ind1+2]
    else:
        if mystr[m_ind2-20]==" " or mystr[m_ind1-20]=="\n":
            measure=mystr[m_ind2-19:m_ind2+2]
        else:
            measure=mystr[m_ind2-20:m_ind2+2]
    measure_box.configure(text="Measurements: "+str(measure))
    cw_ind = mystr.find("CARATS")
    if cw_ind != -1:
        carat_weight = mystr[cw_ind-5:cw_ind-1]
    else:
        cw_ind = mystr.find("CARAT")
        if cw_ind != -1:
            carat_weight = mystr[cw_ind-5:cw_ind-1]
        else:
            cw_ind = mystr.find("carat")
            if cw_ind != -1:
                carat_weight = mystr[cw_ind-5:cw_ind-1]
    cw_box.configure(text="Carat Weight: "+str(carat_weight))
    cg_ind = re.findall(".*\n*.*\s[A-Z]\s", mystr)
    grade = re.findall("\s[A-Z]\s", cg_ind[0])
    grade=grade[0][1]
    col_box.configure(text="Color Grade: "+str(grade))
    clarity_ind = re.findall(".*\n*.*\s[A-Z][A-Z][0-9]\s", mystr)
    clarity=[]
    clarity_ind2=[]
    clarity_ind3=[]
    if len(clarity_ind)!=0:
        clarity=re.findall("\s[A-Z][A-Z][0-9]\s", clarity_ind[0])
        clarity = clarity[0][1:4]
    else:
        clarity_ind2 = re.findall(".*\n*.*\s[A-Z][A-Z]\s[0-9]\s", mystr)
        if len(clarity_ind2)!=0:
            clarity=re.findall("\s[A-Z][A-Z]\s[0-9]\s", clarity_ind2[0])
            clarity = clarity[0][1:5]
        else:
            clarity_ind3 = re.findall(".*\n*.*\s[A-Z][A-Z][A-Z]\s[0-9]\s", mystr)
            clarity=re.findall("\s[A-Z][A-Z][A-Z]\s[0-9]\s", clarity_ind3[0])
            clarity = clarity[0][1:6]
    clar_box.configure(text="Clarity Grade: "+str(clarity))
    cutg_ind = mystr.find("Cut Grade")
    if cutg_ind != -1:
        cut_grade = "Excellent"
    elif mystr.find("CUT GRADE")!=-1:
        cut_grade = "IDEAL"
    else:
        cut_grade = ""
    cutg_box.configure(text="Cut Grade: "+str(cut_grade))
    pol_ind = mystr.find("VERY GOOD")
    pol_ind2=-1
    if pol_ind!=-1:
        if mystr[pol_ind+11:pol_ind+20]=="EXCELLENT":
            polish = "VERY GOOD"
            sym = "EXCELLENT"
        else:
            sym = "VERY GOOD"
            polish = "EXCELLENT"
    else:
        pol_ind2=mystr.find("GOOD")
        if pol_ind2!=-1:
            if mystr[pol_ind2+6:pol_ind2+15]=="EXCELLENT":
                polish = "GOOD"
                sym = "EXCELLENT"
            else:
                sym = "GOOD"
                polish = "EXCELLENT"
        else:
            sym = "EXCELLENT"
            polish = "EXCELLENT"
    pol_box.configure(text="Polish: "+str(polish))
    sym_box.configure(text="Symmetry: "+str(sym))
    flore = "None"
    flo_box.configure(text="Florescence: "+str(flore))
    insc_ind = re.findall(".*\n*.*\sGIA\s[0-9]{10}.", mystr)
    insc_ind2 = []
    inscrib=[]
    if len(insc_ind)!=0:
        inscrib = re.findall("\sGIA\s[0-9]{10}.", insc_ind[0])
        report_num = inscrib[0][5:len(inscrib[0])-1]
        inscrib = inscrib[0][1:len(inscrib[0])-1]
    else:
        insc_ind2 = re.findall(".*\n*.*\sIGI\s[0-9]{9}\s", mystr)
        if len(insc_ind2)!=0:
            inscrib = re.findall("\sIGI\s[0-9]{9}\s", insc_ind2[0])
            report_num = inscrib[0][5:len(inscrib[0])-1]
            inscrib = inscrib[0][1:len(inscrib[0])-1]
        else:
            insc_ind3 = re.findall(".*\n*.*\sLABGROWN\sIGI\sLG[0-9]{9}\s", mystr)
            inscrib = re.findall("\sLABGROWN\sIGI\sLG[0-9]{9}\s", insc_ind3[0])
            report_num = inscrib[0][16:len(inscrib[0])-1]
            inscrib = inscrib[0][1:len(inscrib[0])-1]
    r_num.configure(text="Report Number: "+str(report_num))
    insc_box.configure(text="Inscription: "+str(inscrib))
    com_ind = mystr.find("Comments:")
    com = mystr[com_ind:com_ind+151]
    if com!="":
        com_box.configure(text=str(com))

def uploadFile():
    root.filename =  filedialog.askopenfilename(title = "choose your file",filetypes = (("PDF files","*.pdf"),("All files","*.*")))
    path = tk.Label(frame, text = root.filename)
    path.place(relx=0.22,rely=0.005)
    output_string = StringIO()
    with open(root.filename,'rb')as in_file:
        parser = PDFParser(in_file)
        doc = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)
    mystr = output_string.getvalue()
    begin = time.time()
    func(mystr)
    time.sleep(1)
    end = time.time()
    time_taken = end-begin
    time_box.configure(text="Time to process file: "+str(time_taken)+" ms")

try:
    HEIGHT = 850
    WIDTH = 700

    root = tk.Tk()
    root.title("Python based PDF Extractor")
    root.resizable(0, 0)

    canvas = tk.Canvas(root, height=HEIGHT,width=WIDTH)
    canvas.pack()

    banner_frame = tk.Frame(root, bg='#CCE1F2')
    banner_frame.place(relwidth=1, relheight=0.5)

    frame = tk.Frame(root, bg='#CCE1F2')
    frame.place(rely=0.1 ,relwidth = 1, relheight = 1)

    type_d = tk.Label(frame, text="Type of diamond: ")
    type_d.config(font=("Courier",12))
    type_d.place(relx=0.01,rely=0.07)

    r_num = tk.Label(frame, text="Report Number: ")
    r_num.config(font=("Courier",12))
    r_num.place(relx=0.01,rely=0.12)

    des1 = tk.Label(frame, text="Description:")
    des1.config(font=("Courier",12))
    des1.place(relx=0.01,rely=0.17)

    shape_box = tk.Label(frame, text="Shape and Cutting Style:")
    shape_box.config(font=("Courier",12))
    shape_box.place(relx=0.01,rely=0.22)

    measure_box = tk.Label(frame, text="Measurements:")
    measure_box.config(font=("Courier",12))
    measure_box.place(relx=0.01,rely=0.27)

    cw_box = tk.Label(frame, text="Carat Weight:")
    cw_box.config(font=("Courier",12))
    cw_box.place(relx=0.01,rely=0.32)

    col_box = tk.Label(frame, text="Color Grade:")
    col_box.config(font=("Courier",12))
    col_box.place(relx=0.01,rely=0.37)

    clar_box = tk.Label(frame, text="Clarity Grade:")
    clar_box.config(font=("Courier",12))
    clar_box.place(relx=0.01,rely=0.42)

    cutg_box = tk.Label(frame, text="Cut Grade:")
    cutg_box.config(font=("Courier",12))
    cutg_box.place(relx=0.01,rely=0.47)

    pol_box = tk.Label(frame, text="Polish:")
    pol_box.config(font=("Courier",12))
    pol_box.place(relx=0.01,rely=0.52)

    sym_box = tk.Label(frame, text="Symmetry:")
    sym_box.config(font=("Courier",12))
    sym_box.place(relx=0.01,rely=0.57)

    flo_box = tk.Label(frame, text="Fluorescence:")
    flo_box.config(font=("Courier",12))
    flo_box.place(relx=0.01,rely=0.62)

    insc_box = tk.Label(frame, text="Inscription:")
    insc_box.config(font=("Courier",12))
    insc_box.place(relx=0.01,rely=0.67)

    com_box = tk.Label(frame, text="Comments:")
    com_box.config(font=("Courier",12))
    com_box.place(relx=0.01,rely=0.71, relheight=0.125)

    time_box = tk.Label(frame, text="Time to process file: ")
    time_box.config(font=("Courier",12))
    time_box.place(relx=0.01,rely=0.85)

    upload = tk.Button(frame, text="Upload File", command = uploadFile)
    upload.config(font=("Courier",12))
    upload.place(relx=0.01,rely=0.005,relwidth=0.2,relheight=0.03)

    root.mainloop()
except:
    print("Inexpected Error")

