import os
import tkinter
from tkinter import *
from tkinter.messagebox import showinfo
from tkinter import messagebox
from tkinter.font import Font
from tkinter import filedialog
from tkinter.filedialog import askopenfile
import win32com.client
from PIL import Image, ImageTk
import img2pdf

def raise_frame(frame):
    frame.tkraise()


root = Tk()
root.title('File Converter')
root.iconbitmap("logos.ico")
root.configure(background="#f5f0e1")
root.geometry("740x400")

pdf2pptxpage = Frame(root)
pdf2pptxpage.place(x=0, y=0, width=740, height=400)

pdf2wordpage = Frame(root)
pdf2wordpage.place(x=0, y=0, width=740, height=400)

img2pdfpage = Frame(root)
img2pdfpage.place(x=0, y=0, width=740, height=400)

ppt2pdfpage = Frame(root)
ppt2pdfpage.place(x=0, y=0, width=740, height=400)

wo2pdfpage = Frame(root)
wo2pdfpage.place(x=0, y=0, width=740, height=400)

converterpage = Frame(root)
converterpage.place(x=0, y=0, width=740, height=400)

homepage = Frame(root)
homepage.place(x=0, y=0, width=740, height=400)

# FONT LUCIDA CONSOLE
lucida = Font(
    family = "Lucida Console",
    size = 16,
    weight = "normal",
    )

# HOMEPAGE
# LOGO
logo = PhotoImage(file="logos.png")
labellogo = Label(homepage, image = logo)
labellogo.pack()

#File Converter
fconvert =Button(homepage, text="FILE CONVERTER", width="15", font=lucida,\
                 bg="#ff6e40", fg="white",  command=lambda: raise_frame(converterpage))
fconvert.place(x=270 , y = 200)

#CREDIT
credit = Label(homepage, text="@mdavaf17")
credit.place(x=10, y=375)

# FILE CONVERTER PAGE
# LOGO
logo2 = PhotoImage(file="logos.png")
labellogo2 = Label(converterpage, image = logo2)
labellogo2.grid(row=0, column=2)

# WORD TO PDF
wo2pd =Button(converterpage, text="WORD TO PDF", width="15", font=lucida,\
                 bg="#ff6e40", fg="white", command=lambda: raise_frame(wo2pdfpage))
wo2pd.grid(row=1, column=1, padx = 20 , pady = 50)

# PPT TO PDF
pp2pd =Button(converterpage, text="PPT TO PDF", width="15", font=lucida, \
              bg="#ff6e40", fg="white", command=lambda: raise_frame(ppt2pdfpage))
pp2pd.grid(row=1, column=2, padx = 20)

# IMAGE TO PDF
img2pd =Button(converterpage, text="IMAGE TO PDF", width="15", font=lucida, \
               bg="#ff6e40", fg="white", command=lambda: raise_frame(img2pdfpage))
img2pd.grid(row=1, column=3, padx = 20)

# PDF TO WORD
pd2wo =Button(converterpage, text="PDF TO WORD", width="15", font=lucida, \
              bg="#ff6e40", fg="white")
pd2wo.grid(row=2, column=1, columnspan=2, padx = 20)

# PDF TO PPT
pd2pp =Button(converterpage, text="PDF TO PPT", width="15", font=lucida, \
              bg="#ff6e40", fg="white")
pd2pp.grid(row=2, column=2, columnspan=2, padx = 20)

# BACK HOME
btn_back = Button(converterpage, text="Back", width="5", font="Arial,(12)", \
           command=lambda: raise_frame(homepage), bg="#1e3d59", fg="white")
btn_back.grid(row=3, column=2, pady=30)

# WO2PDFPAGE
def openWord():
    fileword = askopenfile(filetypes = [("Word Files", "*.docx")] )
    getfwname = fileword.name
    newwordname = getfwname.replace("/", "\\")
    out_filew = os.path.splitext(newwordname)[0]
    worddoc = win32com.client.Dispatch('Word.Application')
    doc = worddoc.Documents.Open(newwordname)
    doc.SaveAs(out_filew, FileFormat = 17)
    doc.Close()
    worddoc.Quit()
    showinfo("Complete !", "Word to PDF Succesfully Converted")


# LOGO
logo3 = PhotoImage(file="logos.png")
labellogo3 = Label(wo2pdfpage, image = logo3)
labellogo3.grid(row=0, column=2)

# LABEL CHOOSE FILE
choose_file = Label(wo2pdfpage, text="File", width="15", font=lucida, bg="#1e3d59", \
                    fg="white")
choose_file.grid(row=1, column=1, padx = (50,0) , pady = 50)

# TITIK CHOOSE FILE
titik_choose_file = Label(wo2pdfpage, text=":", width="15", font=lucida, bg="#1e3d59", fg="white")
titik_choose_file.grid(row=1, column=2)

# BROWSE FILE
btn_browse = Button(wo2pdfpage, text="Browse File...", width="15", font=lucida, \
                    bg="#ffc13b", fg="black", relief=RAISED, command=openWord)
btn_browse.grid(row=1, column=3)

# BACK FILE CONVERTER
btn_back2 = Button(wo2pdfpage, text="Back", width="5", font="Arial,(12)", \
           command=lambda: raise_frame(converterpage), bg="#1e3d59", fg="white")
btn_back2.grid(row=3, column=2, pady=30)

# PPT2PDF
def openPpt():
    fileppt = askopenfile(filetypes=[("Presentation Files", "*.pptx")])
    getfname = fileppt.name
    newpptname = getfname.replace("/", "\\")
    out_file = os.path.splitext(newpptname)[0]
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    pdf = powerpoint.Presentations.Open(newpptname, WithWindow=False)
    pdf.SaveAs(out_file, 32)
    pdf.Close()
    powerpoint.Quit()
    showinfo("Complete !", "PPT to PDF Succesfully Converted")


# LOGO
logo4 = PhotoImage(file="logos.png")
labellogo4 = Label(ppt2pdfpage, image = logo3)
labellogo4.grid(row=0, column=2)

# LABEL CHOOSE FILE
choose_file = Label(ppt2pdfpage, text="File", width="15", font=lucida, bg="#1e3d59", \
                    fg="white")
choose_file.grid(row=1, column=1, padx = (50,0) , pady = 50)

# TITIK CHOOSE FILE
titik_choose_file = Label(ppt2pdfpage, text=":", width="15", font=lucida, bg="#1e3d59", fg="white")
titik_choose_file.grid(row=1, column=2)

# BROWSE FILE
btn_browse = Button(ppt2pdfpage, text="Browse File...", width="15", font=lucida, \
                    bg="#ffc13b", fg="black", relief=RAISED, command=openPpt)
btn_browse.grid(row=1, column=3)

# BACK FILE CONVERTER
btn_back2 = Button(ppt2pdfpage, text="Back", width="5", font="Arial,(12)", \
           command=lambda: raise_frame(converterpage), bg="#1e3d59", fg="white")
btn_back2.grid(row=3, column=2, pady=30)

# IMAGE TO PDF
def openImg():
    fileimg = askopenfile(filetypes=[("Image Files", ".jpg .jpeg .png")])
    getImgname = fileimg.name
    pdf = Image.open(getImgname)
    if pdf.mode == "RGBA" :
        pdf = pdf.convert("RGB")
    if ".jpg" in getImgname:
        newImgPdfname = getImgname.replace(".jpg", ".pdf")
    if ".jpeg" in getImgname:
        newImgPdfname = getImgname.replace(".jpeg", ".pdf")
    if ".png" in getImgname:
        newImgPdfname = getImgname.replace(".png", ".pdf")
    pdf.save(newImgPdfname)
    showinfo("Complete !", "Image to PDF Succesfully Converted")


# LOGO
logo5 = PhotoImage(file="logos.png")
labellogo5 = Label(img2pdfpage, image = logo3)
labellogo5.grid(row=0, column=2)

# LABEL CHOOSE FILE
choose_file = Label(img2pdfpage, text="File", width="15", font=lucida, bg="#1e3d59", \
                    fg="white")
choose_file.grid(row=1, column=1, padx = (50,0) , pady = 50)

# TITIK CHOOSE FILE
titik_choose_file = Label(img2pdfpage, text=":", width="15", font=lucida, bg="#1e3d59", fg="white")
titik_choose_file.grid(row=1, column=2)

# BROWSE FILE
btn_browse = Button(img2pdfpage, text="Browse File...", width="15", font=lucida, \
                    bg="#ffc13b", fg="black", relief=RAISED, command=openImg)
btn_browse.grid(row=1, column=3)

# BACK FILE CONVERTER
btn_back3 = Button(img2pdfpage, text="Back", width="5", font="Arial,(12)", \
           command=lambda: raise_frame(converterpage), bg="#1e3d59", fg="white")
btn_back3.grid(row=3, column=2, pady=30)


frame=Frame(root,relief=RAISED,borderwidth=1)
frame.pack(fill=X)

root.mainloop()