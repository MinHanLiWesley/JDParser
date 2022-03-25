# -*- coding: utf-8 -*-
from tkinter import filedialog
from tkinter import *
from tkinter import ttk
from ttkthemes import ThemedStyle
from PIL import Image,ImageTk
from pdf2image import convert_from_path
import pandas as pd


HEADER=[
    'Experience‌ ‌Required‌','Preferred‌ ‌Experience‌','Keys‌ ‌to‌ ‌Success‌ ‌in‌ ‌this‌ ‌Role','Qualifications','What we need to see','Ways to stand out from the crowd','Plus','Keys to Success in this Role'
]

def op():
 sfname = filedialog.askopenfilename(title='選擇', filetypes=[('pdf', '*.pdf'), ('All Files', '*')])
 return sfname
def read(sfname):
 df = pd.read_excel(sfname,header=0)
 cols = list(df.columns)
 return df,cols
def show(frame,df,cols):
 tree = ttk.Treeview(frame)
 tree.pack()
 tree["columns"] = cols
 for i in cols:
    tree.column(i,anchor="center")
    tree.heading(i,text=i,anchor='center')
 for index, row in df.iterrows():
    tree.insert("",'end',text = index,values=list(row))
 return tree
def oas():
 global root
 filename = op()
 data,c = read(filename)
 tree = show(root,data,c)
 tree.place(relx=0,rely=0.1,relheight=0.7,relwidth=1)
 
def save():
    fname = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),
                                        ("All files", "*.*")))
    
    # store workbook
    
    
    # note: this will fail unless user ends the fname with ".xlsx"
    df.to_excel(fname)



def insert_header(event=None):
    global header
    global Lb1
    text = header.get()
    Lb1.insert("end",text)
    HEADER.append(text)
    header.delete(0,END)


def delete_header(event=None):
    global Lb1
    index = Lb1.curselection()
    Lb1.delete(index)

def select_required_header():
    global Lb1
    global Lb2
    Lb2.delete(0,END)
    index = Lb1.curselection()
    required_header  = Lb1.get(index)
    Lb2.insert("end", required_header)
    
def show_pdf():
    global pdf
    filename = op()
    # Here the PDF is converted to list of images
    pages = convert_from_path(filename,size=(800,900))
    # Empty list for storing images
    photos = []
    # Storing the converted images into list
    for i in range(len(pages)):
      photos.append(ImageTk.PhotoImage(pages[i]))
    # Adding all the images to the text widget
    print(photos)
    for photo in photos:
      pdf.image_create(END,image=photo)
      
      # For Seperating the pages
      # pdf.insert(END,'\n\n')
    # Ending of mainloop
    mainloop()
    
    

def Add_menu(my_menu):
    file_menu = Menu(my_menu, tearoff=False)
    my_menu.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="Open", command=show_pdf)
    # file_menu.add_command(label="Clear", command=clear_text_box)
    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=root.quit)

def main():
    global root
    global pdf
    global header
    global Lb1
    global Lb2
    root = Tk()
    style = ThemedStyle(root)
    style.set_theme("breeze")
    root.geometry("1100x500")
    root.resizable(0,0)
    # Menu
    my_menu = Menu(root)
    root.config(menu=my_menu)
    Add_menu(my_menu)
    # B1 = Button(root, text="打開",command = oas).pack()
    pdf_frame = Frame(root,width=500,height=460).pack(side=RIGHT,fill=BOTH,expand=0)
    # Adding Scrollbar to the PDF frame
    scrol_y = Scrollbar(pdf_frame,orient=VERTICAL)
    scrol_x = Scrollbar(pdf_frame,orient=HORIZONTAL)
    # Adding text widget for inserting images
    pdf = Text(pdf_frame,yscrollcommand=scrol_y.set,bg="grey")
    # Setting the scrollbar to the right side
    scrol_y.pack(side=RIGHT,fill=Y)
    scrol_y.config(command=pdf.yview)
    # Finally packing the text widget
    pdf.pack(fill=BOTH,expand=1)
    # Button(root,text="Insert Header",command = insert_header).pack()
    root.bind("<Return>",insert_header)
    Button(root,text="Delete Header",command = delete_header).place(x=800,y=10,height=20)
    Button(root,text="Select",command = select_required_header).place(x=900,y=350,height=20,width=100)
    root.bind("<Delete>",delete_header)
    Lb1 = Listbox(root,bg='white')
    Lb1.place(x=600,y=40,width = 400,height=300)
    scrol_y_lb = Scrollbar(Lb1)
    scrol_y_lb.pack(side=RIGHT,fill=BOTH)
    Lb1.config(yscrollcommand=scrol_y_lb.set)
    scrol_y_lb.config(command=Lb1.yview)
    # header entry
    L1 = Label(root,text="Header")
    L1.place(x=600,y=10)
    L2 = Label(root,text="Required Header: The first section that you may want "\
               "to append in table.\n Every statements after this header will be included.")
    L2.place(x=600,y=370)
    Lb2 = Listbox(root,bg='white')
    Lb2.place(x=600,y=350,width= 300,height=20)
    header = Entry(root,text="Hi",bg='white')
    header.place(x=650,y=10)
    
    
    Button(root,text="Export as",command=save).place(x=750,y=425,height=50,width=100)
    
    
    mainloop()
if __name__=='__main__':
    main()