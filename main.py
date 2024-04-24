#importing the modules
import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime

#application window
window = tkinter.Tk()
window.title("Invoice Generator Form")

#frame
frame=tkinter.Frame(window)
frame.pack(padx=20,pady=10)

#clearing the items after each generation
def clear_item():
    qty_spinbox.delete(0,tkinter.END)
    qty_spinbox.insert(0,'1')
    desc_entry.delete(0,tkinter.END)
    up_spinbox.delete(0,tkinter.END)
    up_spinbox.insert(0,1)

#inserting the values in the tree
invoice_list=[]
def add_item():
    qty=int(qty_spinbox.get())
    desc=desc_entry.get()
    price=float(up_spinbox.get())
    line_total=qty*price
    invoice_item=[qty,desc,price,line_total]
    tree.insert('',0,values=invoice_item)
    clear_item()
    invoice_list.append(invoice_item)

#clearing the tree
def clear_invoice():
    first_name_entry.delete(0,tkinter.END)
    last_name_entry.delete(0,tkinter.END)
    phone_entry.delete(0,tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())

#gnerate invoice
def genrate_invoice():
    doc = DocxTemplate("invoice_template.docx")
    name = first_name_entry.get()+last_name_entry.get()
    phone = phone_entry.get()
    subtotal = sum(item[3] for item in invoice_list) 
    salestax = 0.1
    total = subtotal*(1-salestax)
    
    doc.render({"name":name, 
            "phone":phone,
            "invoice_list": invoice_list,
            "subtotal":subtotal,
            "salestax":str(salestax*100)+"%",
            "total":total})
    
    doc_name = "new_invoice" + name + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
    doc.save(doc_name)
    
    clear_invoice()
    
    
    
#creating widgets
first_name=tkinter.Label(frame,text='First Name')
first_name.grid(row=0,column=0)
first_name_entry=tkinter.Entry(frame)
first_name_entry.grid(row=1,column=0)

last_name=tkinter.Label(frame,text='Last Name')
last_name.grid(row=0,column=1)
last_name_entry=tkinter.Entry(frame)
last_name_entry.grid(row=1,column=1)

phone=tkinter.Label(frame,text='Phone')
phone.grid(row=0,column=2)
phone_entry=tkinter.Entry(frame)
phone_entry.grid(row=1,column=2)

qty_label=tkinter.Label(frame,text='Qty')
qty_label.grid(row=2,column=0)
qty_spinbox=tkinter.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row=3,column=0)

desc_label=tkinter.Label(frame,text='Describtion')
desc_label.grid(row=2,column=1)
desc_entry=tkinter.Entry(frame)
desc_entry.grid(row=3,column=1)

up_label=tkinter.Label(frame,text='Unit Price')
up_label.grid(row=2,column=2)
up_spinbox=tkinter.Spinbox(frame, from_=1, to=100, increment=0.5)
up_spinbox.grid(row=3,column=2)

add_item_button=tkinter.Button(frame,text='Add Item',command=add_item)
add_item_button.grid(row=4,column=2,padx=20,pady=10)

cols=('item','qty','price','total')
tree=ttk.Treeview(frame,columns=cols,show='headings')

tree.grid(row=5,column=0,columnspan=3,padx=20,pady=10)
tree.heading('item',text='Item')
tree.heading('qty',text='QTY')
tree.heading('price',text='Price')
tree.heading('total',text='Total')

save_invoice=tkinter.Button(frame,text='Generate Invoice',command=genrate_invoice)
save_invoice.grid(row=6,column=0,columnspan=3,sticky='nsew',padx=5,pady=5)

new_invoice=tkinter.Button(frame,text='New Invoice',command=clear_invoice)
new_invoice.grid(row=7,column=0,columnspan=3,sticky='nsew',padx=5,pady=5)






window.mainloop()