
import tkinter as t
from openpyxl import load_workbook
import xlrd
import pandas as pd

root=t.Tk()                               #Main window 
f=t.Frame(root)
frame1=t.Frame(root)
frame2=t.Frame(root)
frame3=t.Frame(root)
root.title()
root.geometry("830x395")
root.configure(background="Black")

scrollbar=t.Scrollbar(root)
scrollbar.place(x=800,y=10)

Category=t.StringVar()                    #Declaration of all variables
Prize=t.StringVar()
Brand=t.StringVar()
Size=t.StringVar()
remove_Category=t.StringVar()
remove_Size=t.StringVar()
remove_Brand=t.StringVar()
searchCategory=t.StringVar()
searchBrand=t.StringVar()
searchSize=t.StringVar()
sheet_data=[]
row_data=[]

def emp_dict(*args):                   #To add a new entry and check if entry already exist in excel sheet
    #print("done")
    workbook_name="Users\my-pc\Desktop\sample.xlsx"
    workbook=xlrd.open_workbook(workbook_name)
    worksheet=workbook.sheet_by_index(0)
    
    wb=load_workbook(workbook_name)
    page=wb.active
    
    p=0
    for i in range(worksheet.nrows):
        for j in range(worksheet.ncols):
            cellvalue=worksheet.cell_value(i,j)
            print(cellvalue)   
            sheet_data.append([])
            sheet_data[p]=cellvalue
            p+=1
    print(sheet_data)
    fl=Category.get()
    fsl=fl.lower()
    ll=Price.get()
    lsl=ll.lower()
    if (fsl and lsl) in sheet_data:
        print("found")
        messagebox.showerror("Error","This Product already exist")
    else:
        print("not found")
        for info in args:
            page.append(info)
        messagebox.showinfo("Done","Successfully added the product record")

    wb.save(filename=workbook_name)
    
def add_entries():                       #to append all data and add entries on click the button
    a=" "
    f=Category.get()
    f1=f.lower()
    l=Prize.get()
    l1=l.lower()
    d=Brand.get()
    d1=d.lower()
    de=Size.get()
    de1=de.lower()
    list1=list(a)
    list1.append(f1)
    list1.append(l1)
    list1.append(d1)
    list1.append(de1)
    emp_dict(list1)

def add_info():                                           #for taking user input to add the enteries
    frame2.pack_forget()
    frame3.pack_forget()
    pro_Category=t.Label(frame1,text="Enter category of product:",bg="red",fg="white")
    pro_Category.grid(row=1,column=1,padx=10)
    e1=t.Entry(frame1,textvariable=Category)
    e1.grid(row=1,column=2,padx=10)
    e1.focus()
    pro_prize=t.Label(frame1,text="Prize",bg="red",fg="white")
    pro_prize.grid(row=2,column=1,padx=10)
    e2=t.Entry(frame1,textvariable=Prize)
    e2.grid(row=2,column=2,padx=10)
    pro_Brand=t.Label(frame1,text="Select Brand of Product: ",bg="red",fg="white")
    pro_Brand.grid(row=3,column=1,padx=10)
    Brand.set("Select Option")
    e4=t.OptionMenu(frame1,Brand,"Select Option","Peter England","PRADA","versace","GUCCI","Levis","Lee","Raymond","Wrangle","Park Avenue","Pepe Jeans",'Mufti','Zara')
    e4.grid(row=3,column=2,padx=10)
    Pro_Size=t.Label(frame1,text="Select Size: ",bg="red",fg="white")
    Pro_Size.grid(row=4,column=1,padx=10)
    Size.set("Select Option")
    e5=t.OptionMenu(frame1,Size,"Select Option",'Full',"28",'30','32','34','36','38','M','L','XL','XXL','XXXL','XXXXL','40')
    e5.grid(row=4,column=2,padx=10)
    button4=t.Button(frame1,text="Add entries",command=add_entries)
    button4.grid(row=5,column=2,pady=10)
    
    frame1.configure(background="Red")
    frame1.pack(pady=10)
def clear_all():             #for clearing the entry widgets
    frame1.pack_forget()
    frame2.pack_forget()
    frame3.pack_forget()

    
def remove_emp():                #for taking user input to remove enteries
    clear_all()
    pro_category=t.Label(frame2,text="Enter Category of product",bg="red",fg="white")
    pro_category.grid(row=1,column=1,padx=10)
    e6=t.Entry(frame2,textvariable=remove_Category)
    e6.grid(row=1,column=2,padx=10)
    e6.focus()
    pro_Brand=t.Label(frame2,text="Prize",bg="red",fg="white")
    pro_Brand.grid(row=2,column=1,padx=10)
    e7=t.Entry(frame2,textvariable=remove_Brand)
    e7.grid(row=2,column=2,padx=10)
    pro_Size=t.Label(frame2,text="Size",bg="red",fg="white")
    pro_Size.grid(row=2,column=1,padx=10)
    e8=t.Entry(frame2,textvariable=remove_Size)
    e8.grid(row=2,column=2,padx=10)
    remove_button=t.Button(frame2,text="Click to remove",command=remove_entry)
    remove_button.grid(row=3,column=2,pady=10)
    frame2.configure(background="Red")
    frame2.pack(pady=10)
def remove_entry():  #to remove entry from excel sheet
    rsf=remove_Category.get()
    rsf1=rsf.lower()
    print(rsf1)
    rsl=remove_Brand.get()
    rsl1=rsl.lower()
    print(rsl1)
    workbook_name="sample.xlsx"
    path="Users\my-pc\Desktop\sample.xlsx"
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)

    for row_num in range(sheet.nrows):
        row_value = sheet.row_values(row_num)
        if (row_value[1]==rsf1 and row_value[2]==rsl1):
            print(row_value)
            print("found")
            file="Users\my-pc\Desktop\sample.xlsx"
            x=pd.ExcelFile(file)
            dfs=x.parse(x.sheet_names[0])
            dfs=dfs[dfs['First Name']!=rsf]
            dfs.to_excel("Users\my-pc\Desktop\sample.xlsx",sheet_name='Product',index=False)
            messagebox.showinfo("Done","Successfully removed the Employee record")
    clear_all()
def search_emp():     #can implement search by 1st name,last name,emp id, designation
    clear_all()
    pro_category=t.Label(frame3,text="Enter Category of product",bg="red",fg="white")   #to take user input to seach
    pro_category.grid(row=1,column=1,padx=10)
    e8=t.Entry(frame3,textvariable=searchCategory)
    e8.grid(row=1,column=2,padx=10)
    e8.focus()
    pro_Brand=t.Label(frame3,text="Enter Brand",bg="red",fg="white")
    pro_Brand.grid(row=2,column=1,padx=10)
    e9=t.Entry(frame3,textvariable=searchBrand )
    e9.grid(row=2,column=2,padx=10)
    pro_Size=t.Label(frame3,text="Enter Size",bg="red",fg="white")
    pro_Size.grid(row=2,column=1,padx=10)
    e9=t.Entry(frame3,textvariable=searchSize )
    e9.grid(row=2,column=2,padx=10)
    search_button=t.Button(frame3,text="Click to search",command=search_entry)
    search_button.grid(row=3,column=2,pady=10)
    
    frame3.configure(background="Red")
    frame3.pack(pady=10)

    
def search_entry():
    sf=searchCategory.get()
    ssf1=sf.lower()
    print(ssf1)
    sl=searchSize.get()
    ssl1=sl.lower()
    print(ssl1)
    path='Users\my-pc\Desktop\sample.xlsx'
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)

    for row_num in range(sheet.nrows):
        row_value = sheet.row_values(row_num)
        if (row_value[1]==ssf1 and row_value[2]==ssl1):
            print(row_value)
            print("found")
            messagebox.showinfo("Done","Searched Product Exist")
            clear_all()
    #else:
    if(row_value[1]!=ssf1 and row_value[2]!=ssl1):
        print("Not found")
        messagebox.showerror("Sorry","Product does not Exist")
        clear_all()

        
#Main window buttons and labels
        
label1=t.Label(root,text="Sonu Saree Centre")
label1.config(font=('Italic',16,'bold'),  background="Orange",fg="Yellow", anchor="center")
label1.pack()

label2=t.Label(f,text="Select an action: ",font=('bold',12), background="Black", fg="White")
label2.pack()
button1=t.Button(f,text="Add", background="Brown", fg="White", command=add_info, width=8)
button1.pack()
button2=t.Button(f,text="Remove", background="Brown", fg="white", command=remove_emp, width=8)
button2.pack()
button3=t.Button(f,text="Search", background="Brown", fg="White", command=search_emp, width=8)
button3.pack()
button6=t.Button(f,text="Close", background="Brown", fg="White", width=8, command=root.destroy)
button6.pack()
f.configure(background="Black")
f.pack()

root.mainloop()


    
