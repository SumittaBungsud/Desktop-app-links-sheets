
from asyncio.windows_events import NULL
from msilib.schema import RadioButton

from matplotlib.ft2font import BOLD
from tableName import *
import gspread
from pandas import *
from numpy import *
from oauth2client.service_account import ServiceAccountCredentials
from tkinter import *
from tkinter.ttk import *
from tkcalendar import *
from tkinter.messagebox import askokcancel, showinfo, showerror, showwarning, WARNING, QUESTION

# Get data from Google sheet
scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("API.json", scope)
client = gspread.authorize(creds)

sheet_adv = client.open("Data test").worksheet("Advanced")
sheet_pv = client.open("Data test").worksheet("PV")
data1 = sheet_adv.get_all_records()
data2 = sheet_pv.get_all_records()

# Pandas is border part get data for managing
Adv = DataFrame(data1)  # .set_index("Code")
Pv = DataFrame(data2)  # .set_index("Code")
# sheet_1.update_cell(2, 5, "=now()")


class Window_Result(Tk):
    def __init__(self, form, code):
        # initialize frame
        super().__init__()
        self.title('Searching Results')
        #self.geometry('1025x230')
        self.resizable(1, 1)
        self.form = form
        self.code = code
        if (self.form == 'PV'):
            self.tree = self.create_tree_pv()
        else:
            self.tree = self.create_tree_adv()
        
    def create_tree_pv(self):
        columns = ('code', 'supplier', 'vendor', 'item', 'cost')
        tree = Treeview(self, columns=columns, show='headings', height=20)
        
        tree.column('code', width=100, anchor=CENTER)
        tree.column('supplier', width=200, anchor=W)
        tree.column('vendor', width=200, anchor=W)
        tree.column('item', width=300, anchor=W)
        tree.column('cost', width=100, anchor=W)
        # define headings
        tree.heading('code', text='Code', anchor=CENTER)
        tree.heading('supplier', text='Supplier', anchor=CENTER)
        tree.heading('vendor', text='Vendor', anchor=CENTER)
        tree.heading('item', text='Item', anchor=CENTER)
        tree.heading('cost', text='Cost (THB)', anchor=CENTER)
        
        tree.bind('<<TreeviewSelect>>', self.item_selected_pv)
        tree.grid(row=0, column=0, sticky=NSEW)
        
        # add a scrollbar
        scrollbar = Scrollbar(self, orient=VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky='ns')
        
        # generate sample data
        match = Pv.loc[Pv.Code.str.contains(self.code)][(Pv.Supplier != '') & (Pv.Vendor != '')]
        index = match.index
        if(len(index) > 0):
            order_list = []
            for i in index:
                codet = Pv.Code[i]
                supt = Pv.Supplier[i]
                vent = Pv.Vendor[i]
                itemt = Pv.Items[i]
                costt = Pv['Total (THB)'][i]
                order_list.append((codet, supt, vent, itemt, f'฿ {costt}'))

            # add data to the treeview
            for list in order_list:
                tree.insert('', END, values=list)
                
            return tree     
       
        else: showinfo(title='No data', message='NO DATA EXIST')
            
    
    def create_tree_adv(self):
        columns = ('code', 'supplier', 'vendor', 'remark', 'budget', 'expense')
        tree = Treeview(self, columns=columns, show='headings', height=10)

        tree.column('code', width=100, anchor=CENTER)
        tree.column('supplier', width=250, anchor=W)
        tree.column('vendor', width=200, anchor=W)
        tree.column('remark', width=200, anchor=W)
        tree.column('budget', width=150, anchor=W)
        tree.column('expense', width=150, anchor=W)
        # define headings
        tree.heading('code', text='Code', anchor=CENTER)
        tree.heading('supplier', text='Supplier', anchor=CENTER)
        tree.heading('vendor', text='Vender', anchor=CENTER)
        tree.heading('remark', text='Remark', anchor=CENTER)
        tree.heading('budget', text='Budget (THB)', anchor=CENTER)
        tree.heading('expense', text='Expenditure (THB)', anchor=CENTER)
        
        tree.bind('<<TreeviewSelect>>', self.item_selected_adv)
        tree.grid(row=0, column=0, sticky=NSEW)
        
        # add a scrollbar
        scrollbary = Scrollbar(self, orient=VERTICAL, command=tree.yview)
        scrollbarx = Scrollbar(self, orient=HORIZONTAL, command=tree.xview)
        scrollbary.grid(row=0, column=1, sticky='ns')
        tree.configure(yscroll=scrollbary.set)
        scrollbarx.grid(row=1, column=0, sticky='ns')
        tree.configure(xscroll=scrollbarx.set)
        
        # generate sample data
        match = Adv.loc[Adv.Code.str.contains(self.code)][(Adv.Supplier != '') & (Adv.Vendor != '')]
        index = match.index
        if(len(index) > 0):
            order_list = []
            for i in index:
                codet = Adv.Code[i]
                supt = Adv.Supplier[i]
                vent = Adv.Vendor[i]
                remt = Adv.Remark[i]
                budt = Adv['Budget (THB)'][i]
                expt = Adv['Expenditure (THB)'][i]
                order_list.append((codet, supt, vent, remt, f'฿ {budt}', f'฿ {expt}'))

            # add data to the treeview
            for list in order_list:
                tree.insert('', END, values=list)
                
            return tree
        
        else: showinfo(title='No data', message='NO DATA EXIST')

    
    def item_selected_pv(self, event):
        for selected_item in self.tree.selection():
            item = self.tree.item(selected_item)
            record = item['values'][0]
            
        row = sheet_pv.find(record, in_column=1).row
        rec = sheet_pv.cell(row, 10).value
        # show a message
        msg = f'Status : {rec} '
        showinfo(title='Information', message=msg)
            
        
    def item_selected_adv(self, event):
        for selected_item in self.tree.selection():
            item = self.tree.item(selected_item)
            record = item['values'][0]
            
        for i in range(Adv.shape[0]-1):
            if (Adv["Budget (THB)"][i] != ''):
                key1 = i  
            y=0  
            if ((Adv.Supplier[i] == '') & (Adv.Vendor[i] == '') & (y==0)):
                key2 = i
                y=1
                       
        row = sheet_adv.find(record, in_column=1).row
        rec1 = sheet_adv.cell(row, 8).value
        rec2 = Adv["Budget (THB)"][key1]
        rec3 = Adv.Total[key2]
        # show a message
        msg = f'Status : {rec1} \n Budget Now : {rec3}/{rec2}'
        showinfo(title='Information', message=msg)
        
        
class Application(Notebook):
    def __init__(self, master):
        # initialize frame
        Notebook.__init__(self, master)
        self.grid()
        self.searcher_frame1()
        self.adding_frame2_pv()
        self.adding_frame2_adv()

    # sticky in grid() is compass directional indicator as N,E,S,W or NE(north east)
    def searcher_frame1(self):
        self.frame1 = Frame(self, width=400, height=280)
        self.frame1.grid(columnspan = 7, padx=5, pady=5)
        self.add(self.frame1, text="Search")
        
        self.lb1000 = Label(self.frame1, text="Search by Code : ")
        self.lb1000.grid(column=1, row=2, sticky='WE', padx=5, pady=5)
        
        self.code_find = StringVar()
        self.lb1000_entry = Entry(self.frame1, textvariable=self.code_find, width=25)
        self.lb1000_entry.grid(column=2, row=2, sticky='WE', padx=5, pady=5)
        
        self.bt_sch = Button(self.frame1, text="Search", command=self.seacher_by_code)
        self.bt_sch.grid(column=3, row=2, sticky='WE', padx=5, pady=5)
        
        self.selected_value = StringVar()
        self.std_sch1 = Radiobutton(self.frame1, text='PV', value='PV',
                        variable=self.selected_value)
        self.std_sch1.grid(column=2, row=3, sticky='E', padx=5, pady=5)
        self.std_sch2 = Radiobutton(self.frame1, text='ADV', value='ADV',
                        variable=self.selected_value)
        self.std_sch2.grid(column=3, row=3, sticky='W', padx=5, pady=5)
        
        
    def seacher_by_code(self):
        if((self.code_find.get() != '') & (self.selected_value.get() != '')):
            code_str = self.code_find.get() 
            code_str = code_str.capitalize()
            Window_Result(self.selected_value.get(), code_str)
            
        elif(self.code_find.get() != ''): 
            showinfo(title='Warning', message='Select your \'Form\'!', icon=WARNING)
        elif(self.selected_value.get() != ''): 
            showinfo(title='Warning', message='Fill your \'Code\'!', icon=WARNING)
        else:
            showerror(title='Error', message='Please insert the required input!')


    def adding_frame2_pv(self):
        frame2_pv = Frame(self, width=400, height=280)
        frame2_pv.grid(columnspan = 7, padx=5, pady=5)
        self.add(frame2_pv, text="PV Form ")
        self.grid(column=0, row=1, padx=5, pady=5, sticky='ew')
        
        Lab_pv = LabelFrame(frame2_pv, text="Adding PV")
        Lab_pv.grid(row=2, columnspan=10, sticky='WE',
                  padx=5, pady=5, ipadx=25, ipady=5)
        self.code_pv ='PV'

        # Code
        self.lb1 = Label(Lab_pv, text="Code : ")
        self.lb1.grid(column=0, row=2, sticky='E', padx=5, pady=5)

        self.lb1_code = Label(Lab_pv)
        self.lb1_code.grid(column=1, row=2, sticky='W', padx=5, pady=5)
        
        self.bt_resv = Button(Lab_pv, text="Reserve",
                           command=self.reserve_last_code_pv)
        self.bt_resv.grid(column=2, row=2, sticky='W', padx=5, pady=5)

        # Supplier
        self.lb2 = Label(Lab_pv, text="Supplier : ")
        self.lb2.grid(column=0, row=3, sticky='E', padx=5, pady=5)

        self.supp_1 = StringVar()
        self.lb2_entry = Entry(Lab_pv, textvariable=self.supp_1, width=50)
        self.lb2_entry.grid(column=1, row=3, sticky='E', padx=5, pady=5)

        # Vendor
        self.lb3 = Label(Lab_pv, text="Vendor : ")
        self.lb3.grid(column=0, row=4, sticky='E', padx=5, pady=5)

        self.ven_1 = StringVar()
        self.lb3_entry = Entry(Lab_pv, textvariable=self.ven_1, width=50)
        self.lb3_entry.grid(column=1, row=4, sticky='E', padx=5, pady=5)

        # Remark
        self.lb4 = Label(Lab_pv, text="Remark : ")
        self.lb4.grid(column=0, row=5, sticky='E', padx=5, pady=5)

        self.rem_1 = StringVar()
        self.lb4_entry = Entry(Lab_pv, textvariable=self.rem_1, width=50)
        self.lb4_entry.grid(column=1, row=5, sticky='E', padx=5, pady=5)

        # Items
        self.lb5 = Label(Lab_pv, text="Items : ")
        self.lb5.grid(column=0, row=6, sticky='E', padx=5, pady=5)

        self.item_1 = StringVar()
        self.lb5_entry = Entry(Lab_pv, textvariable=self.item_1, width=50)
        self.lb5_entry.grid(column=1, row=6, sticky='E', padx=5, pady=5)

        # Date PV
        # Calendar
        self.lb6 = Label(Lab_pv, text="Date PV : ")
        self.lb6.grid(column=0, row=7, sticky='E', padx=5, pady=5)

        self.lb6_cal = DateEntry(Lab_pv, width=12, year=2022, month=3, day=28,
                                 background='darkgreen', foreground='white', borderwidth=2)
        self.lb6_cal.grid(column=1, row=7, sticky='W', padx=5, pady=5)

        # Status
        self.lb7 = Label(Lab_pv, text="Status : ")
        self.lb7.grid(column=0, row=8, sticky='E', padx=5, pady=5)

        self.sta_1 = StringVar()
        self.lb7_entry = Entry(Lab_pv, textvariable=self.sta_1, width=50)
        self.lb7_entry.grid(column=1, row=8, sticky='E', padx=5, pady=5)

        # Project Budget
        # combobox be able adding
        self.lb8 = Label(Lab_pv, text="Project Budget : ")
        self.lb8.grid(column=0, row=9, sticky='E', padx=5, pady=5)

        self.lb8_box = Combobox(Lab_pv, width=47)
        self.lb8_box['values'] = ("None", "Optics Lab", "Laser Crosslinks Unit",
                                  "Remote Sensing", "Fibre Optics Battery Sensing", "M.E.D.U.S.A.",
                                  "Lithography", "Satellite Support(ADCS)", "Co-order(See attached file)",
                                  "Satellite Assembly(ODC 1)", "Satellite Assembly(ODC 2)")
        self.lb8_box.current(0)
        self.lb8_box.bind("<<>ComboboxSelected>")
        self.lb8_box.grid(column=1, row=9)

        # Cost
        self.lb9 = Label(Lab_pv, text="Cost : ")
        self.lb9.grid(column=0, row=10, sticky='E', padx=5, pady=5)

        self.double_cost = DoubleVar()
        self.lb9_entry = Entry(Lab_pv, textvariable=self.double_cost, width=50)
        self.lb9_entry.bind('<Return>', self.calculate_total)  # Binding
        self.lb9_entry.grid(column=1, row=10, sticky='E', padx=5, pady=5)

        # Currencf exchang rate
        # filling & combobox of currency
        self.lb10 = Label(Lab_pv, text="Currency Exchange Rate : ")
        self.lb10.grid(column=0, row=11, sticky='E', padx=5, pady=5)

        self.double_cur = DoubleVar()
        self.lb10_entry = Entry(Lab_pv, textvariable=self.double_cur, width=50)
        self.lb10_entry.bind('<Return>', self.calculate_total)  # Binding
        self.lb10_entry.grid(column=1, row=11, sticky='E', padx=5, pady=5)

        self.lb10_box = Combobox(Lab_pv, width=5)
        self.lb10_box['values'] = ("THB", "USD", "SGD")
        self.lb10_box.current(0)
        self.lb10_box.bind("<<>ComboboxSelected>")
        self.lb10_box.grid(column=2, row=11)

        # Total (THB)
        # show result
        self.lb11 = Label(Lab_pv, text="Total (THB) : ")
        self.lb11.grid(column=0, row=12, sticky='E', padx=5, pady=5)

        self.lb11_1 = Label(Lab_pv, text="%.2f" % (0000.0000))
        self.lb11_1.grid(column=1, row=12, sticky='W', padx=5, pady=5)
        
        # Hardcopy  Status
        # combobox
        """self.lb12 = Label(Lab_pv, text="Hardcopy  Status : ")
        self.lb12.grid(column=0, row=13, sticky='E', padx=5, pady=5)

        self.lb12_box = Combobox(Lab_pv, width=47)
        self.lb12_box['values'] = ("None", "Requested quote",
                                   "Digital File", "On Office Box", "Archive")
        self.lb12_box.current(0)
        self.lb12_box.bind("<<>ComboboxSelected>")
        self.lb12_box.grid(column=1, row=13)

        # Attached Files Code
        # waiting...
        self.lb13 = Label(Lab_pv, text="attached files code : ")
        self.lb13.grid(column=0, row=14, sticky='e', padx=5, pady=5)"""
        
        # Submit button
        self.bt1_add = Button(frame2_pv, text="Add", command=self.add_data_pv)
        self.bt1_add.grid(column = 7, row=16, sticky='E', padx=5, pady=5)
        
        # Clear all
        self.bt2_clr = Button(frame2_pv, text="Clear all", command=self.reset_pv)
        self.bt2_clr.grid(column = 8, row=16, sticky='W', padx=5, pady=5)
        
    def add_data_pv(self):
        if(self.supp_1.get() != '') | (self.ven_1.get() != ''):
            answer = askokcancel(
                title='Confirmation',
                message='Do you want to confirm this PV?',
                icon=QUESTION)
            if (answer & ((self.double_cur.get() == 0.0)&(self.lb10_box.get() == 'THB') | (((self.lb10_box.get() == 'USD')|(self.lb10_box.get() == 'SGD'))&(self.double_cur.get() > 0.0)))):
                sheet_pv.update_cell(self.obj3, 2, self.supp_1.get())
                sheet_pv.update_cell(self.obj3, 3, self.ven_1.get())
                sheet_pv.update_cell(self.obj3, 4, self.rem_1.get())
                sheet_pv.update_cell(self.obj3, 5, self.item_1.get())
                               
                if(self.lb10_box.get() == "THB"): 
                    if(self.double_cost.get() == ''):
                        sheet_pv.update_cell(self.obj3, 6, 0)                       
                    else: sheet_pv.update_cell(self.obj3, 6, self.double_cost.get())
                    sheet_pv.update_cell(self.obj3, 7, 0)
                    sheet_pv.update_cell(self.obj3, 8, 0)
                    sheet_pv.update_cell(self.obj3, 12, 0)
                    sheet_pv.update_cell(self.obj3, 13, 0)                   
                elif(self.lb10_box.get() == "USD"):
                    sheet_pv.update_cell(self.obj3, 6, 0)
                    if(self.double_cost.get() == ''):
                        sheet_pv.update_cell(self.obj3, 7, 0)
                        sheet_pv.update_cell(self.obj3, 12, 0)
                    else:
                        sheet_pv.update_cell(self.obj3, 7, self.double_cost.get())
                        sheet_pv.update_cell(self.obj3, 12, self.double_cur.get())
                    sheet_pv.update_cell(self.obj3, 8, 0)
                    sheet_pv.update_cell(self.obj3, 13, 0)
                elif(self.lb10_box.get() == "SGD"):
                    sheet_pv.update_cell(self.obj3, 6, 0)
                    sheet_pv.update_cell(self.obj3, 7, 0)
                    if(self.double_cost.get() == ''):
                        sheet_pv.update_cell(self.obj3, 8, 0)
                        sheet_pv.update_cell(self.obj3, 13, 0)
                    else:
                        sheet_pv.update_cell(self.obj3, 8, self.double_cost.get())
                        sheet_pv.update_cell(self.obj3, 13, self.double_cur.get())
                    sheet_pv.update_cell(self.obj3, 8, self.double_cost.get())
                    sheet_pv.update_cell(self.obj3, 13, self.double_cur.get())
                    sheet_pv.update_cell(self.obj3, 12, 0)
                
                dtt = self.lb6_cal.get()
                dtt_arr = dtt.split('/')
                date = f'{dtt_arr[1]} {date_no[dtt_arr[0]]} 20{dtt_arr[2]}'
                sheet_pv.update_cell(self.obj3, 9, date)
                sheet_pv.update_cell(self.obj3, 15, self.lb8_box.get())
                #sheet_pv.update_cell(self.key_pv+2, 16, self.lb12_box.get())
                
                showinfo(
                    title='Insertion Status',
                    message='The data\'ve been added successfully!')
                
            else: showwarning(
                    title='Warnning!',
                    message='The Exchange Rate is more than 0.0 \n Please change your currency')
        else: 
            showerror(
                title='Informations Error',
                message='Please complete Supplier or Vendor')
        
    def add_data_adv(self): 
        if(self.supp_2.get() != '') | (self.ven_2.get() != ''):
            answer = askokcancel(
                title='Confirmation',
                message='Do you want to confirm this Advanced?',
                icon=QUESTION)
            if answer:
                sheet_adv.update_cell(self.thg3, 2,self.supp_2.get())
                sheet_adv.update_cell(self.thg3, 3,self.ven_2.get())
                sheet_adv.update_cell(self.thg3, 4,self.rem_2.get())
                sheet_adv.update_cell(self.thg3, 6,self.exp_2.get())
                showinfo(
                        title='Insertion Status',
                        message='The data is added successfully!')
        else: 
            showerror(
                title='Informations Error',
                message='Please complete Supplier or Vendor')
        
        
    def reserve_last_code_pv(self):
        try:
            obj1 = sheet_pv.find('', in_column=2).row
            obj2 = sheet_pv.find('', in_column=3).row
            self.obj3 = max(obj1, obj2)
            self.id_pv = sheet_pv.cell(self.obj3, 1).value
            self.lb1_code.config(text=self.id_pv)
            sheet_pv.update_cell(self.obj3, 2, "[ Reserved ]")   
                
        except ValueError as error:
            showerror(title='Error', message=error)
            
            
    def reserve_last_code_adv(self):
        try:
            for i in range(Adv.shape[0]-1):
                if (Adv["Budget (THB)"][i] != ''):
                    key2 = i
                    
            thg1 = sheet_adv.find('', in_column=2).row
            thg2 = sheet_adv.find('', in_column=3).row
            self.thg3 = max(thg1, thg2)
            self.id_adv = sheet_adv.cell(self.thg3, 1).value     
            self.total_budget = Adv["Budget (THB)"][key2]
            self.remain_avail = float(sheet_adv.cell(self.thg3, 7).value)
            self.lb100_code.config(text=self.id_adv)
            self.lb105_t.config(text=" %d THB "%(self.total_budget))
            self.lb106_r.config(text=" %.2f THB "%(self.remain_avail))
            sheet_adv.update_cell(self.thg3, 2, "[ Reserved ]")
            
        except ValueError as error:
            showerror(title='Error', message=error)


    def calculate_total(self, event):  # Function of enter action
        cost = self.double_cost.get()
        currency_rate = self.double_cur.get()
        if(currency_rate <= 0): total = cost
        else: total = cost*currency_rate
        self.lb11_1.config(text="%.2f THB" % (total))
        
        
    def calculate_remain(self, event):
        cost = self.exp_2.get()
        remain = self.remain_avail
        if(remain >= cost):
            deduct = remain-cost
            self.lb108_d.config(text=" %.2f THB"%(deduct))
        else: 
            self.lb107_entry.delete(0, "end")
            self.lb108_d.config(text = '')
            showerror(
            title='Error',
            message='The Budget is not sufficiently!')
            
            
    def reset_pv(self):
        self.lb2_entry.delete(0, "end")
        self.lb3_entry.delete(0, "end")
        self.lb4_entry.delete(0, "end")
        self.lb5_entry.delete(0, "end")
        self.lb6_cal.delete(0, "end")
        self.lb8_box.current(0)
        self.lb9_entry.delete(0, "end")
        self.lb10_entry.delete(0, "end")
        self.lb1_code.config(text = '')
        self.lb11_1.config(text = '')
        if(sheet_pv.cell(self.obj3, 2).value == '[ Reserved ]'):
            sheet_pv.update_cell(self.obj3, 2, '')
        
        
    def reset_adv(self):
        self.lb101_entry.delete(0, "end")
        self.lb102_entry.delete(0, "end")
        self.lb103_entry.delete(0, "end")
        self.lb104_entry.delete(0, "end")
        self.lb107_entry.delete(0, "end")
        self.lb100_code.config(text = '')
        self.lb105_t.config(text = '')
        self.lb106_r.config(text = '')
        self.lb108_d.config(text = '')
        if(sheet_adv.cell(self.thg3, 2).value == '[ Reserved ]'):
            sheet_adv.update_cell(self.thg3, 2, '')
            
    def adding_frame2_adv(self):
        frame2_adv = Frame(self, width=400, height=280)
        frame2_adv.grid(columnspan = 7, padx=5, pady=5)
        self.add(frame2_adv, text="Advanced Form ")
        self.grid(column=0, row=1, padx=5, pady=5, sticky='ew')
        
        Lab_adv = LabelFrame(frame2_adv, text="Adding Advanced")
        Lab_adv.grid(row=2, columnspan=10, sticky='WE',
                  padx=5, pady=5, ipadx=25, ipady=5)
        self.code_adv = 'Advanced'

        # Code
        self.lb100 = Label(Lab_adv, text="Code : ")
        self.lb100.grid(column=0, row=2, sticky='E', padx=5, pady=5)

        self.lb100_code = Label(Lab_adv)
        self.lb100_code.grid(column=1, row=2, sticky='W', padx=5, pady=5)
        
        self.bt100 = Button(Lab_adv, text="Reserve",
                           command=self.reserve_last_code_adv)
        self.bt100.grid(column=2, row=2, sticky='E', padx=5, pady=5)

        # Supplier
        self.lb101 = Label(Lab_adv, text="Supplier : ")
        self.lb101.grid(column=0, row=3, sticky='E', padx=5, pady=5)

        self.supp_2 = StringVar()
        self.lb101_entry = Entry(Lab_adv, width=50)
        self.lb101_entry.grid(column=1, row=3, sticky='E', padx=5, pady=5)

        # Vendor
        self.lb102 = Label(Lab_adv, text="Vendor : ")
        self.lb102.grid(column=0, row=4, sticky='E', padx=5, pady=5)

        self.ven_2 = StringVar()
        self.lb102_entry = Entry(Lab_adv, width=50)
        self.lb102_entry.grid(column=1, row=4, sticky='E', padx=5, pady=5)

        # Remark
        self.lb103 = Label(Lab_adv, text="Remark : ")
        self.lb103.grid(column=0, row=5, sticky='E', padx=5, pady=5)

        self.rem_2 = StringVar()
        self.lb103_entry = Entry(Lab_adv, width=50)
        self.lb103_entry.grid(column=1, row=5, sticky='E', padx=5, pady=5)

        # Items
        self.lb104 = Label(Lab_adv, text="Items : ")
        self.lb104.grid(column=0, row=6, sticky='E', padx=5, pady=5)

        self.item_2 = StringVar()
        self.lb104_entry = Entry(Lab_adv, width=50)
        self.lb104_entry.grid(column=1, row=6, sticky='E', padx=5, pady=5)
        
        # Budget total
        self.lb105 = Label(Lab_adv, text="Total Budget : ")
        self.lb105.grid(column=0, row=7, sticky='E', padx=5, pady=5)
        
        self.lb105_t = Label(Lab_adv)
        self.lb105_t.grid(column=1, row=7, sticky='W', padx=5, pady=5)
        
        # Budget Available
        self.lb106 = Label(Lab_adv, text="Budget Available : ")
        self.lb106.grid(column=0, row=8, sticky='E', padx=5, pady=5)
        
        self.lb106_r = Label(Lab_adv)
        self.lb106_r.grid(column=1, row=8, sticky='W', padx=5, pady=5)

        # Expense
        self.lb107 = Label(Lab_adv, text="Expenditure (THB) : ")
        self.lb107.grid(column=0, row=9, sticky='E', padx=5, pady=5)
        
        self.exp_2 = DoubleVar()
        self.lb107_entry = Entry(Lab_adv,textvariable=self.exp_2, width=50)
        self.lb107_entry.bind('<Return>', self.calculate_remain)  # Binding
        self.lb107_entry.grid(column=1, row=9, sticky='E', padx=5, pady=5)
        
        # Remaining after deduction
        self.lb108 = Label(Lab_adv, text="Remain : ")
        self.lb108.grid(column=0, row=10, sticky='E', padx=5, pady=5)
        
        self.lb108_d = Label(Lab_adv)
        self.lb108_d.grid(column=1, row=10, sticky='W', padx=5, pady=5)
        
        # Hardcopy  Status
        # combobox
        """self.lb108 = Label(Lab_adv, text="Hardcopy  Status : ")
        self.lb108.grid(column=0, row=9, sticky='E', padx=5, pady=5)

        self.lb108_box = Combobox(Lab_adv, width=47)
        self.lb108_box['values'] = ("None", "Requested quote",
                                   "Digital File", "On Office Box", "Archive")
        self.lb108_box.current(0)
        self.lb108_box.bind("<<>ComboboxSelected>")
        self.lb108_box.grid(column=1, row=9)"""

        # Submit button
        self.bt2_add = Button(frame2_adv, text="Add")
        self.bt2_add.grid(column = 8, row=16, sticky='E', padx=5, pady=5)
        
        # Clear all
        self.bt2_clr = Button(frame2_adv, text="Clear all", command=self.reset_adv)
        self.bt2_clr.grid(column = 9, row=16, sticky='W', padx=5, pady=5)
    
        


if __name__ == "__main__":
    #window.bind('<Return>', Application.calculate_total)
    window = Tk()
    window.title("Data Insertion")
    #window.geometry('570x440')
    window.resizable(1,1)
    
    app = Application(window)
    
    window.mainloop()
