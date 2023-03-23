import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from datetime import date
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import calendar
import sys,os

def resource_path(relative_path: str) -> str:
    """ Get absolute path to resource"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

pendapatan_worksheet = 0
pembelian_worksheet = 0
last_row = 0
last_row_pembelian = 0
nomor_jumlah = 0
today = date.today()
now = datetime.now()

def start_up_sequence_sheet():                          # Sequence of things done by the system before starting the application
    month = today.strftime("%b")
    year = str(today.year%100)
    name = month+'-'+year
    if today.day <= 3:      #make new worksheet if the date is less than 3, or the start of the month
        try:
            make_worksheet_sheet(name)
        except:
            pass
            
        try:
            make_worksheet_pembelian_sheet(name)
        except:
            pass

    open_worksheet_sheet(name)

def update_last_row_sheet(final_row):
    pendapatan_worksheet.update_acell('G1', final_row)

def update_last_row_pembelian_sheet(final_row):
    pembelian_worksheet.update_acell('P1', final_row)

def open_worksheet_sheet(input_name):
    global pendapatan_worksheet
    global pembelian_worksheet
    global last_row
    global last_row_pembelian
    pendapatan_worksheet = textsheet.worksheet(input_name)
    pembelian_worksheet = textsheet_pembelian.worksheet(input_name)
    last_row = pendapatan_worksheet.acell('G1').value
    last_row_pembelian = pembelian_worksheet.acell('P1').value

    if last_row == None:
        last_row = 1
    if last_row_pembelian == None:
        last_row_pembelian = 1
    if now.hour <= 11:
        pendapatan_worksheet.update_acell('A'+str(last_row), str(today))
        last_row = int(last_row) + 1
    update_last_row_sheet(int(last_row))
    update_last_row_pembelian_sheet(int(last_row_pembelian))

def print_data_sheet(time, num_o_products, data, total):
    global pendapatan_worksheet
    global last_row
    for i in range(num_o_products):
        row_i = i
        pendapatan_worksheet.update_acell('B'+str(int(last_row)+row_i), data[i][0])
        pendapatan_worksheet.update_acell('C'+str(int(last_row)+row_i), data[i][1])
        pendapatan_worksheet.update_acell('D'+str(int(last_row)+row_i), data[i][2])
        pendapatan_worksheet.update_acell('E'+str(int(last_row)+row_i), data[i][3])
    pendapatan_worksheet.update_acell('A'+str(last_row), time)
    last_row = int(last_row)+num_o_products
    update_last_row_sheet(last_row)
    pendapatan_worksheet.update_acell('F'+str(int(last_row)-1), total)

def make_worksheet_sheet(input_name):                                       # To make sheet for database
    textsheet.add_worksheet(title=input_name, rows=8000, cols=20)

def make_worksheet_pembelian_sheet(input_name):                             # To make sheet to store sales data
    textsheet_pembelian.add_worksheet(title=input_name, rows=8000, cols=20)

def reduce_qty_sheet(row, reduce_num):
    val = int(sheet.acell('F'+str(row)).value)
    val = val - reduce_num
    sheet.update_acell('F'+str(row), val)

def seek_in_database_sheet(key_word):
    matching_cells = []
    a=len(sheet.col_values(1))
    for row_num in range(1, a+1):
        val = sheet.acell('B'+str(row_num)).value
        if key_word.lower() in val.lower():
            matching_cells.append(find_data_sheet(row_num))
    return matching_cells

def cell_format_sheet():
    last_row = find_last_row_sheet() + 1
    sheet.format('G'+str(last_row)+':'+'H'+str(last_row), {'numberFormat':{'type': 'CURRENCY', 'pattern': 'Rp#,###'}})      #Change the currency and pattern

def delete_row_sheet(row_num):
    sheet.delete_row(row_num)
    after_last_row = find_last_row_sheet() + 1
    under_row = after_last_row-row_num
    for i in range(under_row):
        val = sheet.acell('A'+str(row_num+i)).value
        sheet.update_acell('A'+str(row_num+i), int(val)-1)
    empty_row = ['' for cell in range(sheet.col_count)]
    sheet.insert_row(empty_row, index=after_last_row)

def find_last_row_sheet():
    a=len(sheet.col_values(1))
    return a

def find_data_sheet(row):
    values_list = sheet.row_values(row)
    return values_list
    
def input_data_sheet():
    #Check the number of filled columns
    list_of_list = sheet.get_all_values()
    list_of_list.pop(0)
    return list_of_list

def add_data_sheet(current_num, num, name, price, type, finprice, total_stock, quant):
    sheet.update_acell('A'+str(current_num+1), current_num)
    sheet.update_acell('B'+str(current_num+1), name)
    sheet.update_acell('C'+str(current_num+1), type)
    sheet.update_acell('D'+str(current_num+1), num)
    sheet.update_acell('E'+str(current_num+1), quant)
    sheet.update_acell('F'+str(current_num+1), total_stock)
    sheet.update_acell('G'+str(current_num+1), int(finprice))
    sheet.update_acell('H'+str(current_num+1), int(price))

def replace_data_sheet(pos,num, price, finprice, stok_am):      #replace a specific data in database
    sheet.update_cell(pos, 4, num)
    sheet.update_cell(pos, 6, stok_am)
    sheet.update_cell(pos, 7, finprice)
    sheet.update_cell(pos, 8, price)

harga_item_dict = {}
history = []              

def on_close():                                                 # Final warning before exiting
    response= messagebox.askyesno('Exit','Do you want to exit ERM ?\nHistory will be automatically deleted !')
    if response:
        window.destroy()

def totalin():
    global harga_item_dict
    jumlah = 0
    if len(harga_item_dict) > 0:
        for item in harga_item_dict:
            jumlah = jumlah + harga_item_dict[item][0]
    total.set(str(jumlah))


def check_histori():                                            # For checking the history for today's sales
    jumlah = 0

    if len(history) > 0:
        for item in histori_tree.get_children():
            histori_tree.delete(item)

        for item in history:
            histori_tree.insert("", 'end', iid=item[0], text=item[0],
                    values=(item[0], item[1], item[2]))
            jumlah = jumlah + item[2]
        income_total.set('Rp.'+str(jumlah))

def finishing(event):                                           # After finishing a purchase
    global total
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")

    def put_in_spread():
        put_data = []
        for item in tree_cashier.get_children():
            one_item = tree_cashier.item(item)['values']
            put_data.append([one_item[0], one_item[1], one_item[2], one_item[5]])
        print_data_sheet(str(current_time), berapa_byk, put_data, int(total.get()))

    def save_history(total, num):
        global history
        
        history.append([current_time ,num, total])

    def itung_stok_kosongin_dict():
        global harga_item_dict

        for item in harga_item_dict:
            reduce_qty_sheet(int(item)+1, harga_item_dict[item][1])

        harga_item_dict = {}
        totalin()

    def itung_return(*args):
        uang = get_uang_price.get()
        if uang != '':
            kembalian = int(uang) - int(total.get())
            get_return_price.set(str(kembalian))
    
    def empty(event):
        if int(get_uang_price.get()) >= int(total.get()):
            put_in_spread()
            save_history(int(total.get()), berapa_byk)
            itung_stok_kosongin_dict()
            for item in tree_cashier.get_children():
                tree_cashier.delete(item)
            update_database()
            newWindow4.destroy()
    
    berapa_byk = len(tree_cashier.get_children())

    if berapa_byk > 0:                          # Final windows before checking out
        get_return_price = StringVar()
        get_return_price.set('0')
        get_uang_price = StringVar()
        get_uang_price.set('0')
        newWindow4 = Toplevel(window)
        newWindow4.title("Electronic Register and Management")
        ttk.Label(newWindow4, 
                text ="Total Money:", font=("Arial", 17), anchor='e').grid(column = 1, 
                               row = 1,
                               padx = 10,
                               pady = 15)
        get_money = Entry(newWindow4, width=10, font="Arial 35", textvariable=get_uang_price)
        get_money.grid(column = 2, row = 1, padx = 15,pady=10)
        get_uang_price.trace('w', itung_return)
        ttk.Label(newWindow4, 
                text ="Total Purchase:", font=("Arial", 17), anchor='e').grid(column = 1, 
                               row = 2,
                               padx = 10,
                               pady = 15)
        get_price = Entry(newWindow4, width=10, font="Arial 35", textvariable=total)
        get_price.grid(column = 2, row = 2, padx = 15, pady=10)
        ttk.Label(newWindow4, 
                text ="Return Money:", font=("Arial", 17), anchor='e').grid(column = 1, 
                               row = 3,
                               padx = 10,
                               pady = 15)
        return_price_label = Label(newWindow4, width = 10,font="Arial 35", anchor='w',textvariable=get_return_price)
        return_price_label.grid(column = 2, row = 3, padx = 15, pady=10)
        newWindow4.bind('<Return>', empty)

def press_pisan(event):
    global harga_item_dict
    loc_number = tree_cashier.focus()
    tree_cashier.delete(loc_number)
    del harga_item_dict[loc_number]
    totalin()

def change_final_price(pricing, key_val, total_potong):
    global harga_item_dict
    harga_item_dict[key_val] = [pricing, total_potong]
    totalin()

def masuk_item():
    data_poll = []
    search_kasir_input = search_kasir_entry.get()
    get_type_input = get_type_box.get()
    date_input = date_entry.get()

    def input_to_purchasing_spreadsheet():
        brp_byk = len(data_poll)
        cell_row = pembelian_worksheet.acell('P1').value
        pembelian_worksheet.update_acell('A'+str(cell_row), date_input)
        pembelian_worksheet.update_acell('B'+str(cell_row), search_kasir_input)
        pembelian_worksheet.update_acell('C'+str(cell_row), get_type_input)
        for i in range(brp_byk):
            pembelian_worksheet.update_acell('D'+str(int(cell_row)+i), data_poll[i][0])
            pembelian_worksheet.update_acell('E'+str(int(cell_row)+i), data_poll[i][1])
            pembelian_worksheet.update_acell('F'+str(int(cell_row)+i), data_poll[i][2])
            pembelian_worksheet.update_acell('G'+str(int(cell_row)+i), data_poll[i][5])
            pembelian_worksheet.update_acell('H'+str(int(cell_row)+i), data_poll[i][6])
            pembelian_worksheet.update_acell('I'+str(int(cell_row)+i), data_poll[i][7])
        pembelian_worksheet.update_acell('P1', int(cell_row) + brp_byk)
        

    if len(date_input) > 0 and len(search_kasir_input) > 0 and len(get_type_input) > 0:
        all_list = input_data_sheet()
        curr_num = find_last_row_sheet()
        same_num = 0
        row_same = 0

        global nomor_jumlah
        nomor_jumlah = 0

        for children in tree_pembelian.get_children():
            data_pisan = tree_pembelian.item(children)['values']
            data_poll.append(data_pisan)
        
            for i in range(0,curr_num-1):
                if str(data_pisan[0]) == str(all_list[i][3]):
                    same_num = same_num + 1
                    row_same = i
                    break
        
            if same_num >= 1:       #if there is an item that has the same number, same name, and same type, we add the stock to the existing data in the database
                cell = sheet.acell('F'+str(row_same+2)).value
                new_total = int(cell) + data_pisan[2]
                replace_data_sheet(row_same+2,data_pisan[0], data_pisan[5], data_pisan[8], new_total)
            else:
                cell_format_sheet()
                add_data_sheet(curr_num, data_pisan[0], data_pisan[1], data_pisan[5], data_pisan[4], data_pisan[8], data_pisan[2], data_pisan[3])
                update_database()
            same_num = 0
            row_same = 0
            tree_pembelian.delete(children)
        input_to_purchasing_spreadsheet()
        search_kasir_entry.delete(0, 'end')
        get_type_box.delete(0, 'end')
        date_entry.delete(0,'end')
        date_entry.insert(0,"dd/mm/yyyy")
    else:
        pass

def make_item():                                    # To add new item into database
    def itung_total(*args):
        try:
            price = int(get_price_var.get())
            jumlah = int(get_jumlah_var.get())
            diskon = get_percentage_var.get()
            total_price = float(((100-diskon)/100)*(price*jumlah))
            total_harga.set(total_price)
        except:
            total_harga.set('0')

    def cekin_detail():
        if float(total_harga.get()) > 0 and len(get_number.get()) > 0 and len(get_name.get()) > 0 and len(get_type.get()) > 0 and len(get_count.get()) > 0:
            global nomor_jumlah
            nomor_jumlah = nomor_jumlah + 1
            tree_pembelian.insert("", 'end', iid=nomor_jumlah, text=nomor_jumlah, values=(get_number.get(), get_name.get(), get_jumlah.get(), get_count.get(),
                                        get_type.get(), get_price_var.get(), get_percentage_var.get(), total_harga.get(), get_pricing.get()))
            newWindow.destroy()
        else:
            pass

    newWindow = Toplevel(window)                                
    newWindow.title("Add Information")
    ttk.Label(newWindow, 
          text ="Batch Number*", font=("Arial", 12)).grid(column = 0, 
                               row = 0,
                               padx = 15)
    
    get_number = Entry(newWindow, width=20, font="Arial 11")
    get_number.grid(column = 0, row = 1, padx = 15)

    ttk.Label(newWindow, 
          text ="Item Name*", font=("Arial", 12)).grid(column = 1, 
                               row = 0,
                               padx = 10,
                               pady = 5)

    get_name = Entry(newWindow, width=20, font="Arial 11")
    get_name.grid(column = 1, row = 1, padx = 15)

    ttk.Label(newWindow, 
          text ="Quantity*", font=("Arial", 12)).grid(column = 4, 
                               row = 0,
                               padx = 10,
                               pady = 5)
    get_jumlah_var = StringVar()
    get_jumlah = Entry(newWindow, width=10, font="Arial 12", textvariable=get_jumlah_var)
    get_jumlah.grid(column = 4, row = 1, padx = 15)
    get_jumlah_var.trace('w', itung_total)

    ttk.Label(newWindow, 
          text ="Type*", font=("Arial", 12)).grid(column = 2 , 
                               row = 0,
                               padx = 10,
                               pady = 5)

    get_type = ttk.Combobox(newWindow, values=["type_1",                          # To list all types of items, change it as you wish
                                                "type_2",
                                                "type_3",
                                                "type_4",
                                                "type_5",
                                                "type_6"], width=10)
    get_type.grid(column = 2, row = 1, padx = 15)

    ttk.Label(newWindow,
          text ="Quantifier*", font=("Arial", 12)).grid(column = 3 , 
                               row = 0,
                               padx = 10,
                               pady = 5)

    get_count = ttk.Combobox(newWindow, values=["qty_1",                           # To list all quantifier of items, change it as you wish
                                                "qty_2",
                                                "qty_3",
                                                "qty_4",
                                                "qty_5",
                                                "qty_6"], width=10)
    get_count.grid(column = 3, row = 1, padx = 15)

    ttk.Label(newWindow, 
          text ="Original Price*", font=("Arial", 12)).grid(column = 0, 
                               row = 2,
                               padx = 10,
                               pady = 5)

    get_price_var = StringVar()
    get_price = Entry(newWindow, width=10, font="Arial 11", textvariable=get_price_var)
    get_price.grid(column = 0, row = 3, padx = 15)
    get_price_var.trace('w', itung_total)

    ttk.Label(newWindow, 
          text ="Discount*", font=("Arial", 12)).grid(column = 1, 
                               row = 2,
                               padx = 10,
                               pady = 5)

    get_percentage_var = DoubleVar()
    get_percentage_var.set('0')
    get_percentage = Entry(newWindow, width=10, font="Arial 11", textvariable=get_percentage_var)
    get_percentage.grid(column = 1, row = 3, padx = 15)
    get_percentage_var.trace('w', itung_total)

    ttk.Label(newWindow, 
          text ="Selling Price*", font=("Arial", 12)).grid(column = 2, 
                               row = 2,
                               padx = 10,
                               pady = 5)

    get_pricing = Entry(newWindow, width=15, font="Arial 11")
    get_pricing.grid(column = 2, row = 3, padx = 15)

    ttk.Button(newWindow, text= 'Add Item', width=16, command=cekin_detail).grid(column =4, row= 4, pady = 20)

    total_harga = StringVar()
    total_harga.set('0')
    ttk.Label(newWindow, text ="Total Price", font="Arial 12 bold").grid(column = 0, row = 4, pady = 5)
    ttk.Label(newWindow, font="Arial 14 bold", textvariable= total_harga).grid(column = 1, row = 4, pady = 5, sticky=tkinter.W)


def print_data(event):
    def itung_nyusahin(event):
        final_qty = int(total_qty.get())
        final_price = int(total_price.get())
        iid_num = tree_cashier.focus()
        tree_cashier.delete(tree_cashier.focus())
        tree_cashier.insert("", 'end', iid=iid_num, text=iid_num,
            values=(sel_data[0], sel_data[1], str(final_qty), sel_data[3], str(price), str(final_price)))
        change_final_price(final_price, iid_num, final_qty)
        newWindow5.destroy()

    def count_pricing(*args):                   # Final checkout page before finish purchase
        try:
            harga_akhir = int(var.get())*price
        except:
            harga_akhir = '-'
        price_var.set(harga_akhir)

    if tree_cashier.focus() == '':
        pass
    else:
        sel_data = tree_cashier.item(tree_cashier.focus())['values']
        price = int(sel_data[4])

        newWindow5 = Toplevel(window)
        newWindow5.title("Electronic Register and Management")

        var = StringVar()
        var.set('1')
        total_qty = Entry(newWindow5, width=10, font="Arial 14 bold", textvariable=var)
        total_qty.grid(column = 1, row = 0, padx = 10, pady = 5)
        var.trace('w', count_pricing)
        ttk.Label(newWindow5, text ="Total", font=("Arial", 14)).grid(column = 0, 
                               row = 0)
        price_var = StringVar()
        price_var.set(str(price))
        total_price = Entry(newWindow5, width=10, font="Arial 14 bold", textvariable=price_var)
        total_price.grid(column = 1, row = 1, padx = 10, pady=5)
        ttk.Label(newWindow5, text ="Total", font=("Arial", 14)).grid(column = 0, 
                               row = 1)
        newWindow5.bind('<Return>', itung_nyusahin)

def OnDoubleClick(event):
        if tree_database.focus() == '':
            return None
        sel_data = tree_database.item(tree_database.focus())['values']
        sel_data[5] = sel_data[5].replace('.', "").replace('Rp', "")
        tree_cashier.insert("", 'end', iid=tree_database.focus(), text=tree_database.focus(),
            values=(sel_data[0], sel_data[1], str(1), sel_data[3], sel_data[5], int(sel_data[5])))
        change_final_price(int(sel_data[5]), tree_database.focus(), 1)
        search_kasir.delete(0, END)

def delete_data():
    if tree.focus() == '':
        return None
    selected_row = int(tree.focus())+1
    delete_row_sheet(selected_row)
    update_database()

def update_database():
    list_of_data = input_data_sheet()
    #delete items in treeview
    for item in tree.get_children():
      tree.delete(item)
    for item in tree_database.get_children():
      tree_database.delete(item)

    #input the new data inside treeview
    for data in list_of_data:
        tree_database.insert("", 'end', iid=data[0], text=data[0],
            values=(data[3], data[1], data[5], data[4], data[2], data[6]))
        tree.insert("", 'end', iid=data[0], text=data[0],
            values=(data[3], data[1], data[5], data[4], data[6], data[7]))

def search_database_pembelian(*args):
    itemsOnTreeView = tree_database.get_children()

    for eachItem in itemsOnTreeView:
        if search_entry_var_kasir.get().lower() in tree_database.item(eachItem)['values'][1].lower():   # compare strings in lower cases
            search_var = tree_database.item(eachItem)['values']
            tree_database.delete(eachItem)
            tree_database.insert("",0, iid=eachItem, text=eachItem, values=search_var)


def search_database(*args):
    itemsOnTreeView = tree.get_children()

    for eachItem in itemsOnTreeView:
        if search_entry_var.get().lower() in tree.item(eachItem)['values'][1].lower():   # compare strings in  lower cases
            search_var = tree.item(eachItem)['values']
            tree.delete(eachItem)
            tree.insert("",0, iid=eachItem, text=eachItem, values=search_var)

def retur_barang():                                     # The ability to return goods after purchase
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")

    def change_qty(event):
        if len(qty.get()) > 0:
            data_arr = [list_of_data[3], list_of_data[1], int(qty.get()), int(price.get())*-1]
            reduce_qty_sheet(int(selected_item)+1, int(qty.get())*-1)
            print_data_sheet(str(current_time), 1, data_arr, int(price.get())*-1)
            update_database()
            newWindow6.destroy()
    
    def change_final_price(*args):
        if qty.get() == '':
            price.set('0')
        else:
            harga_akhir = int(qty.get())*int(current_price)
            price.set(str(harga_akhir))
    
    if tree_database.focus() == '':
        return None

    selected_item = tree_database.focus()
    list_of_data = find_data_sheet(int(selected_item)+1)
    newWindow6 = Toplevel(window)
    newWindow6.title("Retur Barang")

    qty = StringVar()
    qty.set('1')
    total_qty = Entry(newWindow6, width=10, font="Arial 14 bold", textvariable=qty)
    total_qty.grid(column = 1, row = 0, padx = 10, pady = 5)
    ttk.Label(newWindow6, text ="Jumlah", font=("Arial", 14)).grid(column = 0, row = 0)
    qty.trace('w', change_final_price)

    price = StringVar()
    current_price = list_of_data[6].replace('.', "").replace('Rp', "")
    price.set(current_price)
    total_price = Entry(newWindow6, width=10, font="Arial 14 bold", textvariable=price)
    total_price.grid(column = 1, row = 1, padx = 10, pady=5)
    ttk.Label(newWindow6, text ="Total", font=("Arial", 14)).grid(column = 0, row = 1)
    newWindow6.bind('<Return>', change_qty)


def change_data():                                      # Change established data in the database
    if tree.focus() == '':
        return None
    selected_item = tree.focus()
    list_of_data = find_data_sheet(int(selected_item)+1)

    def find_final_price(*args):
        try:
            final_value = int(float(get_price_var.get())+(float(get_price_var.get())*float(get_percentage_var.get())*0.01))
            get_final_price.delete(0, END)
            get_final_price.insert(0,str(final_value))
        except:
            get_final_price.delete(0, END)
            get_final_price.insert(0,'-')

    def input_data():
        goods_num = get_number.get()
        goods_price = int(get_price.get())
        final_price = int(get_final_price.get())
        stok_amount = int(get_stok.get())

        if len(goods_num) != 0 and goods_price > 0 and final_price >= 0 and stok_amount >= 0:
            replace_data_sheet(int(selected_item)+1, goods_num, goods_price, final_price, stok_amount)
            newWindow.destroy()
            update_database()
        else:
            newWindow2 = Toplevel(newWindow)
            newWindow2.title("Peringatan!")
            ttk.Label(newWindow2, 
                text ="Pastikan anda sudah mengisi semua\n         kotak data yang diperlukan !", font=("Arial", 12), anchor='center').grid(column = 2, 
                               row = 2,
                               padx = 50,
                               pady = 50)

    newWindow = Toplevel(window)
    newWindow.title("Edit Informasi Barang")
    ttk.Label(newWindow, 
          text ="Nomor Batch*", font=("Arial", 12)).grid(column = 0, 
                               row = 0,
                               padx = 15)
    
    get_number = Entry(newWindow, width=20, font="Arial 11")
    get_number.insert(0, list_of_data[3])
    get_number.grid(column = 0, row = 1, padx = 15)

    ttk.Label(newWindow, 
          text ="Nama barang*", font=("Arial", 12)).grid(column = 1, 
                               row = 0,
                               padx = 10,
                               pady = 5)

    get_name = Entry(newWindow, width=20, font="Arial 11", state='disabled')
    get_name.grid(column = 1, row = 1, padx = 15)
    get_name.insert(0, list_of_data[1])
    
    ttk.Label(newWindow, 
          text ="Harga Jual*", font=("Arial", 12)).grid(column = 2, 
                               row = 2,
                               padx = 10,
                               pady = 5)

    get_final_price = Entry(newWindow, width=10, font="Arial 12")
    get_final_price.grid(column = 2, row = 3, padx = 15)
    get_final_price.insert(0, list_of_data[6].replace('.', "").replace('Rp', ""))

    ttk.Label(newWindow, 
          text ="Tipe*", font=("Arial", 12)).grid(column = 2 , 
                               row = 0,
                               padx = 10,
                               pady = 5)

    get_type = ttk.Combobox(newWindow, width=10, state='disabled')
    get_type.grid(column = 2, row = 1, padx = 15)

    ttk.Label(newWindow,
          text ="Satuan*", font=("Arial", 12)).grid(column = 3 , 
                               row = 0,
                               padx = 10,
                               pady = 5)

    get_count = ttk.Combobox(newWindow, width=10, state='disabled')
    get_count.grid(column = 3, row = 1, padx = 15)

    ttk.Label(newWindow, 
          text ="Harga Pokok*", font=("Arial", 12)).grid(column = 0, 
                               row = 2,
                               padx = 10,
                               pady = 5)

    get_price_var = StringVar()
    get_price = Entry(newWindow, width=10, font="Arial 11", textvariable= get_price_var)
    get_price.insert(0, list_of_data[7].replace('.', "").replace('Rp', ""))
    get_price.grid(column = 0, row = 3, padx = 15)
    get_price_var.trace('w', find_final_price)

    ttk.Label(newWindow, 
          text ="Persentase (%)", font=("Arial", 12)).grid(column = 1, 
                               row = 2,
                               padx = 10,
                               pady = 5)

    get_percentage_var= StringVar()
    get_percentage = Entry(newWindow, width=10, font="Arial 11", textvariable= get_percentage_var)
    get_percentage.grid(column = 1, row = 3, padx = 15)
    get_percentage.insert(0, '0')
    get_percentage_var.trace('w', find_final_price)

    ttk.Label(newWindow, 
          text ="Stok*", font=("Arial", 12)).grid(column = 3, 
                               row = 2,
                               padx = 10,
                               pady = 5)

    get_stok = Entry(newWindow, width=10, font="Arial 11")
    get_stok.grid(column = 3, row = 3, padx = 15)
    get_stok.insert(0, list_of_data[5])

    ttk.Button(newWindow, 
            text= 'Ubah data', width=12, command=input_data).grid(column =3,
                                row= 4,
                                pady = 20)


###########################################################
# Connect to Google Sheets
scope = ['https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive"]

credentials = ServiceAccountCredentials.from_json_keyfile_name(resource_path("#your_credentials_json_file_name.json"), scope)        # To link gspread to google sheet, put your credentials.json here
client = gspread.authorize(credentials)

sheet = client.open("#your_database_sheet_name").sheet1                             # The sheet to be used as database, be sure that this sheet already exist, change as you wish

textsheet = client.open("#your_sales_sheet_name")                                # The sheet to store sales data, change as you wish

textsheet_pembelian = client.open("#your_purchasing_sheet_name")                      # The sheet to store purchasing data, change as you wish

start_up_sequence_sheet()

today = date.today()

window = tkinter.Tk()
window.title("Electronic Register and Management")
tabControl = ttk.Notebook(window, padding=10)

tab1 = ttk.Frame(tabControl)
tab2 = ttk.Frame(tabControl)
tab3 = ttk.Frame(tabControl)
tab4 = ttk.Frame(tabControl)

tabControl.add(tab1, text ='Cashier')
tabControl.add(tab2, text ='Database')
tabControl.add(tab3, text ='History')
tabControl.add(tab4, text ='Buying')
tabControl.pack(expand = 1, fill ="both")

style = ttk.Style()
style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Arial', 12)) # Modify the font of the body
style.configure("mystyle.Treeview.Heading", font=('Arial', 13,'bold')) # Modify the font of the headings
style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})]) # Remove the borders

total = StringVar()
total.set('0')

search_entry_var_kasir = StringVar()
search_kasir = tkinter.Entry(tab1, width=30, font="Arial 14", textvariable=search_entry_var_kasir)
search_kasir.grid(column = 1, row = 0, padx = 10, pady = 30)
search_entry_var_kasir.trace('w', search_database_pembelian)

total_label = ttk.Label(tab1, font="Arial 45", textvariable=total)
total_label.grid(column = 4, row = 0, padx = 10, pady = 30)

columns = ('nomor_barang' ,'nama_barang', 'qty', 'satuan','tipe','harga')
tree_database = ttk.Treeview(tab1, columns=columns, show="headings", style="mystyle.Treeview")
tree_database.heading('nomor_barang', text= 'Batch Number')
tree_database.heading('nama_barang', text= 'Item Name')
tree_database.heading('qty', text= 'Total')
tree_database.heading('satuan', text='Qty.')
tree_database.heading('harga', text= 'Price per Piece')
tree_database.heading('tipe', text = 'Type')
tree_database.column('harga', width=160)
tree_database.column('satuan', width=120)
tree_database.column('qty', width=80)
tree_database.column('tipe', width=160)
tree_database['height']=5

tree_database.grid(row = 1, column=0, columnspan=5, pady=10, padx=20)
tree_database.bind("<Double-1>", OnDoubleClick)

list_of_data = input_data_sheet()

#input the new data inside tree_database treeview
for data in list_of_data:
    tree_database.insert("", 'end', iid=data[0], text=data[0], values=(data[3], data[1], data[5], data[4], data[2], data[6]))

columns = ('nomor_barang' ,'nama_barang', 'qty', 'satuan','harga', 'harga_total')
tree_cashier = ttk.Treeview(tab1, columns=columns, show='headings', style="mystyle.Treeview")
tree_cashier.heading('nomor_barang', text= 'Batch Number')
tree_cashier.heading('nama_barang', text= 'ItemName')
tree_cashier.heading('qty', text= 'Qty.')
tree_cashier.heading('satuan', text='Quantifier')
tree_cashier.heading('harga', text= 'Price per Piece')
tree_cashier.heading('harga_total', text= 'Total Price')
tree_cashier.column('satuan', width=120)
tree_cashier.column('harga', width=160)
tree_cashier.column('harga_total', width=160)
tree_cashier.column('qty', width=80)
tree_cashier['height']=7
tree_cashier.bind("<Double-1>", print_data)
tree_cashier.bind('<BackSpace>', press_pisan)
window.bind('<Escape>', finishing)

tree_cashier.grid(row = 5, column=0, columnspan=5, padx=20, pady=10)

ttk.Button(tab1, text= 'Return Item', width=16, command=retur_barang).grid(column = 4, row= 2, pady = 20, padx=25, sticky = tkinter.E)

ttk.Label(tab2, text ="DATABASE\n#your_institution_name", font=("Arial 17 bold")).grid(column = 3, row = 0, pady = 20)

search_entry_var = StringVar()
search_entry = tkinter.Entry(tab2, width=30, font="Arial 14", textvariable=search_entry_var)
search_entry.grid(column = 1, row = 0, pady = 20)
search_entry_var.trace('w', search_database)

ttk.Button(tab2, text= 'Delete Item', width=16, command=delete_data).grid(column =1, row= 6, pady = 20, sticky=tkinter.W)

ttk.Button(tab2, text= 'Change Item Info', width=16, command=change_data).grid(column =2, row= 6, pady = 20, sticky=tkinter.W)

ttk.Button(tab2, text= 'Check Database', width=16, command=update_database).grid(column = 3, row= 6, pady = 20)

columns = ('nomor_barang' ,'nama_barang', 'qty', 'satuan','harga', "harga_pokok")
tree = ttk.Treeview(tab2, columns=columns, show="headings", style="mystyle.Treeview")
tree.heading('nomor_barang', text= 'Batch Number')
tree.heading('nama_barang', text= 'Item Name')
tree.heading('qty', text= 'Qtv.')
tree.heading('satuan', text='Quantifier')
tree.heading('harga', text= 'Sell Price')
tree.heading('harga_pokok', text = 'Buying Price')
tree.column('nomor_barang', width=140)
tree.column('satuan', width=100)
tree.column('qty', width=80)
tree['height']=20

tree.grid(row = 5, column=0, columnspan=5, padx=20, pady=10)

income_total = StringVar()
income_total.set('Rp.0')

columns = ('waktu' ,'tot_jml_barang', 'harga_total')
histori_tree = ttk.Treeview(tab3, columns=columns, show="headings", style="mystyle.Treeview")
histori_tree.heading('waktu', text= 'Time')
histori_tree.heading('tot_jml_barang', text= 'Total Items')
histori_tree.heading('harga_total', text= 'Total Price')
histori_tree['height']= 25


date_entry = tkinter.Entry(tab4, width=10, font="Arial 13")
date_entry.grid(row = 0, column=0, pady=15)
date_entry.insert(0, "dd/mm/yyyy")

search_kasir_entry = tkinter.Entry(tab4, width=25, font="Arial 13")
search_kasir_entry.grid(row = 0, column=2, pady=15)

get_type_box = ttk.Combobox(tab4, values=["APL","AAM","Alida"], width=15)
get_type_box.grid(column = 1, row = 0, pady = 15)

columns = ('nomor_barang','nama_barang', 'qty', 'satuan','tipe','harga_pokok', 'disc', 'harga_total','harga_satuan')
tree_pembelian = ttk.Treeview(tab4, columns=columns, show="headings", style="mystyle.Treeview")
tree_pembelian.heading('nomor_barang', text= 'Batch Number')
tree_pembelian.heading('nama_barang', text= 'Item Name')
tree_pembelian.heading('qty', text= 'Qty.')
tree_pembelian.heading('satuan', text='Quantifier')
tree_pembelian.heading('tipe', text='Type')
tree_pembelian.heading('disc', text= 'Discount')
tree_pembelian.heading('harga_pokok', text ='Buying Price')
tree_pembelian.heading('harga_total', text ='Total Price')
tree_pembelian.heading('harga_satuan', text ='Sell Price')
tree_pembelian.column('nama_barang', width = 160)
tree_pembelian.column('harga_pokok', width=140)
tree_pembelian.column('harga_total', width=140)
tree_pembelian.column('harga_satuan', width=140)
tree_pembelian.column('tipe', width=90)
tree_pembelian.column('disc', width= 60)
tree_pembelian.column('nomor_barang', width=80)
tree_pembelian.column('satuan', width=60)
tree_pembelian.column('qty', width=60)
tree_pembelian['height']=10
tree_pembelian.grid(row = 1, column=0, columnspan=3, padx=20, pady=10)

ttk.Button(tab4, text= 'Add Item', width=16, command=make_item).grid(column =1, row= 2, pady = 20)
ttk.Button(tab4, text= 'Finish Input', width=16, command=masuk_item).grid(column =2, row= 2, pady = 20)

histori_tree.grid(row = 0, column=0, columnspan=3, padx=20, pady=10)

date_label = Label(tab3, text ="SALES HISTORY\n#your_institution_name\n"+ str(today.strftime("%d-%b-%Y")), font=("Arial 22 bold"), anchor='n')
date_label.grid(column = 3, row = 0, pady = 20)

income_label = ttk.Label(tab3, font="Arial 32", textvariable=income_total)
income_label.grid(column = 1, row = 1)

ttk.Button(tab3, text= 'Check History', width=20, command=check_histori).grid(column = 3, row= 1, pady = 20)

window.protocol('WM_DELETE_WINDOW',on_close)
window.mainloop()