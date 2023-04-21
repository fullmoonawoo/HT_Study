from datetime import *
import shutil
import xlsxwriter
import openpyxl as op
import tkinter as tk
from tkinter import *
from tkinter import messagebox
import os


HEIGHT = 750
WIDTH = 1000

root = tk.Tk()
root.title("BODYCOTE HT Study")
sc_width = root.winfo_screenwidth()
sc_height = root.winfo_screenheight()
x = (sc_width / 2) - (WIDTH / 2)
y = (sc_height / 3) - (HEIGHT / 3)
root.geometry(('%dx%d+%d+%d' % (WIDTH, HEIGHT, x, y)))


bg_color = "gray44"
button_color = "gray64"
font_color = "gray4"
font_color2 = "gray4"
root.config(bg=bg_color)
root.rowconfigure(1, weight=1)
root.columnconfigure(7, weight=1)

# Frames
header = Frame(root, height=200, width=1000)
header.grid(row=0, column=0)

background_image = tk.PhotoImage(file="C:\\Users\\lukas\\python_files\\HT_Study\\bc_htstudy_back.png")
background_label = tk.Label(header, image=background_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

workspace = Frame(root, height=400, width=1000, bg=bg_color)
workspace.grid(row=1, column=0, sticky="WE", ipady=10, pady=4)
workspace.columnconfigure(6, weight=1)
workspace.rowconfigure(19, weight=1)

cellwidth = 35
cellheight = 1

# Date and time
date = datetime.today()
finaldate = date.strftime("%H:%M:%S / %b-%d-%Y")
year = date.strftime("%Y")

# Destinations
serveraddress = "C:\\Users\\lukas.rausa.EUROPE\\OneDrive - BODYCOTE\\Protokoly\\" + str(year) + "\\LINAMAR LTH\\"
#serveraddress = "C:\\Users\\majster.vlkanova\\OneDrive - BODYCOTE\\Protokoly\\" + str(year) + "\\LINAMAR LTH\\"
localaddress = "C:\\Users\\lukas\\python_files\\HT_Study\\archive\\"
#localaddress = "C:\\Users\\majster.vlkanova\\ht_docs\\"

# Widgets
part_label = tk.Label(workspace, text="Vyber typ dielu:", bg=bg_color, fg=font_color)
part_label.grid(row=4, column=1, padx=4, pady=10, ipady=2)

part_type = IntVar()
P1 = Radiobutton(workspace, text="Output Gear High D", variable=part_type, value=1, bg=bg_color, fg=font_color)
P1.grid(row=4, column=2, pady=4, ipady=2)

P2 = Radiobutton(workspace, text="Output Gear Low D", variable=part_type, value=2, bg=bg_color, fg=font_color)
P2.grid(row=5, column=2, pady=4, ipady=2)


bc_number = tk.Label(workspace, text="Číslo sprievodky:", width=13, bg=bg_color, fg=font_color)
bc_number.grid(row=6, column=1, columnspan=1, padx=4, ipady=2)
nnote = tk.Entry(workspace, width=int(cellwidth))
nnote.grid(row=6, column=2, padx=4)

for x in range(1, 11):
    position_label = Label(workspace, text="Pozícia {}".format(x), width=10, bg=bg_color, fg=font_color)
    position_label.grid(padx=6, pady=4, ipadx=20, ipady=1, row=6+x, column=1)

position_container = []

for r in range(1, 11):
    for c in range(1, 4):
        position_entry = tk.Entry(workspace, width=int(cellwidth))
        position_entry.grid(row=6 + r, column=1 + c, padx=4)
        position_container.append(position_entry)


# Functions
def save_values():
    nnoteresult = str(nnote.get())
    fnnoteresult = nnoteresult.replace("/", "")
    parttyperesult = int(part_type.get())
    if parttyperesult == int(0):
        messagebox.showinfo("Information", "Vyber typ dielu !")
        exit(mainloop())

    # Creating of excel file
    workbook = xlsxwriter.Workbook(str(localaddress) + fnnoteresult + "-HT Study.xlsx")
    worksheet_htstudy = workbook.add_worksheet("HT Study")
    filetitle = fnnoteresult+"-HT Study.xlsx"

    # Setting the sizes of cells
    cellposition = ("A:A", "B:B", "C:C", "D:D", "E:E", "F:F", "G:G", "H:H")
    colwidth = (2.14, 4, 36.43, 11.43, 12.57, 39.71, 4.71, 2.14)

    for (x, y) in zip(cellposition, colwidth):
        worksheet_htstudy.set_column(x, y)

    rowheight = (
        15.75, 65, 6, 18, 36, 18.75, 15.75, 15, 15, 45, 24, 24, 24, 24, 24, 24, 24, 24, 24, 24, 24, 24, 24, 24, 24, 24, 24, 24, 24, 24, 24, 24,
        24, 24, 24, 24, 24, 24, 24, 24, 15, 15.75, 15.75
    )

    for (r, h) in zip(range(45), rowheight):
        worksheet_htstudy.set_row(r, h)

    worksheet_htstudy.set_paper(9)
    worksheet_htstudy.fit_to_pages(8, 43)
    worksheet_htstudy.set_margins(0.1, 0.1, 0.1, 0.1)
    worksheet_htstudy.print_area(0, 0, 42, 7)
    worksheet_htstudy.set_print_scale(80)
    worksheet_htstudy.center_horizontally()
    worksheet_htstudy.center_vertically()

    # Outside Borders
    bborder_format = workbook.add_format({'bottom': 6})
    rborder_format = workbook.add_format({'right': 6})
    lborder_format = workbook.add_format({'left': 6})

    for rb in range(1, 7):
        worksheet_htstudy.write(41, rb, '', bborder_format)

    for cb in range(2, 42):
        worksheet_htstudy.write(cb, 0, '', rborder_format)
        worksheet_htstudy.write(cb, 7, '', lborder_format)

    # Writing titles and formating
    basetitle_format = workbook.add_format(
        {'font_name': 'Calibri', 'bold': 1, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    base_format = workbook.add_format(
        {'font_name': 'Calibri', 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1, 'border': 1})
    bottom_format = workbook.add_format(
        {'font_name': 'Calibri', 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'text_wrap': 1, 'bottom': 1})
    datecell_format = workbook.add_format(
        {'font_name': 'Calibri', 'font_size': 11, 'align': 'right', 'valign': 'vcenter', 'text_wrap': 1, 'bold': 1})
    main_title = workbook.add_format({
        'bold': 1,
        'font_size': 22,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': 'wrap',
        'border': 6
    })

    worksheet_htstudy.merge_range('B2:G2', '                                     LINAMAR Heat treatment study sample parts', main_title)
    worksheet_htstudy.merge_range('D4:F4', '', base_format)
    worksheet_htstudy.merge_range('D5:F5', '', base_format)
    worksheet_htstudy.merge_range('D6:F6', '', base_format)
    worksheet_htstudy.merge_range('D7:F7', '', base_format)

    # Order of arguments r-c-t-f
    worksheet_htstudy.write(3, 2, 'Drawing number:', basetitle_format)
    worksheet_htstudy.write(4, 2, 'Part name:', basetitle_format)
    worksheet_htstudy.write(5, 2, 'Furnace ID:', basetitle_format)
    worksheet_htstudy.write(6, 2, 'Heat Treated Qty:', basetitle_format)
    worksheet_htstudy.write(9, 2, 'HT study Part ID marking', basetitle_format)
    worksheet_htstudy.write(9, 3, 'Furnace position (1-10)', base_format)
    worksheet_htstudy.write(10, 3, '1', base_format)
    worksheet_htstudy.write(11, 3, '1', base_format)
    worksheet_htstudy.write(12, 3, '1', base_format)
    worksheet_htstudy.write(13, 3, '2', base_format)
    worksheet_htstudy.write(14, 3, '2', base_format)
    worksheet_htstudy.write(15, 3, '2', base_format)
    worksheet_htstudy.write(16, 3, '3', base_format)
    worksheet_htstudy.write(17, 3, '3', base_format)
    worksheet_htstudy.write(18, 3, '3', base_format)
    worksheet_htstudy.write(19, 3, '4', base_format)
    worksheet_htstudy.write(20, 3, '4', base_format)
    worksheet_htstudy.write(21, 3, '4', base_format)
    worksheet_htstudy.write(22, 3, '5', base_format)
    worksheet_htstudy.write(23, 3, '5', base_format)
    worksheet_htstudy.write(24, 3, '5', base_format)
    worksheet_htstudy.write(25, 3, '6', base_format)
    worksheet_htstudy.write(26, 3, '6', base_format)
    worksheet_htstudy.write(27, 3, '6', base_format)
    worksheet_htstudy.write(28, 3, '7', base_format)
    worksheet_htstudy.write(29, 3, '7', base_format)
    worksheet_htstudy.write(30, 3, '7', base_format)
    worksheet_htstudy.write(31, 3, '8', base_format)
    worksheet_htstudy.write(32, 3, '8', base_format)
    worksheet_htstudy.write(33, 3, '8', base_format)
    worksheet_htstudy.write(34, 3, '9', base_format)
    worksheet_htstudy.write(35, 3, '9', base_format)
    worksheet_htstudy.write(36, 3, '9', base_format)
    worksheet_htstudy.write(37, 3, '10', base_format)
    worksheet_htstudy.write(38, 3, '10', base_format)
    worksheet_htstudy.write(39, 3, '10', base_format)
    worksheet_htstudy.write(13, 5, 'Description (looking into open door)', basetitle_format)
    worksheet_htstudy.write(14, 5, '1 = left bottom front', base_format)
    worksheet_htstudy.write(15, 5, '2 = right bottom front', base_format)
    worksheet_htstudy.write(16, 5, '3 = left bottom back', base_format)
    worksheet_htstudy.write(17, 5, '4 = right bottom back', base_format)
    worksheet_htstudy.write(18, 5, '5 = middle of batch', base_format)
    worksheet_htstudy.write(19, 5, '6 = left top front', base_format)
    worksheet_htstudy.write(20, 5, '7 = right top front', base_format)
    worksheet_htstudy.write(21, 5, '8 = left top back', base_format)
    worksheet_htstudy.write(22, 5, '9 = right top back', base_format)
    worksheet_htstudy.write(23, 5, '10 = top shelf middle', base_format)
    worksheet_htstudy.write(40, 4, "Date:", datecell_format)
    # Image
    worksheet_htstudy.insert_image('F26', 'pozicie.png', {'x_offset': 5, 'y_offset': 5} and {'x_scale': 0.7, 'y_scale': 0.7})
    worksheet_htstudy.insert_image('B2', 'linamar_logo.png', {'x_offset': 4, 'y_offset': 5})
    # Write values and save
    if parttyperesult == 1:
        worksheet_htstudy.write(4, 3, str("Output gear High D"), basetitle_format)
    if parttyperesult == 2:
        worksheet_htstudy.write(4, 3, str("Output gear Low D"), basetitle_format)

    worksheet_htstudy.write(6, 3, str(nnoteresult), base_format)

    for idx, numb in enumerate(position_container):
        worksheet_htstudy.write(10 + idx, 2, str(numb.get()), base_format)

    worksheet_htstudy.write(40, 5, str(finaldate), bottom_format)
    workbook.close()
    if parttyperesult == 1:
        partname = str(" (Output Gear High D)")
        foldertitle = fnnoteresult + partname + "-HT Study"
        os.mkdir(str(serveraddress) + str(foldertitle))
        src_path = str(localaddress) + filetitle
        dst_path = str(serveraddress) + str(foldertitle) + '\\' + filetitle
        shutil.copy(src_path, dst_path)
    if parttyperesult == 2:
        partname = str(" (Output Gear Low D)")
        foldertitle = fnnoteresult + partname + "-HT Study"
        os.mkdir(str(serveraddress) + str(foldertitle))
        src_path = str(localaddress) + filetitle
        dst_path = str(serveraddress) + str(foldertitle) + '\\' + filetitle
        shutil.copy(src_path, dst_path)
    save_button["state"] = DISABLED
    print_button["state"] = ACTIVE


def clearcell():
    nnote.delete(0, END)

    for entry in position_container:
        entry.delete(0, END)

    save_button["state"] = ACTIVE
    print_button["state"] = DISABLED
    return part_type.set(int(0))


def printfile():
    nnoteresult = str(nnote.get())
    fnnoteresult = nnoteresult.replace("/", "")
    wb2 = op.load_workbook(str(localaddress)+str(fnnoteresult)+"-HT Study.xlsx")
    ws = wb2.active
    ws.title = str(fnnoteresult)+"-HT Study.xlsx"
    os.startfile(str(localaddress)+ws.title, "print")


# Buttons
save_button = tk.Button(workspace, text="Uložiť HT Study", command=save_values, width=18, height=4, bg=button_color, fg=font_color2)
save_button.grid(row=19, column=4, padx=4, pady=4, sticky="WE")

clear_button = tk.Button(workspace, text="Vyčistiť bunky !", command=clearcell, width=18, height=1, bg=button_color,
                         fg=font_color2)
clear_button.grid(row=7, column=5, padx=4, pady=4, sticky="WE")

print_button = tk.Button(workspace, text="Vytlačiť", command=printfile, width=18, height=4, bg=button_color,
                         fg=font_color2)
print_button.grid(row=19, column=2, padx=4, pady=4, sticky="WE")
print_button["state"] = DISABLED

root.mainloop()
