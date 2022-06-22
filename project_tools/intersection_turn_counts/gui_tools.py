from tkinter import messagebox
from tkinter import *
from tkinter import ttk
from PIL import ImageTk, Image


def add_labels(frame, labels, row_start=0, col_start=0, row_increment=1, col_increment=0, font=None):
    row = row_start
    col = col_start
    for l in labels:
        Label(frame, text=l, font=font).grid(column=col, row=row)
        row += row_increment
        col += col_increment
    return


def add_select_box(frame, initial_vars, variables_list, row_start=0, col_start=0, row_increment=1, col_increment=0):
    row = row_start
    col = col_start
    variables = []
    for var in initial_vars:
        var_str = StringVar()
        var_str.set(var)
        variables.append(var_str)
    for i, v in enumerate(variables):
        OptionMenu(frame, variables[i], *variables_list).grid(row=row, column=col)
        row += row_increment
        col += col_increment
    return variables


def get_approaches(variables, master):
    output_list = []
    for v in variables:
        output_list.append(v.get())
    master.destroy
    return output_list


def set_apply(apply):
    apply = False
    return apply


def quit_and_log(master):
    quit_master = messagebox.askyesnocancel("Quit",
                                            "Do you want to quit?  No data will be extracted and the site will "
                                            "be added to the log.")
    if quit_master is None:
        print('ddddd')


    elif quit_master:
        print('aaaaa')
        master.destroy()
    else:
        print('cccccc')
        master.destroy()


def apply_on_closing(master):
    if messagebox.askokcancel("Quit", "Do you want to extract data and quit? "):
        master.destroy()


def size_image_by_width(image_, width_target=900):
    width = image_.width
    height = image_.height
    image_ = image_.resize((int(width_target), int(width_target / width * height)), Image.LANCZOS)

    return image_


def run_traffic_survey_gui(xl_image, map_image):
    direction_list = ['N', 'NE', 'E', 'SE', 'S', 'SW', 'W', 'NW']

    movements = {'1': 'N_N', '2': 'N_S'}
    movements_excel_list = ['1', '2', '3', '4']
    approach_from_list = ['N', 'S', 'E', 'W']
    approach_to_list = ['N', 'S', 'E', 'W']
    img_path = r"C:\Users\matth\OneDrive\Pictures\Camera Roll\20220511_004853.jpg"

    master = Tk()
    master.title('intersection movement analysis')
    frm = ttk.Frame(master, padding=10)
    frm.grid(row=0, column=0, sticky='n')
    frm2 = ttk.Frame(master, padding=10)
    frm2.grid(row=1, column=0, sticky='s')
    variable = StringVar(frm)

    # w.pack()
    grid_location = 1

    add_labels(frm, labels=['xl mvt', 'i', 'j', 'k'], row_start=0, col_start=0, row_increment=0, col_increment=1,
               font='bold')
    add_labels(frm, labels=movements_excel_list, row_start=1, col_start=0, row_increment=1, col_increment=0)
    variables_from = add_select_box(frm, initial_vars=approach_from_list, variables_list=direction_list,
                                    row_start=1, col_start=1, row_increment=1, col_increment=0)
    variables_to = add_select_box(frm, initial_vars=approach_from_list, variables_list=direction_list,
                                  row_start=2, col_start=1, row_increment=1, col_increment=0)
    # Button(frm, text="Extract Data", command=get_approaches(variables_from)).grid(row=6, column=5)
    Button(frm2, text="Cancel", command=quit_and_log).grid(row=1, column=1)
    Button(frm2, text="Extract Data", command=apply_on_closing).grid(row=0, column=1)
    # image1 = img.resize((800,200), Image.LANCZOS)
    xl_image = size_image_by_width(xl_image, width_target=500)
    map_image = size_image_by_width(map_image, width_target=700)
    xl_image_tk = ImageTk.PhotoImage(xl_image, master=master)
    map_image_tk = ImageTk.PhotoImage(map_image, master=master)
    label1 = Label(master, image=xl_image_tk).grid(row=0, column=1)
    label2 = Label(master, image=map_image_tk).grid(row=1, column=1)
    # Position image
    # label1.place(x=0, y=0)

    master.mainloop()

    print('done')
    '''
    for movement, item in movements.items():
        Label(frm, text=movement).grid(column=1, row=grid_location)
        variable.set(item)  # default value
        w = OptionMenu(frm, variable, *direction_list)
        w.grid(row=grid_location, column=2)
        grid_location += 1
        # buttons.append(btn)  # adding button reference
    '''
    return

