import configparser
import logging
import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import script

# Initialize logging
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler("debug.log"), logging.StreamHandler()],
)

# Intialize config parser
config = configparser.ConfigParser()
config.read("config.ini")
initial_dir = config.get("Settings", "ICP_DB_FOLDER_PATH", fallback=os.getcwd())

# Initialize global variables
db_file_path = None
selected_result_indexes = []
conn = None
filtered_db_file_path = None

# Initialize tkinter GUI
root = tk.Tk()
root.geometry("500x600")  # set screen dimensions
root.pack_propagate(False)  # Do not auto adjust window size.
root.resizable(0, 0)  # Disable window resizing.
root.iconbitmap("favicon.ico")
root.title("ICP Results Filter and Extration")

# Frame for open file dialog
frm_file = tk.LabelFrame(root, text="Open ICP Results Database File")
frm_file.place(height=70, width=450, relx=0.5, rely=0.15, anchor=tk.CENTER)

# Frame for Data and TreeView
frm_dataview = tk.LabelFrame(root, text="Select data to be filtered out")
frm_dataview.place(height=300, width=500, rely=0.20)

# Frame for selecting data to exclude
frm_selected_data = tk.LabelFrame(
    root,
    text="Selected Data (The following indexes will be filtered and removed)",
)
frm_selected_data.place(height=75, width=450, relx=0.5, rely=0.85, anchor=tk.CENTER)

# Buttons
btn_browse = tk.Button(
    frm_file,
    text="Browse File",
    command=lambda: loadDataToTreeView(openFileDialog()),
)
btn_browse.pack(side=tk.BOTTOM, anchor="e", padx=8, pady=8)
btn_select_data = tk.Button(
    root,
    text="Exclude selected from output",
    command=lambda: getSelectItemsFromTreeView(),
)
btn_select_data.place(relx=0.65, rely=0.7)
btn_cfm = tk.Button(
    root,
    text="Confirm",
    command=lambda: confirmAction(),
)
btn_cfm.pack(side=tk.BOTTOM, anchor="e", padx=15, pady=15)

# Labels
lbl_title = ttk.Label(
    root,
    text="This program will duplicate the target database and filter out Xinsha and other user defined tests",
    wraplength=460,
)
lbl_title.pack(side=tk.TOP, anchor="nw", padx=10, pady=10)
lbl_file = ttk.Label(frm_file, text="No File Selected", wraplength=450)
lbl_file.place(rely=0, relx=0)
lbl_selection_hint = ttk.Label(
    root,
    text="(Hold Ctrl/Shift and click to select multiple items)",
    font=("TkSmallCaptionFont", 8),
)
lbl_selection_hint.place(relx=0, rely=0.7)
v_selected_indexes = tk.StringVar()
v_selected_indexes.set("Selected Result Indexes:")
lbl_data_selected = ttk.Label(
    frm_selected_data, textvariable=v_selected_indexes, wraplength=450
)
lbl_data_selected.place(rely=0, relx=0)

# Treeview Widget
tv = ttk.Treeview(frm_dataview)
tv.place(relheight=1, relwidth=1)
treescrolly = tk.Scrollbar(frm_dataview, orient="vertical", command=tv.yview)
treescrollx = tk.Scrollbar(frm_dataview, orient="horizontal", command=tv.xview)
tv.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
treescrollx.pack(side="bottom", fill="x")
treescrolly.pack(side="right", fill="y")


def openFileDialog():
    """Open File dialog for user to choose database file"""
    global db_file_path
    global initial_dir
    db_file_path = filedialog.askopenfilename(
        initialdir=initial_dir,
        title="Select ICP Results Database File",
        filetype=(("Microsoft Access Database Files", "*.mdb"), ("All Files", "*.*")),
    )
    if db_file_path:
        logging.debug(f"File dialog selected file name: {db_file_path}")
        lbl_file["text"] = db_file_path
        # Write to config parser
        try:
            config["Settings"]["ICP_DB_FOLDER_PATH"] = os.path.dirname(db_file_path)
            with open("config.ini", "w") as f:
                config.write(f)
        except Exception as e:
            logging.debug(e)
        return db_file_path
    return


def loadDataToTreeView(db_file_path):
    """Load selected database into tkinter treeview using dataframes.
    script.py implements pandas and pyodbc to read database file.

    Args:
        db_file_path (_type_): _description_
    """
    if not db_file_path:
        return
    clearTreeView()
    try:
        logging.debug("Attempting to duplicate database...")
        progressbar = showProgressBar()
        root.update()
        global filtered_db_file_path
        filtered_db_file_path = script.duplicateDatabaseFile(db_file_path)
        logging.debug(f"Database duplicated. File path: {filtered_db_file_path}")
        logging.debug("Attempting to connect to database and query base view...")
        global conn
        conn = script.openDatabase(fp=filtered_db_file_path)
        df = script.getBaseViewTable(conn=conn)
        progressbar.destroy()
    except ValueError:
        logging.error("Invalid DB file chosen!")
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        logging.error("DB file does not exist!")
        tk.messagebox.showerror("Information", f"No such file as {db_file_path}")
        return None
    except Exception as e:
        logging.error(e)
        tk.messagebox.showerror(
            "Error",
            f"Database connection error or unexpected error. Please check if database is opened or locked for editing. Otherwise report the error to developer",
        )
    # Parsing dataframe to tkinter's treeview
    tv["column"] = list(df.columns)
    tv["show"] = "headings"
    for column in tv["columns"]:
        tv.heading(column, text=column)
    # Converts dataframe into a list of lists to populate treeview
    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        # inserts each list into the treeview
        tv.insert("", "end", values=row)
    return None


def confirmAction():
    """Handles user confirmation dialogue. If confirmed, database operations will proceed."""
    if db_file_path is None or conn is None:
        logging.debug(
            f"User confirmed without defining db file path ({db_file_path}) or a db connection ({conn}) does not exist"
        )
        message = messagebox.showerror(
            "Error!",
            "No database file or existing database connection. Plase browse and choose database file before proceeding.",
        )
        return
    message = messagebox.askquestion("Please confirm", "Confirm to proceed?")
    if message == "yes":
        logging.debug("User confirmed action")
        runDbOperation(selected_result_indexes=selected_result_indexes)
    else:
        logging.debug("User cancelled confirmation")
    return


def runDbOperation(selected_result_indexes=[]):
    """Handles the running of dropXinshaResultIndexes function in script.py by passing selected result index to function

    Args:
        selected_result_indexes (list, optional): _description_. Defaults to [].
    """
    try:
        logging.debug("Attempting to run database filter operations...")
        progressbar = showProgressBar()
        root.update()
        status = script.dropXinshaResultIndexes(
            conn=conn, additional_result_indexes=selected_result_indexes
        )
    except Exception as e:
        logging.error(e)
    progressbar.destroy()
    if status is True:
        messagebox.showinfo(
            "Success",
            f"Database filter and extraction operations successful! Filtered database in {filtered_db_file_path}",
        )
        # Open file explorer on output database directory for user
        command = (
            rf'explorer /select,"{os.path.dirname(filtered_db_file_path)}"'.replace(
                "/", "\\"
            )
        )
        subprocess.Popen(command)
        root.destroy()


def clearTreeView():
    """Clear tkinter's tree view"""
    tv.delete(*tv.get_children())
    return None


def getSelectItemsFromTreeView():
    """Get selected items from tree view and assign the selected result indexes as global var"""
    selected_items = tv.selection()
    selected_result_indexes = [tv.item(i)["values"][0] for i in selected_items]
    logging.debug(f"User selected result indexes: {selected_result_indexes}")
    # print("\n".join([str(tv.item(i)["values"]) for i in selected_items]))
    v_selected_indexes.set(f"Selected Result Indexes: {str(selected_result_indexes)}")


def showProgressBar():
    """Show tkinter progress bar. This will be used for long running operations"""
    progressbar = ttk.Progressbar(mode="indeterminate")
    progressbar.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
    progressbar.start()
    return progressbar


root.protocol("WM_DELETE_WINDOW", logging.debug("App Closed"))
logging.debug("App Initialized")
root.mainloop()
