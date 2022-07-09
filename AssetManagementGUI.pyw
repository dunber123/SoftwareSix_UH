from tkinter import messagebox
import pyodbc
import tkinter as tk
from tkinter import ttk

global cnxn
cnxn = """Driver={SQL Server Native Client 11.0};
                Server=CoT-CIS3365-10.cougarnet.uh.edu;
                Port=1433;
                Database=SoftwareSix;
                UID=username;
                PWD=password;
                Trusted_Connection=no;"""

# ----------------------------------------
#          Begin Login
# ----------------------------------------
def SignIn():
    global root
    global Username
    global password
    # Initialize Login Window
    root = tk.Tk()
    root.title("aeSolutions Asset Management Database")
    root.configure(bg='gray15')

    # Define login window dimensions and place on center of screen
    window_width = 300
    window_height = 200
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    # Creating buttons and labels
    Lbl_Username = tk.Label(root, text="Username -")
    Lbl_Username.place(x=50, y=20, width=90)
    Username = tk.Entry(root, width=35)
    Username.place(x=150, y=20, width=100)
    Lbl_Password = tk.Label(root, text="Password -")
    Lbl_Password.place(x=50, y=50, width=90)
    password = tk.Entry(root, show="*", width=35)
    password.place(x=150, y=50, width=100)
    Btn_Login = tk.Button(root, text="Login", bg='DodgerBlue3', command=testCredentials)
    Btn_Login.place(x=150, y=90, width=100)
    root.mainloop()

# Test user credentials during log in: If unsuccessful connection, Error is displayed.
def testCredentials():
    global user
    global passw

    for i in range(10):
        user = Username.get()
        passw = password.get()
        try:
            conn = pyodbc.connect(
                "Driver={SQL Server Native Client 11.0};"
                "Server=CoT-CIS3365-10.cougarnet.uh.edu;"
                "Port=1433;"
                "Database=SoftwareSix;"
                "UID=" + str(user) + ";"
                "PWD=" + str(passw) + ";"
                "Trusted_Connection=no;"
            )
            root.withdraw()
            homePage()
        except pyodbc.Error as ex:
            messagebox.showinfo("Error", "Invalid Credentials")

        break

# ----------------------------------------
#          Begin Main Navigation
# ----------------------------------------

# Bring user to application home page after successful login
def homePage():
    global homeWindow
    global client_ID

    def logOut():
        root.quit()

    def updateClient_ID(event):
        global client_ID
        client_ID = ''
        clientName = Cmb_Clients.get()
        for client in clients:
            if clientName == client[1]:
                client_ID = client[0]
        return

    def validateClient_ID(useCase):
        global client_ID
        if client_ID != '':
            if useCase == "Client Assets":
                showAssetTree(client_ID)
            elif useCase == "Client Information":
                manageClientInfo(client_ID)
            elif useCase == "Reports":
                reports(client_ID)
            elif useCase == "OtherTables":
                otherTables()
            elif useCase == "New Client":
                newClient()
        else:
            messagebox.showinfo("Error", "Please Select a Client")

    root.withdraw()

    # Create second window and its dimensions
    homeWindow = tk.Toplevel(root)
    homeWindow.title("Home Page")
    homeWindow.configure(bg='gray15')
    window_width = 440
    window_height = 250
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    homeWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    # Create buttons
    btn_ManageAssets = tk.Button(homeWindow, width=5, text="Manage Client Assets", bg='white',
                                 command=lambda: validateClient_ID("Client Assets"))
    btn_ManageClients = tk.Button(homeWindow, width=5, text="Manage Client Information", bg='white',
                                  command=lambda: validateClient_ID("Client Information"))
    btn_OtherTables = tk.Button(homeWindow, width=5, text="Manage Other Tables", bg='white',
                                command=lambda: validateClient_ID("OtherTables"))
    btn_LogOut = tk.Button(homeWindow, width=5, text="Log Out", bg='firebrick1', command=logOut)
    Cmb_Clients = ttk.Combobox(homeWindow, width=30)

    btn_ManageAssets.grid(row=1, column=1, columnspan=1, pady=10, padx=5, ipadx=75, ipady=10)
    btn_ManageClients.grid(row=1, column=2, columnspan=1, pady=10, padx=5, ipadx=75, ipady=10)
    btn_OtherTables.grid(row=2, column=1, columnspan=3, pady=10, padx=5, ipadx=75, ipady=10)
    btn_LogOut.grid(row=4, column=1, columnspan=2, pady=10, padx=5, ipadx=75)
    Cmb_Clients.grid(row=0, column=1, columnspan=3)
    Cmb_Clients.bind("<<ComboboxSelected>>", updateClient_ID)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    cursor.execute("""SELECT p.Client_ID, tbl_Companies.Name
                            FROM tbl_Client_Info As p
                            INNER JOIN tbl_Contacts ON p.Contact_ID = tbl_Contacts.Contact_ID
                            INNER JOIN tbl_Companies ON tbl_Contacts.Company_ID = tbl_Companies.Company_ID
                            WHERE p.Deleted IS NULL""")
    clients = cursor.fetchall()

    clientList = []
    for client in clients:
        clientList.append(client[1])
    Cmb_Clients['values'] = clientList
    Cmb_Clients.current(0)
    client_ID = clients[0][0]


# This function is responsible for closing/opening windows when using back/home buttons
def home(windowToClose, windowToOpen):
    try:
        if windowToClose == linkAssetsWindow and windowToOpen == AssetTreeViewWindow:
            windowToOpen.destroy()
            showAssetTree(client_ID)
        elif windowToClose == siwAssetsWindow and windowToOpen == AssetTreeViewWindow:
            windowToOpen.destroy()
            showAssetTree(client_ID)
    except:
        windowToClose.destroy()
        windowToOpen.deiconify()


def manageClientInfo(client_ID):
    global ClientInfoWindow

    def updateListboxSel(event):
        global Contact_ID
        global Client_ID

        currSelection = Lbl_AllCompanies.curselection()

        if currSelection != ():
            Company = Lbl_AllCompanies.get(currSelection)
            indexNum = currSelection[0] + 1
            for record in Companies:
                if Company == record[0]:
                    Company_ID = record[1]
                    Client_ID = record[2]
                    Contact_ID = record[3]
            populateGeneralInfo(Client_ID, Contact_ID)

    def populateGeneralInfo(client_ID, Contact_ID):
        global industries
        global countries
        global states

        cursor.execute("""SELECT tbl_Industries.IndustryName,
                                tbl_Countries.CountryName,
                                tbl_States.StateName,
                                tbl_Contacts.ContactName,
                                tbl_Contacts.ContactEmail,
                                tbl_Contacts.ContactPhone,
                                tbl_Contacts.ContactRelation,
                                tbl_Contacts.Contact_ID

                                FROM tbl_Client_Info as p
                                INNER JOIN tbl_Industries ON p.Industry_ID = tbl_Industries.Industry_ID
                                INNER JOIN tbl_Countries ON p.Country_ID = tbl_Countries.Country_ID
                                INNER JOIN tbl_States ON p.State_ID = tbl_States.State_ID
                                INNER JOIN tbl_Contacts ON p.Contact_ID = tbl_Contacts.Contact_ID
                                WHERE p.Client_ID = """ + str(client_ID))
        records = cursor.fetchall()
        for record in records:
            Cmb_Industry.set(record[0])
            Cmb_Country.set(record[1])
            Cmb_State.set(record[2])
            ContactName_Var.set(record[3])
            email_Var.set(record[4])
            Number_var.set(record[5])
            Relation_var.set(record[6])
            Contact_ID = record[7]

        cursor.execute("""SELECT tbl_Industries.Industry_ID, tbl_Industries.IndustryName
                                    FROM tbl_Industries""")
        industries = cursor.fetchall()
        IndustriesList = []
        for record in industries:
            IndustriesList.append(record[1])
        Cmb_Industry["values"] = IndustriesList

        cursor.execute("""SELECT tbl_Countries.Country_ID, tbl_Countries.CountryName
                            FROM tbl_Countries""")
        countries = cursor.fetchall()
        countryList = []
        for record in countries:
            countryList.append(record[1])
        Cmb_Country["values"] = countryList

        cursor.execute("""SELECT tbl_States.State_ID, tbl_States.StateName
                                FROM tbl_States""")
        states = cursor.fetchall()
        StatesList = []
        for record in states:
            StatesList.append(record[1])
        Cmb_State["values"] = StatesList
        return

    def edit():
        global btn_Save
        global btn_Cancel
        btn_Save = tk.Button(Frame1, text="Save", bg='white', command=save)
        btn_Cancel = tk.Button(Frame1, text="Cancel", bg='white', command=cancel)
        btn_Save.place(x=150, y=400, width=120)
        btn_Cancel.place(x=300, y=400, width=120)
        Cmb_Industry['state'] = "normal"
        Cmb_Country['state'] = "normal"
        Cmb_State['state'] = "normal"
        Entry_Email['state'] = "normal"
        Entry_ContactName['state'] = "normal"
        Entry_Relation['state'] = "normal"
        Entry_PhoneNumber['state'] = "normal"

        return

    def delete():
        SQL = """UPDATE tbl_Client_Info
                    SET Deleted = 1
                    WHERE tbl_Client_Info.Client_ID = {0}""".format(Client_ID)
        cursor.execute(SQL)
        cursor.commit()
        SQL = """UPDATE tbl_Contacts
                            SET Deleted = 1
                            WHERE tbl_Contacts.Contact_ID = {0}""".format(Contact_ID)
        cursor.execute(SQL)
        cursor.commit()
        homeWindow.destroy()
        homePage()
        manageClientInfo(client_ID)
        return

    def cancel():
        Cmb_Industry['state'] = "disabled"
        Cmb_Country['state'] = "disabled"
        Cmb_State['state'] = "disabled"
        Entry_Email['state'] = "disabled"
        Entry_ContactName['state'] = "disabled"
        Entry_Relation['state'] = "disabled"
        Entry_PhoneNumber['state'] = "disabled"
        btn_Save.place_forget()
        btn_Cancel.place_forget()
        return

    def save():
        State_ID = 'NULL'
        Country_ID = 'NULL'
        Industry_ID = 'NULL'
        for record in states:
            if Cmb_State.get() == record[1]:
                State_ID = record[0]
        for record in industries:
            if Cmb_Industry.get() == record[1]:
                Industry_ID = record[0]
        for record in countries:
            if Cmb_Country.get() == record[1]:
                Country_ID = record[0]

        SQL = """UPDATE tbl_Client_Info
                            SET Industry_ID = {0}, Country_ID = {1}, State_ID = {2}
                            WHERE Client_ID = {3}""".format(Industry_ID, Country_ID, State_ID, client_ID)
        cursor.execute(SQL)
        cursor.commit()
        SQL = """Update tbl_Contacts
                    SET ContactName = '{0}', ContactEmail = '{1}', ContactPhone= '{2}', ContactRelation = '{3}'
                    WHERE Contact_ID = {4}""".format(Entry_ContactName.get(), Entry_Email.get(),
                                                     Entry_PhoneNumber.get(), Entry_Relation.get(), Contact_ID)
        cursor.execute(SQL)
        cursor.commit()
        ClientInfoWindow.destroy()
        manageClientInfo(client_ID)
        return

    # make window
    homeWindow.withdraw()
    ClientInfoWindow = tk.Toplevel(homeWindow)
    ClientInfoWindow.geometry("600x300")
    ClientInfoWindow.title("Client Information")
    ClientInfoWindow.configure(bg='gray15')
    window_width = 815
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    ClientInfoWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    Frame1 = tk.Frame(ClientInfoWindow, padx=5, pady=5, bg='white')
    Frame1.place(x=160, y=5, width=500, height=500)

    Lbl_AllCompanies = tk.Listbox(ClientInfoWindow)
    Lbl_AllCompanies.place(x=5, y=5, width=150)
    Lbl_AllCompanies.bind('<<ListboxSelect>>', updateListboxSel)

    label_GeneralInfo = tk.Label(Frame1, text="General Info:", bg='white')
    label_GeneralInfo.config(font=('TkDefaultFont', 15, 'bold'))
    label_GeneralInfo.place(x=185, y=5, width=120)

    label_Industry = tk.Label(Frame1, text="Industry", bg='white', anchor='w')
    label_Industry.config(font=('TkDefaultFont', 10, 'bold'))
    label_Industry.place(x=5, y=50, width=120)
    Cmb_Industry = ttk.Combobox(Frame1)
    Cmb_Industry['state'] = "disabled"
    Cmb_Industry.place(x=130, y=50, width=200)

    label_Country = tk.Label(Frame1, text="Country", bg='white', anchor='w')
    label_Country.config(font=('TkDefaultFont', 10, 'bold'))
    label_Country.place(x=5, y=100, width=120)
    Cmb_Country = ttk.Combobox(Frame1)
    Cmb_Country['state'] = "disabled"
    Cmb_Country.place(x=130, y=100, width=200)

    label_State = tk.Label(Frame1, text="State", bg='white', anchor='w')
    label_State.config(font=('TkDefaultFont', 10, 'bold'))
    label_State.place(x=5, y=150, width=120)
    Cmb_State = ttk.Combobox(Frame1)
    Cmb_State['state'] = "disabled"
    Cmb_State.place(x=130, y=150, width=200)

    label_ContactName = tk.Label(Frame1, text="Contact Name", bg='white', anchor='w')
    label_ContactName.config(font=('TkDefaultFont', 10, 'bold'))
    label_ContactName.place(x=5, y=200, width=120)
    ContactName_Var = tk.StringVar()
    Entry_ContactName = tk.Entry(Frame1, textvariable=ContactName_Var)
    Entry_ContactName['state'] = "disabled"
    Entry_ContactName.place(x=130, y=200, width=200)

    label_Email = tk.Label(Frame1, text="Email", bg='white', anchor='w')
    label_Email.config(font=('TkDefaultFont', 10, 'bold'))
    label_Email.place(x=5, y=250, width=120)
    email_Var = tk.StringVar()
    Entry_Email = tk.Entry(Frame1, textvariable=email_Var)
    Entry_Email['state'] = "disabled"
    Entry_Email.place(x=130, y=250, width=200)

    label_PhoneNumber = tk.Label(Frame1, text="Phone Number", bg='white', anchor='w')
    label_PhoneNumber.config(font=('TkDefaultFont', 10, 'bold'))
    label_PhoneNumber.place(x=5, y=300, width=120)
    Number_var = tk.StringVar()
    Entry_PhoneNumber = tk.Entry(Frame1, textvariable=Number_var)
    Entry_PhoneNumber['state'] = "disabled"
    Entry_PhoneNumber.place(x=130, y=300, width=200)

    label_Relation = tk.Label(Frame1, text="Relation", bg='white', anchor='w')
    label_Relation.config(font=('TkDefaultFont', 10, 'bold'))
    label_Relation.place(x=5, y=350, width=120)
    Relation_var = tk.StringVar()
    Entry_Relation = tk.Entry(Frame1, textvariable=Relation_var)
    Entry_Relation['state'] = "disabled"
    Entry_Relation.place(x=130, y=350, width=200)

    btn_Edit = tk.Button(Frame1, text="Edit", bg='white', command=edit)
    btn_Edit.place(x=5, y=400, width=120)

    # Create and place buttons on right side of window + Back and Home buttons
    btn_ManageCountries = tk.Button(ClientInfoWindow, width=5, text="Manage Countries", bg='white',
                                    command=ManageCountries)
    btn_ManageStates = tk.Button(ClientInfoWindow, width=5, text="Manage States", bg='white', command=ManageStates)
    btn_ManageIndustries = tk.Button(ClientInfoWindow, width=5, text="Manage Industries", bg='white',
                                     command=manageIndustries)
    btn_NewClient = tk.Button(ClientInfoWindow, width=5, text="Add New Client", bg='white', command=newClient)
    btn_Deleted = tk.Button(Frame1, width=5, text="Delete Client", bg='white', command=delete)
    btn_home = tk.Button(ClientInfoWindow, width=5, text="Home", bg='white',
                         command=lambda: home(ClientInfoWindow, homeWindow))
    btn_Deleted.place(x=5, y=450, width=120)
    btn_home.place(x=665, y=5, width=150)
    btn_ManageCountries.place(x=665, y=55, width=150)
    btn_ManageStates.place(x=665, y=105, width=150)
    btn_ManageIndustries.place(x=665, y=155, width=150)
    btn_NewClient.place(x=665, y=205, width=150)

    # create connection to database and execute SQL code
    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()

    cursor.execute("""SELECT tbl_Companies.Name, tbl_Companies.Company_ID, tbl_Client_Info.Client_ID, tbl_Contacts.Contact_ID
                FROM tbl_Companies
                INNER JOIN tbl_Contacts ON tbl_Companies.Company_ID = tbl_Contacts.Company_ID
                INNER JOIN tbl_Client_Info ON tbl_Contacts.Contact_ID = tbl_Client_Info.Contact_ID
                WHERE tbl_Client_Info.Deleted IS NULL""")
    Companies = cursor.fetchall()
    companyNames = []
    for record in Companies:
        Lbl_AllCompanies.insert('end', record[0])
    Lbl_AllCompanies.select_set(0)

    return


def ManageCountries():
    def updateWidgets(event):
        try:
            if Entry_Country.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                CountryInfo = Selection.get('values')
                Country_Var.set(CountryInfo[1])
                Code_Var.set(CountryInfo[2])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:

            global Entry_CountryCode
            global Entry_Country

            global Country_Var
            global Code_Var

            global Btn_Cancel
            global Btn_Save

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                CountryInfo = Selection.get('values')

                CountryName = 'NULL'
                CountryCode = 'NULL'
                CountryName = Entry_Country.get()
                CountryCode = Entry_CountryCode.get()

                SQL = """UPDATE tbl_Countries
                        SET CountryName = '{0}', Abbreviation = '{1}'
                        WHERE Country_ID = {2}""".format(CountryName, CountryCode, CountryInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                CountryWindow.destroy()
                ManageCountries()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            CategoryInfo = Selection.get('values')

            Country_Var = tk.StringVar()
            Code_Var = tk.StringVar()

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(CountryWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Country = tk.Label(Frame1, text="Country Name", bg='white', anchor='w')
            lbl_CountryCode = tk.Label(Frame1, text="Country Code", bg='white', anchor='w')
            Entry_Country = tk.Entry(Frame1, textvariable=Country_Var, width=50)
            Entry_CountryCode = tk.Entry(Frame1, textvariable=Code_Var, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Country.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_CountryCode.grid(row=1, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Country.grid(row=0, column=1, pady=5, padx=5, sticky='w')
            Entry_CountryCode.grid(row=1, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            Country_Var.set(CategoryInfo[1])
            Code_Var.set(CategoryInfo[2])

            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        CountryInfo = Selection.get('values')
        SQL = """UPDATE tbl_Countries
                    SET Deleted = 1
                    WHERE tbl_Countries.Country_ID = {0}""".format(CountryInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global Entry_Country

            global Btn_Cancel
            global Btn_Save

            def save():
                CountryName = 'NULL'
                CountryCode = 'NULL'
                CountryName = Entry_Country.get()
                CountryCode = Entry_CountryCode.get()

                SQL = """INSERT INTO tbl_Countries (CountryName, Abbreviation)
                                VALUES ('{0}', '{1}')""".format(CountryName, CountryCode)
                cursor.execute(SQL)
                cursor.commit()
                CountryWindow.destroy()
                ManageCountries()
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(CountryWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Country = tk.Label(Frame1, text="Country Name", bg='white', anchor='w')
            Entry_Country = tk.Entry(Frame1, width=50)
            lbl_CountryCode = tk.Label(Frame1, text="Country Code", bg='white', anchor='w')
            Entry_CountryCode = tk.Entry(Frame1, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Country.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_CountryCode.grid(row=1, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Country.grid(row=0, column=1, pady=5, padx=5, sticky='w')
            Entry_CountryCode.grid(row=1, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            return

    CountryWindow = tk.Toplevel(ClientInfoWindow)
    ClientInfoWindow.withdraw()
    CountryWindow.geometry("600x300")
    CountryWindow.title("Countries")
    CountryWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    CountryWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(CountryWindow, width=15, text="Home", bg='white',
                         command=lambda: home(CountryWindow, homeWindow))
    btn_back = tk.Button(CountryWindow, width=15, text="Back", bg='white',
                         command=lambda: home(CountryWindow, ClientInfoWindow))
    btn_edit = tk.Button(CountryWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(CountryWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry())
    btn_delete = tk.Button(CountryWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(CountryWindow, column=('#1', '#2', '#3'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Country Name')
    tree.heading('#3', text='Country Code')
    tree.column('#1', width=40)
    tree.column('#2', width=500)
    tree.column('#3', width=350)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_Countries.Country_ID,
                tbl_Countries.CountryName,
                tbl_Countries.Abbreviation
                FROM tbl_Countries
                WHERE tbl_Countries.Deleted IS NULL"""
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1], record[2]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def ManageStates():
    def updateWidgets(event):
        try:
            if Entry_Country.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                CountryInfo = Selection.get('values')
                Country_Var.set(CountryInfo[1])
                Code_Var.set(CountryInfo[2])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:

            global Entry_CountryCode
            global Entry_Country

            global Country_Var
            global Code_Var

            global Btn_Cancel
            global Btn_Save

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                CountryInfo = Selection.get('values')

                CountryName = 'NULL'
                CountryCode = 'NULL'
                CountryName = Entry_Country.get()
                CountryCode = Entry_CountryCode.get()

                SQL = """UPDATE tbl_States
                        SET StateName = '{0}', Abbreviation = '{1}'
                        WHERE State_ID = {2}""".format(CountryName, CountryCode, CountryInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                StatesWindow.destroy()
                ManageStates()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            CategoryInfo = Selection.get('values')

            Country_Var = tk.StringVar()
            Code_Var = tk.StringVar()

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(StatesWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Country = tk.Label(Frame1, text="State Name", bg='white', anchor='w')
            lbl_CountryCode = tk.Label(Frame1, text="State Code", bg='white', anchor='w')
            Entry_Country = tk.Entry(Frame1, textvariable=Country_Var, width=50)
            Entry_CountryCode = tk.Entry(Frame1, textvariable=Code_Var, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Country.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_CountryCode.grid(row=1, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Country.grid(row=0, column=1, pady=5, padx=5, sticky='w')
            Entry_CountryCode.grid(row=1, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            Country_Var.set(CategoryInfo[1])
            Code_Var.set(CategoryInfo[2])

            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        CountryInfo = Selection.get('values')
        SQL = """UPDATE tbl_States
                    SET Deleted = 1
                    WHERE tbl_States.State_ID = {0}""".format(CountryInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global Entry_Country

            global Btn_Cancel
            global Btn_Save

            def save():
                CountryName = 'NULL'
                CountryCode = 'NULL'
                CountryName = Entry_Country.get()
                CountryCode = Entry_CountryCode.get()

                SQL = """INSERT INTO tbl_States (StateName, Abbreviation)
                                VALUES ('{0}', '{1}')""".format(CountryName, CountryCode)
                cursor.execute(SQL)
                cursor.commit()
                StatesWindow.destroy()
                ManageStates()
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(StatesWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Country = tk.Label(Frame1, text="State Name", bg='white', anchor='w')
            Entry_Country = tk.Entry(Frame1, width=50)
            lbl_CountryCode = tk.Label(Frame1, text="State Code", bg='white', anchor='w')
            Entry_CountryCode = tk.Entry(Frame1, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Country.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_CountryCode.grid(row=1, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Country.grid(row=0, column=1, pady=5, padx=5, sticky='w')
            Entry_CountryCode.grid(row=1, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            return

    StatesWindow = tk.Toplevel(ClientInfoWindow)
    ClientInfoWindow.withdraw()
    StatesWindow.geometry("600x300")
    StatesWindow.title("States")
    StatesWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    StatesWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(StatesWindow, width=15, text="Home", bg='white',
                         command=lambda: home(StatesWindow, homeWindow))
    btn_back = tk.Button(StatesWindow, width=15, text="Back", bg='white',
                         command=lambda: home(StatesWindow, ClientInfoWindow))
    btn_edit = tk.Button(StatesWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(StatesWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry())
    btn_delete = tk.Button(StatesWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(StatesWindow, column=('#1', '#2', '#3'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='State Name')
    tree.heading('#3', text='State Code')
    tree.column('#1', width=40)
    tree.column('#2', width=500)
    tree.column('#3', width=350)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_States.State_ID,
                tbl_States.StateName,
                tbl_States.Abbreviation
                FROM tbl_States
                WHERE tbl_States.Deleted IS NULL"""
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1], record[2]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def manageIndustries():
    def updateWidgets(event):
        try:
            if Entry_Category.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                CategoryInfo = Selection.get('values')
                Category_Var.set(CategoryInfo[1])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:

            global lbl_Category
            global Entry_Category

            global Category_Var

            global Btn_Cancel
            global Btn_Save

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                USBInfo = Selection.get('values')

                CategoryName = 'NULL'
                CategoryName = Entry_Category.get()

                SQL = """UPDATE tbl_Industries
                        SET IndustryName = '{0}'
                        WHERE Industry_ID = {1}""".format(CategoryName, USBInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                IndustriesWindow.destroy()
                manageIndustries()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            CategoryInfo = Selection.get('values')

            Category_Var = tk.StringVar()

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(IndustriesWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Category = tk.Label(Frame1, text="Industry Name", bg='white', anchor='w')
            Entry_Category = tk.Entry(Frame1, textvariable=Category_Var, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Category.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Category.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            Category_Var.set(CategoryInfo[1])

            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        CategoryInfo = Selection.get('values')
        SQL = """UPDATE tbl_Industries
                    SET Deleted = 1
                    WHERE tbl_Industries.Industry_ID = {0}""".format(CategoryInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global lbl_Category
            global Entry_Category

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save():
                categoryName = 'NULL'
                categoryName = Entry_Category.get()

                SQL = """INSERT INTO tbl_Industries (IndustryName)
                                VALUES ('{0}')""".format(categoryName)
                cursor.execute(SQL)
                cursor.commit()
                IndustriesWindow.destroy()
                manageIndustries()
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(IndustriesWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Category = tk.Label(Frame1, text="Category Name", bg='white', anchor='w')
            Entry_Category = tk.Entry(Frame1, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Category.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Category.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            return

    IndustriesWindow = tk.Toplevel(ClientInfoWindow)
    ClientInfoWindow.withdraw()
    IndustriesWindow.geometry("600x300")
    IndustriesWindow.title("Industries")
    IndustriesWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    IndustriesWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(IndustriesWindow, width=15, text="Home", bg='white',
                         command=lambda: home(IndustriesWindow, homeWindow))
    btn_back = tk.Button(IndustriesWindow, width=15, text="Back", bg='white',
                         command=lambda: home(IndustriesWindow, ClientInfoWindow))
    btn_edit = tk.Button(IndustriesWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(IndustriesWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry())
    btn_delete = tk.Button(IndustriesWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(IndustriesWindow, column=('#1', '#2'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Industry Name')
    tree.column('#1', width=40)
    tree.column('#2', width=850)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_Industries.Industry_ID,
                tbl_Industries.IndustryName
                FROM tbl_Industries
                WHERE tbl_Industries.Deleted IS NULL"""
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def newClient():
    global ClientInfoWindow

    def cancel():
        Cmb_Industry['state'] = "disabled"
        Cmb_Country['state'] = "disabled"
        Cmb_State['state'] = "disabled"
        Entry_Email['state'] = "disabled"
        Entry_ContactName['state'] = "disabled"
        Entry_Relation['state'] = "disabled"
        Entry_PhoneNumber['state'] = "disabled"
        btn_Save.place_forget()
        btn_Cancel.place_forget()
        return

    def save():
        State_ID = 'NULL'
        Country_ID = 'NULL'
        Industry_ID = 'NULL'
        for record in states:
            if Cmb_State.get() == record[1]:
                State_ID = record[0]
        for record in industries:
            if Cmb_Industry.get() == record[1]:
                Industry_ID = record[0]
        for record in countries:
            if Cmb_Country.get() == record[1]:
                Country_ID = record[0]
        SQL = """INSERT INTO tbl_Companies (Name)
                            Values ('{0}')""".format(Entry_Company.get())
        cursor.execute(SQL)
        cursor.commit()
        SQL = """SELECT tbl_Companies.Company_ID
                    FROM tbl_Companies
                    WHERE tbl_Companies.Name = '{0}'""".format(Entry_Company.get())
        cursor.execute(SQL)
        records = cursor.fetchall()
        for record in records:
            Company_ID = record[0]
        SQL = """INSERT INTO tbl_Contacts (Company_ID, ContactName, ContactEmail, ContactPhone, ContactRelation)
                            Values ({0}, '{1}', '{2}', '{3}', '{4}')""".format(Company_ID, Entry_ContactName.get(),
                                                                               Entry_Email.get(),
                                                                               Entry_PhoneNumber.get(),
                                                                               Entry_Relation.get())
        cursor.execute(SQL)
        cursor.commit()

        SQL = """SELECT tbl_Contacts.Contact_ID
                    FROM tbl_Contacts
                    WHERE tbl_Contacts.Company_ID = {0}""".format(Company_ID)
        cursor.execute(SQL)
        records = cursor.fetchall()
        for record in records:
            Contact_ID = record[0]
        SQL = """INSERT INTO tbl_Client_Info (Industry_ID, Country_ID, State_ID, Contact_ID)
                                   Values ({0}, {1}, {2}, {3})""".format(Industry_ID,
                                                                         Country_ID,
                                                                         State_ID,
                                                                         Contact_ID)
        cursor.execute(SQL)
        cursor.commit()
        homeWindow.destroy()
        homePage()
        manageClientInfo(client_ID)
        return

    # make window
    ClientInfoWindow.withdraw()
    newClientWindow = tk.Toplevel(homeWindow)
    newClientWindow.geometry("600x300")
    newClientWindow.title("Client Information")
    newClientWindow.configure(bg='gray15')
    window_width = 815
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    newClientWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    Frame1 = tk.Frame(newClientWindow, padx=5, pady=5, bg='white')
    Frame1.place(x=160, y=5, width=500, height=500)

    label_GeneralInfo = tk.Label(Frame1, text="General Info:", bg='white')
    label_GeneralInfo.config(font=('TkDefaultFont', 15, 'bold'))
    label_GeneralInfo.place(x=185, y=5, width=120)

    label_Company = tk.Label(Frame1, text="Company Name", bg='white', anchor='w')
    label_Company.config(font=('TkDefaultFont', 10, 'bold'))
    label_Company.place(x=5, y=50, width=120)
    Entry_Company = tk.Entry(Frame1)
    Entry_Company.place(x=130, y=50, width=200)

    label_Industry = tk.Label(Frame1, text="Industry", bg='white', anchor='w')
    label_Industry.config(font=('TkDefaultFont', 10, 'bold'))
    label_Industry.place(x=5, y=100, width=120)
    Cmb_Industry = ttk.Combobox(Frame1)
    Cmb_Industry.place(x=130, y=100, width=200)

    label_Country = tk.Label(Frame1, text="Country", bg='white', anchor='w')
    label_Country.config(font=('TkDefaultFont', 10, 'bold'))
    label_Country.place(x=5, y=150, width=120)
    Cmb_Country = ttk.Combobox(Frame1)
    Cmb_Country.place(x=130, y=150, width=200)

    label_State = tk.Label(Frame1, text="State", bg='white', anchor='w')
    label_State.config(font=('TkDefaultFont', 10, 'bold'))
    label_State.place(x=5, y=200, width=120)
    Cmb_State = ttk.Combobox(Frame1)
    Cmb_State.place(x=130, y=200, width=200)

    label_ContactName = tk.Label(Frame1, text="Contact Name", bg='white', anchor='w')
    label_ContactName.config(font=('TkDefaultFont', 10, 'bold'))
    label_ContactName.place(x=5, y=250, width=120)
    ContactName_Var = tk.StringVar()
    Entry_ContactName = tk.Entry(Frame1, textvariable=ContactName_Var)
    Entry_ContactName.place(x=130, y=250, width=200)

    label_Email = tk.Label(Frame1, text="Email", bg='white', anchor='w')
    label_Email.config(font=('TkDefaultFont', 10, 'bold'))
    label_Email.place(x=5, y=300, width=120)
    email_Var = tk.StringVar()
    Entry_Email = tk.Entry(Frame1, textvariable=email_Var)
    Entry_Email.place(x=130, y=300, width=200)

    label_PhoneNumber = tk.Label(Frame1, text="Phone Number", bg='white', anchor='w')
    label_PhoneNumber.config(font=('TkDefaultFont', 10, 'bold'))
    label_PhoneNumber.place(x=5, y=350, width=120)
    Number_var = tk.StringVar()
    Entry_PhoneNumber = tk.Entry(Frame1, textvariable=Number_var)
    Entry_PhoneNumber.place(x=130, y=350, width=200)

    label_Relation = tk.Label(Frame1, text="Relation", bg='white', anchor='w')
    label_Relation.config(font=('TkDefaultFont', 10, 'bold'))
    label_Relation.place(x=5, y=400, width=120)
    Relation_var = tk.StringVar()
    Entry_Relation = tk.Entry(Frame1, textvariable=Relation_var)
    Entry_Relation.place(x=130, y=400, width=200)

    # Create and place buttons on right side of window + Back and Home buttons
    btn_home = tk.Button(newClientWindow, width=5, text="Home", bg='white',
                         command=lambda: home(newClientWindow, homeWindow))
    btn_home.place(x=665, y=5, width=150)
    btn_Save = tk.Button(Frame1, text="Save", bg='white', command=save)
    btn_Save.place(x=150, y=450, width=120)

    # create connection to database and execute SQL code
    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()

    cursor.execute("""SELECT tbl_Industries.Industry_ID, tbl_Industries.IndustryName
                                        FROM tbl_Industries""")
    industries = cursor.fetchall()
    IndustriesList = []
    for record in industries:
        IndustriesList.append(record[1])
    Cmb_Industry["values"] = IndustriesList

    cursor.execute("""SELECT tbl_Countries.Country_ID, tbl_Countries.CountryName
                                FROM tbl_Countries""")
    countries = cursor.fetchall()
    countryList = []
    for record in countries:
        countryList.append(record[1])
    Cmb_Country["values"] = countryList

    cursor.execute("""SELECT tbl_States.State_ID, tbl_States.StateName
                                    FROM tbl_States""")
    states = cursor.fetchall()
    StatesList = []
    for record in states:
        StatesList.append(record[1])
    Cmb_State["values"] = StatesList
    return


# This page shows the asset tree and allows the user to add/ edit / delete client asset information
def showAssetTree(client_ID):
    global AssetTreeViewWindow
    global tree_ID
    global clientSelected
    global buildingSelected
    global SIWAssetSelected
    global SwitchAssetSelected
    global assetSelected
    buildingSelected = False
    clientSelected = False
    SIWAssetSelected = False
    SwitchAssetSelected = False
    assetSelected = False

    def unlinkAsset(asset_ID, assetSelected, tree_ID):
        if assetSelected == "SIW Asset":
            tree.delete(tree_ID)
            sqlScript = "UPDATE tbl_SIWAssets SET Building_ID = NULL WHERE SIWAsset_ID = " + str(
                asset_ID)
        elif assetSelected == "Switch Asset":
            tree.delete(tree_ID)
            sqlScript = "UPDATE tbl_SwitchAssets SET Building_ID = NULL WHERE SwitchAsset_ID = " + str(
                asset_ID)
        conn = pyodbc.connect(cnxn)
        cursor = conn.cursor()
        cursor.execute(sqlScript)
        conn.commit()
        return

    def checkVar(functionToCall, assetVar):
        global tree_ID
        global assetTypeSelected
        global asset_ID
        if functionToCall == "LinkAsset":
            if assetVar == True:
                linkAsset(asset_ID)
            else:
                messagebox.showinfo("Error", "Please select a building first")
        elif functionToCall == "UnlinkAsset":
            if assetVar == True:
                unlinkAsset(asset_ID, assetTypeSelected, tree_ID)
            else:
                messagebox.showinfo("Error", "Please select a Switch or SIW Asset first")

    # This function updates the asset_ID based on the users selection in the treeview. The boolean variables that are
    # assigned in this function will also help with error handling in the future (ex: LinkAsset function will not
    # execute unless buildingSelected = True
    def updateAsset_ID(event):
        global tree_ID
        global clientSelected
        global buildingSelected
        global SIWAssetSelected
        global SwitchAssetSelected
        global assetTypeSelected
        global assetSelected
        global asset_ID
        buildingSelected = False
        clientSelected = False
        SIWAssetSelected = False
        SwitchAssetSelected = False
        assetSelected = False
        assetTypeSelected = ""
        tree_ID = int(tree.focus())

        for record in clientList:
            if record[2] == tree_ID:
                asset_ID = record[0]
                clientSelected = True
        if clientSelected == False:
            for record in buildingsList:
                if record[3] == tree_ID:
                    asset_ID = record[0]
                    buildingSelected = True
        if buildingSelected == False:
            for record in SIWAssetList:
                if record[3] == tree_ID:
                    asset_ID = record[0]
                    SIWAssetSelected = True
                    assetSelected = True
                    assetTypeSelected = "SIW Asset"
        if SIWAssetSelected == False:
            for record in SwitchList:
                if record[3] == tree_ID:
                    asset_ID = record[0]
                    SwitchAssetSelected = True
                    assetSelected = True
                    assetTypeSelected = "Switch Asset"

    # make window
    homeWindow.withdraw()
    AssetTreeViewWindow = tk.Toplevel(homeWindow)
    AssetTreeViewWindow.title("Asset Tree")
    AssetTreeViewWindow.configure(bg='gray15')
    window_width = 600
    window_height = 450
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    AssetTreeViewWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    # make tree view widget
    tree = ttk.Treeview(AssetTreeViewWindow, height=20)
    tree.heading('#0', text='Client Assets', anchor='w')
    tree.column('#0', width=250)
    tree.grid(row=0, column=0, rowspan=20, sticky='n', padx=5, pady=5)
    tree.bind('<<TreeviewSelect>>', updateAsset_ID)

    # add buttons to window
    btn_addSwitchAsset = tk.Button(AssetTreeViewWindow, width=10, text="Manage Switch Assets", bg='white',
                                   command=lambda: manageSwitchAssets(client_ID))
    btn_addSIWAsset = tk.Button(AssetTreeViewWindow, width=10, text="Manage SIW Assets", bg='white',
                                command=lambda: manageSIWAssets(client_ID))
    btn_linkAsset = tk.Button(AssetTreeViewWindow, width=10, text="Link Assets", bg='white',
                              command=lambda: checkVar("LinkAsset", buildingSelected))
    btn_unlinkAsset = tk.Button(AssetTreeViewWindow, width=10, text="Unlink Assets", bg='white',
                                command=lambda: checkVar("UnlinkAsset", assetSelected))
    btn_manageBuildings = tk.Button(AssetTreeViewWindow, width=10, text="Manage Client Buildings", bg='white',
                                    command=lambda: manageBuildings(client_ID))
    btn_home = tk.Button(AssetTreeViewWindow, width=10, text="Home", bg='azure4',
                         command=lambda: home(AssetTreeViewWindow, homeWindow))

    btn_addSwitchAsset.grid(row=0, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    btn_addSIWAsset.grid(row=1, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    btn_manageBuildings.grid(row=2, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    btn_linkAsset.grid(row=3, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    btn_unlinkAsset.grid(row=4, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    btn_home.grid(row=5, column=1, pady=4, padx=5, ipadx=50, sticky='n')

    # create connection to database populate tree view with asset data
    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()

    # get client data from database and assign to variable.
    cursor.execute("""SELECT p.Client_ID, tbl_Companies.Name
                        FROM tbl_Client_Info As p
                        INNER JOIN tbl_Contacts ON p.Contact_ID = tbl_Contacts.Contact_ID
                        INNER JOIN tbl_Companies ON tbl_Contacts.Company_ID = tbl_Companies.Company_ID
                        WHERE p.Client_ID =  """ + str(client_ID))
    clients = cursor.fetchall()

    iidCount = 1
    clientList = []
    buildingsList = []
    SIWAssetList = []
    SwitchList = []

    # iterate through clients and add to toplevel of treeview. Data is reassigned to list variable for future use.
    for record in clients:
        clientRecord = []
        clientRecord.append(record[0])  # ID
        clientRecord.append(record[1])  # Name
        clientRecord.append(iidCount)  # IID
        clientList.append(clientRecord)
        iidCount += 1
    for record in clientList:
        tree.insert('', tk.END, text=record[1], iid=record[2])

    # next, get Building assets and assign results to variable
    cursor.execute("""SELECT
	                        p.Client_ID,
	                        tbl_Companies.Name,
	                        tbl_Buildings.BuildingName,
	                        tbl_Buildings.Building_ID
                        FROM tbl_Client_Info As p
	                    INNER JOIN tbl_Contacts ON p.Contact_ID = tbl_Contacts.Contact_ID
                        INNER JOIN tbl_Companies ON tbl_Contacts.Company_ID = tbl_Companies.Company_ID
                        INNER JOIN tbl_Buildings ON p.Client_ID = tbl_Buildings.Client_ID
                        WHERE tbl_Buildings.Deleted IS NULL AND p.Client_ID = """ + str(client_ID))
    buildings = cursor.fetchall()

    # iterate through buildings, assign to second level of tree view and to corresponding clients. Data is reassigned
    # to a list variable for future use. (easier than using the variable generated by cursor.fetchall())
    for record in buildings:
        for client in clientList:
            if record[0] == client[0]:
                clientIID = client[2]
        buildingRecord = []
        buildingRecord.append(record[3])  # ID
        buildingRecord.append(record[2])  # Name
        buildingRecord.append(clientIID)  # ClientIID
        buildingRecord.append(iidCount)  # IID
        buildingsList.append(buildingRecord)
        iidCount += 1
    for record in buildingsList:
        tree.insert(record[2], tk.END, text=record[1], iid=record[3])

    if len(buildingsList) != 0:
        # now get SIW assets, assign results to variable, iterate and assign to third level to corresponding client
        # and buildings. Data is also reassigned to new variable for reasons stated above.
        cursor.execute("""SELECT
                            p.Client_ID,
                            tbl_Companies.Name,
                            tbl_Buildings.BuildingName,
                            tbl_Buildings.Building_ID,
                            tbl_SIWAssets.HostName,
                            tbl_SIWAssets.SIWAsset_ID
                            FROM tbl_Client_Info As p
                            INNER JOIN tbl_Contacts ON p.Contact_ID = tbl_Contacts.Contact_ID
                            INNER JOIN tbl_Companies ON tbl_Contacts.Company_ID = tbl_Companies.Company_ID
                            INNER JOIN tbl_Buildings ON p.Client_ID = tbl_Buildings.Client_ID
                            INNER JOIN tbl_SIWAssets ON tbl_buildings.Building_ID = tbl_SIWAssets.Building_ID
                            WHERE tbl_SIWAssets.Building_ID IS NOT NULL AND tbl_SIWAssets.Deleted IS NULL
                            """)

        # now do the same thing for switch assets
        SIWAssets = cursor.fetchall()
        for record in SIWAssets:
            for building in buildingsList:
                if record[3] == building[0]:
                    buildingIID = building[3]
                    SIWRecord = []
                    SIWRecord.append(record[5])  # ID
                    SIWRecord.append(record[4])  # Name
                    SIWRecord.append(buildingIID)  # BuildingIID
                    SIWRecord.append(iidCount)  # IID
                    SIWAssetList.append(SIWRecord)
                    iidCount += 1
        for record in SIWAssetList:
            tree.insert(record[2], tk.END, text=record[1], iid=record[3])

        cursor.execute("""SELECT
                                p.Client_ID,
                                tbl_Companies.Name,
                                tbl_Buildings.BuildingName,
                                tbl_Buildings.Building_ID,
                                tbl_SwitchAssets.HostName,
                                tbl_SwitchAssets.SwitchAsset_ID
                            FROM tbl_Client_Info As p
                            INNER JOIN tbl_Contacts ON p.Contact_ID = tbl_Contacts.Contact_ID
                            INNER JOIN tbl_Companies ON tbl_Contacts.Company_ID = tbl_Companies.Company_ID
                            INNER JOIN tbl_Buildings ON p.Client_ID = tbl_Buildings.Client_ID
                            INNER JOIN tbl_SwitchAssets ON tbl_buildings.Building_ID = tbl_SwitchAssets.Building_ID
                            WHERE tbl_SwitchAssets.Building_ID IS NOT NULL AND tbl_SwitchAssets.Deleted IS NULL
                            """)

        SwitchAssets = cursor.fetchall()
        for record in SwitchAssets:
            for building in buildingsList:
                if record[3] == building[0]:
                    buildingIID = building[3]
                    Switchrecord = []
                    iidCount += 1
                    Switchrecord.append(record[5])  # ID
                    Switchrecord.append(record[4])  # Name
                    Switchrecord.append(buildingIID)  # BuildingIID
                    Switchrecord.append(iidCount)  # IID
                    SwitchList.append(Switchrecord)
        for record in SwitchList:
            tree.insert(record[2], tk.END, text=record[1], iid=record[3])
        # end treeview creation
    else:
        return


def manageBuildings(client_ID):
    global buildingsWindow

    def updateWidgets(event):
        try:
            if entry_BuildingName.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                UpdateInfo = Selection.get('values')
                Building_Var.set(UpdateInfo[1])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        except:
            pass
        finally:
            global entry_BuildingName
            global entry_AssetType
            global Btn_Cancel
            global Btn_Save
            global Building_Var

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                BuildingName = Selection.get('values')

                SQL = """UPDATE tbl_Buildings
                        SET BuildingName = '{0}', AssetType_ID = {1}
                        WHERE Building_ID = {2}""".format(entry_BuildingName.get(), 4,
                                                          BuildingName[0])
                cursor.execute(SQL)
                cursor.commit()
                buildingsWindow.destroy()
                manageBuildings(client_ID)

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            BuildingName = Selection.get('values')

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(buildingsWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)
            Building_Var = tk.StringVar()

            lbl_BuildingName = tk.Label(Frame1, text="Building Name", bg='white', anchor='w')
            entry_BuildingName = tk.Entry(Frame1, textvariable=Building_Var, width=30)
            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_BuildingName.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            entry_BuildingName.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            Btn_Cancel.grid(row=8, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=8, column=1, pady=5, padx=5, sticky='w')
            curItem = tree.focus()
            Selection = tree.item(curItem)
            UpdateInfo = Selection.get('values')
            Building_Var.set(UpdateInfo[1])

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        VulnInfo = Selection.get('values')
        SQL = """UPDATE tbl_Buildings
                    SET Deleted = 1
                    WHERE tbl_Buildings.Building_ID= {0}""".format(VulnInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        AssetTreeViewWindow.destroy()
        showAssetTree(client_ID)
        manageBuildings(client_ID)
        return

    def newEntry(client_ID):
        global Frame1
        try:
            Frame1.grid_forget()
        except:
            pass
        finally:
            global Btn_Cancel
            global Btn_Save

            def save(client_ID):

                BuildingName = 'NULL'

                if entry_BuildingName.get() != '':
                    BuildingName = entry_BuildingName.get()

                SQL = """INSERT INTO tbl_Buildings (Client_ID, BuildingName, AssetType_ID)
                                VALUES ({0},'{1}',{2})""".format(client_ID, BuildingName, 4)
                cursor.execute(SQL)
                cursor.commit()
                AssetTreeViewWindow.destroy()
                showAssetTree(client_ID)
                manageBuildings(client_ID)
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(buildingsWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)
            lbl_BuildingName = tk.Label(Frame1, text="Building Name", bg='white', anchor='w')
            entry_BuildingName = tk.Entry(Frame1, width=30)
            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=lambda: save(client_ID))

            # Place widgets in window

            lbl_BuildingName.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            entry_BuildingName.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            Btn_Cancel.grid(row=8, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=8, column=1, pady=5, padx=5, sticky='w')

    buildingsWindow = tk.Toplevel(AssetTreeViewWindow)
    AssetTreeViewWindow.withdraw()
    buildingsWindow.geometry("600x300")
    buildingsWindow.title("Buildings")
    buildingsWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    buildingsWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(buildingsWindow, width=15, text="Home", bg='white',
                         command=lambda: home(buildingsWindow, homeWindow))
    btn_back = tk.Button(buildingsWindow, width=15, text="Back", bg='white',
                         command=lambda: home(buildingsWindow, AssetTreeViewWindow))
    btn_edit = tk.Button(buildingsWindow, width=15, text="Edit Entry", bg='white', command=edit)
    btn_new = tk.Button(buildingsWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry(client_ID))
    btn_delete = tk.Button(buildingsWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(buildingsWindow, column=('#1', '#2'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Building')
    tree.column('#1', width=40)
    tree.column('#2', width=850)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """select tbl_buildings.Building_ID, tbl_Buildings.BuildingName from tbl_Buildings WHERE tbl_Buildings.Deleted IS NULL AND Client_ID = """ + str(
        client_ID)
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def addSIWAsset():
    def populateCombo(comboToPopulate):
        modelList = []
        conn = pyodbc.connect(cnxn)
        cursor = conn.cursor()

        if comboToPopulate == "Ent_Model":
            cursor.execute("""SELECT tbl_Models.ModelName
                              FROM tbl_Models""")
            models = cursor.fetchall()
            for model in models:
                modelList.append(model[0])
            Ent_Model['values'] = modelList
        elif comboToPopulate == "Ent_NICs":
            NICsList = ['0', '1', '2', '3']
            Ent_Nics['values'] = NICsList

        return

    def addAsset():
        if Ent_HostName.get() != '':
            conn = pyodbc.connect(cnxn)
            cursor = conn.cursor()
            SQL = '''INSERT INTO tbl_SIWAssets (HostName, AssetType_ID)
                      VALUES ('{0}',{1})'''.format(Ent_HostName.get(), 19)
            cursor.execute(SQL)
            cursor.commit()

            SQL = """SELECT tbl_SIWAssets.SIWAsset_ID 
                    FROM tbl_SIWAssets
                    WHERE HostName = '{0}'""".format(Ent_HostName.get())
            cursor.execute(SQL)
            asset_ID = cursor.fetchall()
            asset_ID = asset_ID[0][-1]

            SQL = '''INSERT INTO tbl_SIW_Info (SIWAsset_ID, Description, DateCollected)
                      VALUES ({0},'{1}', '{2}')'''.format(asset_ID, Ent_Description.get(), Ent_DateCollected.get())
            cursor.execute(SQL)
            cursor.commit()

            SQL = """SELECT tbl_Models.Model_ID FROM tbl_Models WHERE ModelName = '{0}'""".format(Ent_Model.get())
            cursor.execute(SQL)
            ModelRecord = cursor.fetchall()
            model_ID = ModelRecord[0][0]

            SQL = '''INSERT INTO tbl_SIW_Hardware_Info (SIWAsset_ID, Model_ID, SerialNumber, OSName, NICs)
                      VALUES ({0},{1},'{2}', '{3}', {4})'''.format(asset_ID, model_ID, Ent_SerialNumber.get(),
                                                                   Ent_OSName.get(), Ent_Nics.get())
            cursor.execute(SQL)
            cursor.commit()

            SQL = '''INSERT INTO tbl_SIW_Vuln_Info (SIWAsset_ID)
                      VALUES ({0})'''.format(asset_ID)
            cursor.execute(SQL)
            cursor.commit()
            addSIWAssetWindow.destroy()
            siwAssetsWindow.destroy()
            manageSIWAssets(client_ID)

    global addSIWAssetWindow
    siwAssetsWindow.withdraw()
    addSIWAssetWindow = tk.Toplevel(AssetTreeViewWindow)
    addSIWAssetWindow.title("Add a New SIW Asset")
    addSIWAssetWindow.configure(bg='gray15')
    window_width = 600
    window_height = 450
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    addSIWAssetWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    Lbl_HostName = tk.Label(addSIWAssetWindow, width=5, text="HostName")
    Lbl_Description = tk.Label(addSIWAssetWindow, width=5, text="Description")
    Lbl_DateCollected = tk.Label(addSIWAssetWindow, width=5, text="Date Collected")
    Lbl_Model = tk.Label(addSIWAssetWindow, width=5, text="Model")
    Lbl_SerialNumber = tk.Label(addSIWAssetWindow, width=5, text="Serial Number")
    Lbl_OSName = tk.Label(addSIWAssetWindow, width=5, text="Operating System")
    Lbl_Nics = tk.Label(addSIWAssetWindow, width=5, text="Number of NICs")

    Ent_HostName = tk.Entry(addSIWAssetWindow, width=10)
    Ent_Description = tk.Entry(addSIWAssetWindow, width=10)
    Ent_DateCollected = tk.Entry(addSIWAssetWindow, width=10)
    Ent_Model = ttk.Combobox(addSIWAssetWindow, width=8)
    Ent_SerialNumber = tk.Entry(addSIWAssetWindow, width=10)
    Ent_OSName = tk.Entry(addSIWAssetWindow, width=10)
    Ent_Nics = ttk.Combobox(addSIWAssetWindow, width=8)

    btn_back = tk.Button(addSIWAssetWindow, width=5, text="Back", bg='azure4',
                         command=lambda: home(addSIWAssetWindow, siwAssetsWindow))
    btn_Home = tk.Button(addSIWAssetWindow, width=5, text="Home", bg='azure4',
                         command=lambda: home(AssetTreeViewWindow, homeWindow))
    btn_addSIWAsset = tk.Button(addSIWAssetWindow, width=5, text="Save", bg='azure4', command=addAsset)

    Lbl_HostName.grid(row=0, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    Ent_HostName.grid(row=0, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    Lbl_Description.grid(row=1, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    Ent_Description.grid(row=1, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    Lbl_DateCollected.grid(row=2, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    Ent_DateCollected.grid(row=2, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    Lbl_Model.grid(row=3, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    Ent_Model.grid(row=3, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    Lbl_SerialNumber.grid(row=4, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    Ent_SerialNumber.grid(row=4, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    Lbl_OSName.grid(row=5, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    Ent_OSName.grid(row=5, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    Lbl_Nics.grid(row=6, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    Ent_Nics.grid(row=6, column=1, pady=5, padx=5, ipadx=50, sticky='n')

    btn_back.grid(row=7, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    btn_Home.grid(row=8, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    btn_addSIWAsset.grid(row=7, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    populateCombo("Ent_Model")
    populateCombo("Ent_NICs")

    return


def linkAsset(building_ID):
    global clientRecords
    global deviceAsset_Name
    global linkAssetsWindow
    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    cursor.execute("""SELECT BuildingName FROM tbl_Buildings
                                WHERE Building_ID = """ + str(building_ID))
    buildingName = cursor.fetchall()

    def updateDeviceAsset_ID(event):
        global deviceAsset_Name
        global assetType
        global deviceAssetID
        deviceAssetID = ' '
        assetType = ""
        currSelection = lbl_deviceAssets.curselection()
        if currSelection != ():
            deviceAsset_Name = lbl_deviceAssets.get(currSelection)
        for asset in siwAssets:
            if deviceAsset_Name == asset[1]:
                deviceAssetID = asset[0]
                assetType = "SIWAsset"
        for asset in switchAssets:
            if deviceAsset_Name == asset[1]:
                deviceAssetID = asset[0]
                assetType = "SwitchAsset"

    def populateListBox(ListBox, Records):
        if ListBox == 'lbl_Buildings':
            for i in range(len(Records)):
                record = Records[i]
                lbl_deviceAssets.insert('end', record[1])
            lbl_deviceAssets.selection_set(0)
            lbl_deviceAssets.event_generate("<<ListboxSelect>>")

    def link(deviceAssetID, building_ID, assetType):
        if deviceAssetID != ' ':
            conn = pyodbc.connect(cnxn)
            cursor = conn.cursor()
            if assetType == "SIWAsset":
                sqlScript = "UPDATE tbl_SIWAssets SET Building_ID = " + str(building_ID) + "WHERE SIWAsset_ID = " + str(
                    deviceAssetID)
                cursor.execute(sqlScript)
                conn.commit()
            if assetType == "SwitchAsset":
                sqlScript = "UPDATE tbl_SwitchAssets SET Building_ID = " + str(
                    building_ID) + "WHERE SwitchAsset_ID = " + str(deviceAssetID)
                cursor.execute(sqlScript)
                conn.commit()
            conn.close()
            lbl_deviceAssets.delete(lbl_deviceAssets.curselection())
        else:
            return

    AssetTreeViewWindow.withdraw()

    # make window
    linkAssetsWindow = tk.Toplevel(AssetTreeViewWindow)
    linkAssetsWindow.title("Link Assets")
    linkAssetsWindow.configure(bg='gray15')
    window_width = 600
    window_height = 450
    screen_width = linkAssetsWindow.winfo_screenwidth()
    screen_height = linkAssetsWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    linkAssetsWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    lbl_deviceAssets = tk.Listbox(linkAssetsWindow)
    lbl_deviceAssets.bind('<<ListboxSelect>>', updateDeviceAsset_ID)
    lbl_buildingName = tk.Label(linkAssetsWindow, text="Link to: " + buildingName[0][0])
    btn_home = tk.Button(linkAssetsWindow, width=10, text="Home", bg='azure4',
                         command=lambda: home(AssetTreeViewWindow, homeWindow))
    btn_back = tk.Button(linkAssetsWindow, width=10, text="Back to Asset Tree", bg='azure4',
                         command=lambda: home(linkAssetsWindow, AssetTreeViewWindow))
    btn_link = tk.Button(linkAssetsWindow, width=10, text="Link", bg='white',
                         command=lambda: link(deviceAssetID, building_ID, assetType))

    lbl_buildingName.grid(row=0, column=0, columnspan=1, rowspan=1, pady=4, padx=5, ipadx=50, sticky='n')
    lbl_deviceAssets.grid(row=1, column=0, columnspan=1, rowspan=10, pady=4, padx=5, ipadx=50, sticky='n')
    btn_home.grid(row=1, column=1, columnspan=1, rowspan=1, pady=4, padx=5, ipadx=50, sticky='n')
    btn_back.grid(row=2, column=1, columnspan=1, rowspan=1, pady=4, padx=5, ipadx=50, sticky='n')
    btn_link.grid(row=12, column=0, columnspan=1, rowspan=1, pady=4, padx=5, ipadx=50, sticky='n')

    cursor.execute("""SELECT * FROM tbl_SIWAssets
                        WHERE Building_ID IS NULL""")
    siwAssets = cursor.fetchall()

    cursor.execute("""SELECT * FROM tbl_SwitchAssets
                        WHERE Building_ID IS NULL""")
    switchAssets = cursor.fetchall()

    # assign query results to list box
    populateListBox('lbl_Buildings', siwAssets)
    populateListBox('lbl_Buildings', switchAssets)

    return


# ----------------------------------------
#            Begin Reports
# ----------------------------------------

def reports(client_ID):
    def reportPick(reportToGenerate, client_ID):
        if Cmb_Reports.get() != "":
            if Cmb_Reports.get() == "Switch Report by MAC Address":
                generateReport("Switch Report by MAC Address", client_ID)
        else:
            messagebox.showinfo("Error", "Please Choose a Report")

        return

    homeWindow.withdraw()
    reportWindow = tk.Toplevel(homeWindow)
    reportWindow.title("aeSolutions")
    reportWindow.configure(bg='gray15')
    window_width = 475
    window_height = 250
    screen_width = reportWindow.winfo_screenwidth()
    screen_height = reportWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    reportWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    # Create buttons
    btn_GenerateReports = tk.Button(reportWindow, width=5, text="Generate Report", bg='white',
                                    command=lambda: reportPick(reportToGenerate, (client_ID)))
    Cmb_Reports = ttk.Combobox(reportWindow, width=50)

    btn_GenerateReports.grid(row=1, column=1, columnspan=1, pady=10, padx=5, ipadx=75, ipady=20)
    Cmb_Reports.grid(row=0, column=1, columnspan=3)
    # Cmb_Reports.bind("<<ComboboxSelected>>", updateClient_ID)
    reportList = ["Switch Report by MAC Address",
                  "Switch Access Report",
                  "Switch Port Description Report",
                  "Computer Security Summary Report",
                  "Switch Multiple MACs on Access Ports",
                  "SMBv1 Patch Report",
                  "Network Security Report",
                  "Zone Security Report",
                  "Asset Summary by Zone Report",
                  "Software Vulnerabilities Report (CVE)",
                  "Total Number of Switches / Computers by Zone",
                  "Updates by Category Report",
                  "USB History Report"]
    reportToGenerate = ""
    Cmb_Reports["values"] = reportList
    return


def generateReport(reportToGenerate, client_ID):
    SQL = """SELECT p.HostName, tbl_SwitchPorts.PortNumber, tbl_SwitchPorts.Description, tbl_SwitchPorts.AdminStatus, tbl_SwitchPorts.PortType, tbl_SwitchPorts.NativeVLAN, tbl_SwitchPorts.VLANTag, tbl_Buildings.Client_ID

    FROM tbl_SwitchAssets as p
	INNER JOIN tbl_SwitchPorts ON p.SwitchAsset_ID = tbl_SwitchPorts.SwitchAsset_ID
	INNER JOIN tbl_Buildings ON p.Building_ID = tbl_Buildings.Building_ID
	INNER JOIN tbl_Client_Info ON tbl_Buildings.Client_ID = tbl_Client_Info.Client_ID

	WHERE tbl_Client_Info.Client_ID = {0}""".format(client_ID)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()

    # get client data from database and assign to variable.
    cursor.execute(SQL)
    records = cursor.fetchall()

    pd.DataFrame(records).to_excel('output.xlsx', header=False, index=False)
    return


# ----------------------------------------
#          Begin Other Tables
# ----------------------------------------

def otherTables():
    global otherTablesWindow
    homeWindow.withdraw()
    otherTablesWindow = tk.Toplevel(homeWindow)
    otherTablesWindow.title("Manage Other Tables")
    otherTablesWindow.configure(bg='gray15')
    window_width = 600
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    otherTablesWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    # add buttons to window
    btn_ManageVulnNames = tk.Button(otherTablesWindow, width=10, text="Vulnerabilty Names", bg='white',
                                    command=manageVulnNames)
    btn_ManageUSBClasses = tk.Button(otherTablesWindow, width=10, text="USB Classes", bg='white',
                                     command=manageUSBClasses)
    btn_ManageUpdateCategories = tk.Button(otherTablesWindow, width=10, text="Update Categories", bg='white',
                                           command=manageUpdateCategories)
    btn_ManageProtocols = tk.Button(otherTablesWindow, width=10, text="Protocols", bg='white', command=manageProtocols)
    btn_ManageAttackTypes = tk.Button(otherTablesWindow, width=10, text="Attack Types", bg='white',
                                      command=manageAttackTypes)
    btn_ManageSoftwareSources = tk.Button(otherTablesWindow, width=10, text="Software Sources", bg='white',
                                          command=manageSoftwareSources)
    btn_ManageMisconfigurations = tk.Button(otherTablesWindow, width=10, text="Misconfigurations", bg='white',
                                            command=manageMisconfigurations)
    btn_ManageAssetTypes = tk.Button(otherTablesWindow, width=10, text="Asset Types", bg='white',
                                     command=manageAssetTypes)
    btn_ManageManModels = tk.Button(otherTablesWindow, width=10, text="Models/Manufactureres", bg='white',
                                    command=manageModelsManufacturers)
    btn_CVEs = tk.Button(otherTablesWindow, width=10, text="CVEs", bg='white',
                         command=manageCVEs)
    btn_home = tk.Button(otherTablesWindow, width=10, text="Home", bg='azure4',
                         command=lambda: home(otherTablesWindow, homeWindow))
    btn_ManageVulnNames.grid(row=0, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    btn_ManageUSBClasses.grid(row=1, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    btn_ManageUpdateCategories.grid(row=2, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    btn_ManageProtocols.grid(row=3, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    btn_ManageAttackTypes.grid(row=4, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    btn_ManageSoftwareSources.grid(row=5, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    btn_ManageMisconfigurations.grid(row=6, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    btn_ManageAssetTypes.grid(row=7, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    btn_ManageManModels.grid(row=8, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    btn_CVEs.grid(row=9, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    btn_home.grid(row=0, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    return


def manageVulnNames():
    def updateWidgets(event):
        try:
            if Entry_Vulnerability.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                VulnInfo = Selection.get('values')
                vuln_Var.set(VulnInfo[1])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:

            global lbl_Vulnerability
            global Entry_Vulnerability

            global vuln_Var

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                VulnInfo = Selection.get('values')

                VulnName = 'NULL'
                VulnName = Entry_Vulnerability.get()

                SQL = """UPDATE tbl_VulnNames
                        SET Name = '{0}'
                        WHERE VulnName_ID = {1}""".format(VulnName, VulnInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageVulnNames()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            VulnInfo = Selection.get('values')

            vuln_Var = tk.StringVar()

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Vulnerability = tk.Label(Frame1, text="Vulnerability", bg='white', anchor='w')
            Entry_Vulnerability = tk.Entry(Frame1, textvariable=vuln_Var, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Vulnerability.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Vulnerability.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            curItem = tree.focus()
            Selection = tree.item(curItem)
            VulnInfo = Selection.get('values')
            vuln_Var.set(VulnInfo[1])

            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        VulnInfo = Selection.get('values')
        SQL = """UPDATE tbl_VulnNames
                    SET Deleted = 1
                    WHERE tbl_VulnNames.VulnName_ID = {0}""".format(VulnInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global lbl_Vulnerability
            global Entry_Vulnerability

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save():
                VulnName = 'NULL'
                VulnName = Entry_Vulnerability.get()

                SQL = """INSERT INTO tbl_VulnNames (Name)
                                VALUES ('{0}')""".format(VulnName)
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageVulnNames()
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Vulnerability = tk.Label(Frame1, text="Vulnerability", bg='white', anchor='w')
            Entry_Vulnerability = tk.Entry(Frame1, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=lambda: save())

            # Place widgets in window
            lbl_Vulnerability.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Vulnerability.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            return

    VulnWindow = tk.Toplevel(otherTablesWindow)
    otherTablesWindow.withdraw()
    VulnWindow.geometry("600x300")
    VulnWindow.title("Vulnerabilities")
    VulnWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    VulnWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(VulnWindow, width=15, text="Home", bg='white',
                         command=lambda: home(VulnWindow, homeWindow))
    btn_back = tk.Button(VulnWindow, width=15, text="Back", bg='white',
                         command=lambda: home(VulnWindow, otherTablesWindow))
    btn_edit = tk.Button(VulnWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(VulnWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry())
    btn_delete = tk.Button(VulnWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(VulnWindow, column=('#1', '#2'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Vulnerability Name')
    tree.column('#1', width=40)
    tree.column('#2', width=850)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_VulnNames.VulnName_ID,
                tbl_VulnNames.Name
                FROM tbl_VulnNames
                WHERE tbl_VulnNames.Deleted IS NULL"""
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def manageUSBClasses():
    def updateWidgets(event):
        try:
            if Entry_Vulnerability.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                USBInfo = Selection.get('values')
                Class_Var.set(USBInfo[1])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:

            global lbl_ClassName
            global Entry_ClassName

            global Class_Var

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                USBInfo = Selection.get('values')

                className = 'NULL'
                className = Entry_ClassName.get()

                SQL = """UPDATE tbl_USBClasses
                        SET Name = '{0}'
                        WHERE Class_ID = {1}""".format(className, USBInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageUSBClasses()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            USBInfo = Selection.get('values')

            Class_Var = tk.StringVar()

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_ClassName = tk.Label(Frame1, text="Class Name", bg='white', anchor='w')
            Entry_ClassName = tk.Entry(Frame1, textvariable=Class_Var, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_ClassName.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_ClassName.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            Class_Var.set(USBInfo[1])

            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        VulnInfo = Selection.get('values')
        SQL = """UPDATE tbl_VulnNames
                    SET Deleted = 1
                    WHERE tbl_VulnNames.VulnName_ID = {0}""".format(VulnInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global lbl_Class
            global Entry_Class

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save():
                className = 'NULL'
                className = Entry_Class.get()

                SQL = """INSERT INTO tbl_USBClasses (Name)
                                VALUES ('{0}')""".format(className)
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageUSBClasses()
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Class = tk.Label(Frame1, text="Class Name", bg='white', anchor='w')
            Entry_Class = tk.Entry(Frame1, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Class.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Class.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            return

    VulnWindow = tk.Toplevel(otherTablesWindow)
    otherTablesWindow.withdraw()
    VulnWindow.geometry("600x300")
    VulnWindow.title("USB Classes")
    VulnWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    VulnWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(VulnWindow, width=15, text="Home", bg='white',
                         command=lambda: home(VulnWindow, homeWindow))
    btn_back = tk.Button(VulnWindow, width=15, text="Back", bg='white',
                         command=lambda: home(VulnWindow, otherTablesWindow))
    btn_edit = tk.Button(VulnWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(VulnWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry())
    btn_delete = tk.Button(VulnWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(VulnWindow, column=('#1', '#2'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Class Name')
    tree.column('#1', width=40)
    tree.column('#2', width=850)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_USBClasses.Class_ID,
                tbl_USBClasses.Name
                FROM tbl_USBClasses
                WHERE tbl_USBClasses.Deleted IS NULL"""
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def manageUpdateCategories():
    def updateWidgets(event):
        try:
            if Entry_Vulnerability.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                CategoryInfo = Selection.get('values')
                Category_Var.set(CategoryInfo[1])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:

            global lbl_Category
            global Entry_Category

            global Category_Var

            global Btn_Cancel
            global Btn_Save

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                USBInfo = Selection.get('values')

                CategoryName = 'NULL'
                CategoryName = Entry_Category.get()

                SQL = """UPDATE tbl_UpdatesCategories
                        SET Name = '{0}'
                        WHERE Category_ID = {1}""".format(CategoryName, USBInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageUpdateCategories()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            CategoryInfo = Selection.get('values')

            Category_Var = tk.StringVar()

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Category = tk.Label(Frame1, text="Category Name", bg='white', anchor='w')
            Entry_Category = tk.Entry(Frame1, textvariable=Category_Var, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Category.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Category.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            Category_Var.set(CategoryInfo[1])

            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        CategoryInfo = Selection.get('values')
        SQL = """UPDATE tbl_UpdatesCategories
                    SET Deleted = 1
                    WHERE tbl_UpdatesCategories.Category_ID = {0}""".format(CategoryInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global lbl_Category
            global Entry_Category

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save():
                categoryName = 'NULL'
                categoryName = Entry_Category.get()

                SQL = """INSERT INTO tbl_UpdatesCategories (Name)
                                VALUES ('{0}')""".format(categoryName)
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageUpdateCategories()
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Category = tk.Label(Frame1, text="Category Name", bg='white', anchor='w')
            Entry_Category = tk.Entry(Frame1, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Category.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Category.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            return

    VulnWindow = tk.Toplevel(otherTablesWindow)
    otherTablesWindow.withdraw()
    VulnWindow.geometry("600x300")
    VulnWindow.title("Updates Categories")
    VulnWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    VulnWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(VulnWindow, width=15, text="Home", bg='white',
                         command=lambda: home(VulnWindow, homeWindow))
    btn_back = tk.Button(VulnWindow, width=15, text="Back", bg='white',
                         command=lambda: home(VulnWindow, otherTablesWindow))
    btn_edit = tk.Button(VulnWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(VulnWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry())
    btn_delete = tk.Button(VulnWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(VulnWindow, column=('#1', '#2'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Category Name')
    tree.column('#1', width=40)
    tree.column('#2', width=850)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_UpdatesCategories.Category_ID,
                tbl_UpdatesCategories.Name
                FROM tbl_UpdatesCategories
                WHERE tbl_UpdatesCategories.Deleted IS NULL"""
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def manageAttackTypes():
    def updateWidgets(event):
        try:
            if Entry_Attack.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                AInfo = Selection.get('values')
                Attack_Var.set(AInfo[1])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:

            global lbl_Attack
            global Entry_Attack

            global Attack_Var

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                AInfo = Selection.get('values')

                attackName = 'NULL'
                attackName = Entry_Attack.get()

                SQL = """UPDATE tbl_AttackTypes
                        SET Name = '{0}'
                        WHERE tbl_AttackTypes.AttackType_ID = {1}""".format(attackName, AInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageAttackTypes()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            AInfo = Selection.get('values')

            Attack_Var = tk.StringVar()

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Attack = tk.Label(Frame1, text="Attack Name", bg='white', anchor='w')
            Entry_Attack = tk.Entry(Frame1, textvariable=Attack_Var, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Attack.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Attack.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            Attack_Var.set(AInfo[1])

            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        AInfo = Selection.get('values')
        SQL = """UPDATE tbl_AttackTypes
                    SET Deleted = 1
                    WHERE tbl_AttackTypes.AttackType_ID = {0}""".format(AInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global lbl_AttackTypes
            global Entry_AttackTypes

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save():
                AttackName = 'NULL'
                AttackName = Entry_AttackTypes.get()

                SQL = """INSERT INTO tbl_AttackTypes (Name)
                                VALUES ('{0}')""".format(AttackName)
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageAttackTypes()
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_AttackTypes = tk.Label(Frame1, text="Attack Name", bg='white', anchor='w')
            Entry_AttackTypes = tk.Entry(Frame1, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_AttackTypes.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_AttackTypes.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            return

    VulnWindow = tk.Toplevel(otherTablesWindow)
    otherTablesWindow.withdraw()
    VulnWindow.geometry("600x300")
    VulnWindow.title("Vulnerabilities")
    VulnWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    VulnWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(VulnWindow, width=15, text="Home", bg='white',
                         command=lambda: home(VulnWindow, homeWindow))
    btn_back = tk.Button(VulnWindow, width=15, text="Back", bg='white',
                         command=lambda: home(VulnWindow, otherTablesWindow))
    btn_edit = tk.Button(VulnWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(VulnWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry())
    btn_delete = tk.Button(VulnWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(VulnWindow, column=('#1', '#2'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Attack Type Name')
    tree.column('#1', width=40)
    tree.column('#2', width=850)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_AttackTypes.AttackType_ID,
                tbl_AttackTypes.Name
                FROM tbl_AttackTypes
                WHERE tbl_AttackTypes.Deleted IS NULL"""
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def manageSoftwareSources():
    def updateWidgets(event):
        try:
            if Entry_Software.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                SSInfo = Selection.get('values')
                Software_Var.set(SSInfo[1])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:

            global lbl_Software
            global Entry_Software

            global Software_Var

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                SSInfo = Selection.get('values')

                softwareName = 'NULL'
                softwareName = Entry_Software.get()

                SQL = """UPDATE tbl_SoftwareSources
                        SET Name = '{0}'
                        WHERE tbl_SoftwareSources.Source_ID = {1}""".format(softwareName, SSInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageSoftwareSources()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            SSInfo = Selection.get('values')

            Software_Var = tk.StringVar()

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Software = tk.Label(Frame1, text="Software Name", bg='white', anchor='w')
            Entry_Software = tk.Entry(Frame1, textvariable=Software_Var, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Software.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Software.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            Software_Var.set(SSInfo[1])

            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        SSInfo = Selection.get('values')
        SQL = """UPDATE tbl_SoftwareSources
                    SET Deleted = 1
                    WHERE tbl_SoftwareSources.Source_ID = {0}""".format(SSInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global lbl_SS
            global Entry_SS

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save():
                softwareName = 'NULL'
                softwareName = Entry_SS.get()

                SQL = """INSERT INTO tbl_SoftwareSources (Name)
                                VALUES ('{0}')""".format(softwareName)
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageSoftwareSources()
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_SS = tk.Label(Frame1, text="Source Name", bg='white', anchor='w')
            Entry_SS = tk.Entry(Frame1, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_SS.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_SS.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            return

    VulnWindow = tk.Toplevel(otherTablesWindow)
    otherTablesWindow.withdraw()
    VulnWindow.geometry("600x300")
    VulnWindow.title("Software Sources")
    VulnWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    VulnWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(VulnWindow, width=15, text="Home", bg='white',
                         command=lambda: home(VulnWindow, homeWindow))
    btn_back = tk.Button(VulnWindow, width=15, text="Back", bg='white',
                         command=lambda: home(VulnWindow, otherTablesWindow))
    btn_edit = tk.Button(VulnWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(VulnWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry())
    btn_delete = tk.Button(VulnWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(VulnWindow, column=('#1', '#2'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Class Name')
    tree.column('#1', width=40)
    tree.column('#2', width=850)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_SoftwareSources.Source_ID,
                tbl_SoftwareSources.Name
                FROM tbl_SoftwareSources
                WHERE tbl_SoftwareSources.Deleted IS NULL"""
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def manageProtocols():
    def updateWidgets(event):
        try:
            if Entry_Protocol.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                PInfo = Selection.get('values')
                Proto_Var.set(PInfo[1])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:

            global lbl_Protocol
            global Entry_Protocol

            global Proto_Var

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                PInfo = Selection.get('values')

                protoName = 'NULL'
                protoName = Entry_Protocol.get()

                SQL = """UPDATE tbl_Protocols
                        SET Name = '{0}'
                        WHERE Protocol_ID = {1}""".format(protoName, PInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageProtocols()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            PInfo = Selection.get('values')

            Proto_Var = tk.StringVar()

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Protocol = tk.Label(Frame1, text="Protocol", bg='white', anchor='w')
            Entry_Protocol = tk.Entry(Frame1, textvariable=Proto_Var, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Protocol.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Protocol.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            Proto_Var.set(PInfo[1])

            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        ProtocolInfo = Selection.get('values')
        SQL = """UPDATE tbl_Protocols
                    SET Deleted = 1
                    WHERE tbl_Protocols.Protocol_ID = {0}""".format(ProtocolInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global lbl_Pro
            global Entry_Pro

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save():
                protoName = 'NULL'
                protoName = Entry_Pro.get()

                SQL = """INSERT INTO tbl_Protocols (Name)
                                VALUES ('{0}')""".format(protoName)
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageProtocols()
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Pro = tk.Label(Frame1, text="Protocol", bg='white', anchor='w')
            Entry_Pro = tk.Entry(Frame1, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Pro.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Pro.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            return

    VulnWindow = tk.Toplevel(otherTablesWindow)
    otherTablesWindow.withdraw()
    VulnWindow.geometry("600x300")
    VulnWindow.title("Vulnerabilities")
    VulnWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    VulnWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(VulnWindow, width=15, text="Home", bg='white',
                         command=lambda: home(VulnWindow, homeWindow))
    btn_back = tk.Button(VulnWindow, width=15, text="Back", bg='white',
                         command=lambda: home(VulnWindow, otherTablesWindow))
    btn_edit = tk.Button(VulnWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(VulnWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry())
    btn_delete = tk.Button(VulnWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(VulnWindow, column=('#1', '#2'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Class Name')
    tree.column('#1', width=40)
    tree.column('#2', width=850)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_Protocols.Protocol_ID,
                tbl_Protocols.Name
                FROM tbl_Protocols
                WHERE tbl_Protocols.Deleted IS NULL"""
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def manageAssetTypes():
    def updateWidgets(event):
        try:
            if lbl_AssetType.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                AssetTypeInfo = Selection.get('values')
                AssetType_Var.set(AssetTypeInfo[1])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:

            global lbl_AssetType
            global Entry_AssetType

            global AssetType_Var

            global Btn_Cancel
            global Btn_Save

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                AssetTypeInfo = Selection.get('values')

                AssetTypeName = 'NULL'
                AssetTypeName = Entry_AssetType.get()

                SQL = """UPDATE tbl_AssetTypes
                        SET Name = '{0}'
                        WHERE AssetType_ID = {1}""".format(AssetTypeName, AssetTypeInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageAssetTypes()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            AssetTypeInfo = Selection.get('values')

            AssetType_Var = tk.StringVar()

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_AssetType = tk.Label(Frame1, text="Asset Type Name", bg='white', anchor='w')
            Entry_AssetType = tk.Entry(Frame1, textvariable=AssetType_Var, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_AssetType.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_AssetType.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            AssetType_Var.set(AssetTypeInfo[1])

            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        AssetTypeInfo = Selection.get('values')
        SQL = """UPDATE tbl_AssetTypes
                    SET Deleted = 1
                    WHERE tbl_AssetTypes.AssetType_ID = {0}""".format(AssetTypeInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global lbl_AssetType
            global Entry_AssetType

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save():
                assetTypeName = 'NULL'
                assetTypeName = Entry_AssetType.get()

                SQL = """INSERT INTO tbl_AssetTypes (Name)
                                VALUES ('{0}')""".format(assetTypeName)
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageAssetTypes()
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_AssetType = tk.Label(Frame1, text="Asset Type Name", bg='white', anchor='w')
            Entry_AssetType = tk.Entry(Frame1, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_AssetType.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_AssetType.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            return

    VulnWindow = tk.Toplevel(otherTablesWindow)
    otherTablesWindow.withdraw()
    VulnWindow.geometry("600x300")
    VulnWindow.title("Vulnerabilities")
    VulnWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    VulnWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(VulnWindow, width=15, text="Home", bg='white',
                         command=lambda: home(VulnWindow, homeWindow))
    btn_back = tk.Button(VulnWindow, width=15, text="Back", bg='white',
                         command=lambda: home(VulnWindow, otherTablesWindow))
    btn_edit = tk.Button(VulnWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(VulnWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry())
    btn_delete = tk.Button(VulnWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(VulnWindow, column=('#1', '#2'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Asset Type Name')
    tree.column('#1', width=40)
    tree.column('#2', width=850)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_AssetTypes.AssetType_ID,
                tbl_AssetTypes.Name
                FROM tbl_AssetTypes
                WHERE tbl_AssetTypes.Deleted IS NULL"""
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def manageMisconfigurations():
    def updateWidgets(event):
        try:
            if Entry_Vulnerability.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                MisconfigInfo = Selection.get('values')
                Misconfig_Var.set(MisconfigInfo[1])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:

            global lbl_Misconfig
            global Entry_Misconfig

            global Misconfig_Var

            global Btn_Cancel
            global Btn_Save
            global MisconfigList

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                MisconfigInfo = Selection.get('values')

                MisconfigName = 'NULL'
                MisconfigName = Entry_Misconfig.get()

                SQL = """UPDATE tbl_Misconfigurations
                        SET Name = '{0}'
                        WHERE Error_ID = {1}""".format(MisconfigName, MisconfigInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageMisconfigurations()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            MisconfigInfo = Selection.get('values')

            Misconfig_Var = tk.StringVar()

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Misconfig = tk.Label(Frame1, text="Misconfigurations", bg='white', anchor='w')
            Entry_Misconfig = tk.Entry(Frame1, textvariable=Misconfig_Var, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Misconfig.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Misconfig.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            curItem = tree.focus()
            Selection = tree.item(curItem)
            MisconfigInfo = Selection.get('values')
            Misconfig_Var.set(MisconfigInfo[1])

            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        MisconfigInfo = Selection.get('values')
        SQL = """UPDATE tbl_Misconfigurations
                    SET Deleted = 1
                    WHERE tbl_Misconfigurations.Error_ID = {0}""".format(MisconfigInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global lbl_Misconfig
            global Entry_Misconfig

            global Btn_Cancel
            global Btn_Save
            global MisconfigList

            def save():
                MisconfigName = 'NULL'
                MisconfigName = Entry_Misconfig.get()

                SQL = """INSERT INTO tbl_Misconfigurations (Name)
                                VALUES ('{0}')""".format(MisconfigName)
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                manageMisconfigurations()
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Misconfig = tk.Label(Frame1, text="Misconfigurations", bg='white', anchor='w')
            Entry_Misconfig = tk.Entry(Frame1, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=lambda: save())

            # Place widgets in window
            lbl_Misconfig.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_Misconfig.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            return

    VulnWindow = tk.Toplevel(otherTablesWindow)
    otherTablesWindow.withdraw()
    VulnWindow.geometry("600x300")
    VulnWindow.title("Vulnerabilities")
    VulnWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    VulnWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(VulnWindow, width=15, text="Home", bg='white',
                         command=lambda: home(VulnWindow, homeWindow))
    btn_back = tk.Button(VulnWindow, width=15, text="Back", bg='white',
                         command=lambda: home(VulnWindow, otherTablesWindow))
    btn_edit = tk.Button(VulnWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(VulnWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry())
    btn_delete = tk.Button(VulnWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(VulnWindow, column=('#1', '#2'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Misconfigurations Name')
    tree.column('#1', width=40)
    tree.column('#2', width=850)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_Misconfigurations.Error_ID,
                tbl_Misconfigurations.Name
                FROM tbl_Misconfigurations
                WHERE tbl_Misconfigurations.Deleted IS NULL"""
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def manageModelsManufacturers():
    global ModelsWindow

    def updateWidgets(event):
        try:
            if Frame1.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                ModelInfo = Selection.get('values')
                Misconfig_Var.set(ModelInfo[1])
                Cmb_Manufacturer.set(ModelInfo[2])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global Entry_ModelName
            global Cmb_Manufacturer

            global Misconfig_Var

            global Btn_Cancel
            global Btn_Save
            global MisconfigList

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                ModelInfo = Selection.get('values')

                ModelName = 'NULL'
                Manufacturer_ID = 'NULL'
                ModelName = Entry_ModelName.get()
                for record in manufacturers:
                    if Cmb_Manufacturer.get() == record[0]:
                        Manufacturer_ID = record[1]
                SQL = """UPDATE tbl_Models
                        SET ModelName = '{0}', Manufacturer_ID = {1}
                        WHERE Model_ID = {2}""".format(ModelName, Manufacturer_ID, ModelInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                ModelsWindow.destroy()
                manageModelsManufacturers()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            ModelInfo = Selection.get('values')

            Misconfig_Var = tk.StringVar()

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(ModelsWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_ModelName = tk.Label(Frame1, text="Model Name", bg='white', anchor='w')
            lbl_Manufacturer = tk.Label(Frame1, text="Manufacturer", bg='white', anchor='w')
            Entry_ModelName = tk.Entry(Frame1, textvariable=Misconfig_Var, width=50)
            Cmb_Manufacturer = ttk.Combobox(Frame1)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_ModelName.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Manufacturer.grid(row=1, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_ModelName.grid(row=0, column=1, pady=5, padx=5, sticky='w')
            Cmb_Manufacturer.grid(row=1, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            SQL = """SELECT
                            tbl_Manufacturers.ManufacturerName,
                            tbl_Manufacturers.Manufacturer_ID

                            FROM tbl_Manufacturers
                            WHERE tbl_Manufacturers.Deleted IS NULL"""
            cursor.execute(SQL)
            manufacturers = cursor.fetchall()
            manufacturerList = []
            for record in manufacturers:
                manufacturerList.append(record[0])
            Cmb_Manufacturer["values"] = manufacturerList
            Cmb_Manufacturer.set(ModelInfo[2])
            Misconfig_Var.set(ModelInfo[1])

            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        ModelInfo = Selection.get('values')
        SQL = """UPDATE tbl_Models
                    SET Deleted = 1
                    WHERE tbl_Models.Model_ID = {0}""".format(ModelInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global lbl_Misconfig
            global Entry_Misconfig

            global Btn_Cancel
            global Btn_Save
            global MisconfigList

            def save():
                ModelName = 'NULL'
                Manufacturer_ID = 'NULL'
                ModelName = Entry_ModelName.get()
                for record in manufacturers:
                    if Cmb_Manufacturer.get() == record[0]:
                        Manufacturer_ID = record[1]
                SQL = """Insert Into tbl_Models (ModelName, Manufacturer_ID)
                                        Values('{0}', {1})""".format(ModelName, Manufacturer_ID)
                cursor.execute(SQL)
                cursor.commit()
                ModelsWindow.destroy()
                manageModelsManufacturers()
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(ModelsWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_ModelName = tk.Label(Frame1, text="Model Name", bg='white', anchor='w')
            lbl_Manufacturer = tk.Label(Frame1, text="Manufacturer", bg='white', anchor='w')
            Entry_ModelName = tk.Entry(Frame1, width=50)
            Cmb_Manufacturer = ttk.Combobox(Frame1)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_ModelName.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Manufacturer.grid(row=1, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_ModelName.grid(row=0, column=1, pady=5, padx=5, sticky='w')
            Cmb_Manufacturer.grid(row=1, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            SQL = """SELECT
                                        tbl_Manufacturers.ManufacturerName,
                                        tbl_Manufacturers.Manufacturer_ID

                                        FROM tbl_Manufacturers
                                        WHERE tbl_Manufacturers.Deleted IS NULL"""
            cursor.execute(SQL)
            manufacturers = cursor.fetchall()
            manufacturerList = []
            for record in manufacturers:
                manufacturerList.append(record[0])
            Cmb_Manufacturer["values"] = manufacturerList

            return

    ModelsWindow = tk.Toplevel(otherTablesWindow)
    otherTablesWindow.withdraw()
    ModelsWindow.title("Vulnerabilities")
    ModelsWindow.configure(bg='gray15')
    window_width = 1020
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    ModelsWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(ModelsWindow, width=15, text="Home", bg='white',
                         command=lambda: home(ModelsWindow, homeWindow))
    btn_back = tk.Button(ModelsWindow, width=15, text="Back", bg='white',
                         command=lambda: home(ModelsWindow, otherTablesWindow))
    btn_edit = tk.Button(ModelsWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(ModelsWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry())
    btn_delete = tk.Button(ModelsWindow, width=15, text="Delete Entry", bg='white', command=delete)
    btn_Manufacturers = tk.Button(ModelsWindow, width=15, text="Manufacturers", bg='white', command=viewManufacturers)

    btn_home.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=5, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')
    btn_Manufacturers.grid(row=1, column=3, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(ModelsWindow, column=('#1', '#2', '#3', '#4'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Model Name')
    tree.heading('#3', text='Manufacturer Name')
    tree.heading('#4', text='Manufacturer ID')
    tree.column('#1', width=40)
    tree.column('#2', width=350)
    tree.column('#3', width=300)
    tree.column('#4', width=150)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_Models.Model_ID,
                tbl_Models.ModelName,
                tbl_Manufacturers.Manufacturer_ID,
                tbl_Manufacturers.ManufacturerName

                FROM tbl_Models
                INNER JOIN tbl_Manufacturers ON tbl_Models.Manufacturer_ID = tbl_Manufacturers.Manufacturer_ID
                WHERE tbl_Models.Deleted IS NULL"""
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1], record[3], record[2]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def viewManufacturers():
    def updateWidgets(event):
        try:
            if Frame1.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                ModelInfo = Selection.get('values')
                Misconfig_Var.set(ModelInfo[1])
                Cmb_Manufacturer.set(ModelInfo[2])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global Entry_ModelName
            global Cmb_Manufacturer

            global Misconfig_Var

            global Btn_Cancel
            global Btn_Save
            global MisconfigList

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                ModelInfo = Selection.get('values')

                ManufacturerName = 'NULL'
                ManufacturerName = Entry_ManufacturerName.get()
                SQL = """UPDATE tbl_Manufacturers
                        SET ManufacturerName = '{0}'
                        WHERE Manufacturer_ID = {1}""".format(ManufacturerName, ModelInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                ManufacturersWindow.destroy()
                viewManufacturers()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            ModelInfo = Selection.get('values')

            Misconfig_Var = tk.StringVar()

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(ManufacturersWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_ManufacturerName = tk.Label(Frame1, text="Manufacturer Name", bg='white', anchor='w')
            Entry_ManufacturerName = tk.Entry(Frame1, textvariable=Misconfig_Var, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_ManufacturerName.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_ManufacturerName.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            Misconfig_Var.set(ModelInfo[1])

            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        ModelInfo = Selection.get('values')
        SQL = """UPDATE tbl_Manufacturers
                    SET Deleted = 1
                    WHERE tbl_Manufacturers.Manufacturer_ID = {0}""".format(ModelInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global lbl_Misconfig
            global Entry_Misconfig

            global Btn_Cancel
            global Btn_Save
            global MisconfigList

            def save():
                ManufacturerName = 'NULL'
                ManufacturerName = Entry_ManufacturerName.get()
                SQL = """INSERT INTO tbl_Manufacturers (ManufacturerName)
                            VALUES ('{0}')""".format(ManufacturerName)
                cursor.execute(SQL)
                cursor.commit()
                ManufacturersWindow.destroy()
                viewManufacturers()
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(ManufacturersWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_ManufacturerName = tk.Label(Frame1, text="Manufacturer Name", bg='white', anchor='w')
            Entry_ManufacturerName = tk.Entry(Frame1, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_ManufacturerName.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Entry_ManufacturerName.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            return

    ManufacturersWindow = tk.Toplevel(ModelsWindow)
    ModelsWindow.withdraw()
    ManufacturersWindow.title("Manufacturers")
    ManufacturersWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    ManufacturersWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(ManufacturersWindow, width=15, text="Home", bg='white',
                         command=lambda: home(ManufacturersWindow, homeWindow))
    btn_back = tk.Button(ManufacturersWindow, width=15, text="Back", bg='white',
                         command=lambda: home(ManufacturersWindow, ModelsWindow))
    btn_edit = tk.Button(ManufacturersWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(ManufacturersWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry())
    btn_delete = tk.Button(ManufacturersWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=5, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(ManufacturersWindow, column=('#1', '#2'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Manufacturer Name')
    tree.column('#1', width=40)
    tree.column('#2', width=850)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_Manufacturers.Manufacturer_ID,
                tbl_Manufacturers.ManufacturerName

                FROM tbl_Manufacturers
                WHERE tbl_Manufacturers.Deleted IS NULL"""
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def manageCVEs():
    global CVEsWindow

    def updateWidgets(event):
        try:
            if lbl_CVE.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                SoftwareInfo = Selection.get('values')
                CVEName_Var.set(SoftwareInfo[1])
                Cmb_AttackType.set(SoftwareInfo[2])
                CVSSv3_Var.set(SoftwareInfo[3])
                CVSSv2_Var.set(SoftwareInfo[4])
                URL_Var.set(SoftwareInfo[5])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global lbl_CVE
        global Cmb_AttackType
        global CVEName_Var
        global CVSSv3_Var
        global CVSSv2_Var
        global URL_Var
        global Frame1
        try:
            Frame1.grid_forget()
        finally:

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                SoftwareInfo = Selection.get('values')

                CVE_Name = entry_CVEName.get()
                CVSSv3 = entry_CVSSv3.get()
                CVSSv2 = entry_CVSSv2.get()
                URL = entry_URL.get()
                for record in attackTypes:
                    if Cmb_AttackType.get() == record[0]:
                        AttackType_ID = record[1]

                SQL = """UPDATE tbl_CVEs
                        SET Name = '{0}', CVSSv3 = '{1}', CVSSv2 = '{2}', URL = '{3}', AttackType_ID = {4}
                        WHERE CVE_ID = {5}""".format(CVE_Name,
                                                     CVSSv3, CVSSv2, URL, AttackType_ID,
                                                     SoftwareInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                CVEsWindow.destroy()
                manageCVEs()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            SoftwareInfo = Selection.get('values')
            # Create Entry Widgets/ Combobox
            CVEName_Var = tk.StringVar()
            CVSSv3_Var = tk.StringVar()
            CVSSv2_Var = tk.StringVar()
            URL_Var = tk.StringVar()

            Frame1 = tk.Frame(CVEsWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_CVE = tk.Label(Frame1, text="CVE Name", bg='white', anchor='w')
            lbl_AttackType = tk.Label(Frame1, text="Attack Type", bg='white', anchor='w')
            lbl_CVSSv3 = tk.Label(Frame1, text="CVSSv3", bg='white', anchor='w')
            lbl_CVSSv2 = tk.Label(Frame1, text="CVSSv2", bg='white', anchor='w')
            lbl_URL = tk.Label(Frame1, text="URL", bg='white', anchor='w')

            entry_CVEName = tk.Entry(Frame1, textvariable=CVEName_Var, width=40)
            Cmb_AttackType = ttk.Combobox(Frame1)
            entry_CVSSv3 = tk.Entry(Frame1, textvariable=CVSSv3_Var, width=15)
            entry_CVSSv2 = tk.Entry(Frame1, textvariable=CVSSv2_Var, width=15)
            entry_URL = tk.Entry(Frame1, textvariable=URL_Var, width=15)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_CVE.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_AttackType.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_CVSSv3.grid(row=4, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_CVSSv2.grid(row=5, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_URL.grid(row=6, column=0, pady=5, padx=5, sticky='w', columnspan=2)

            entry_CVEName.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            Cmb_AttackType.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_CVSSv3.grid(row=4, column=1, pady=5, padx=5, sticky='w')
            entry_CVSSv2.grid(row=5, column=1, pady=5, padx=5, sticky='w')
            entry_URL.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            Btn_Cancel.grid(row=8, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=8, column=1, pady=5, padx=5, sticky='w')
            # Set the first two entry widgets values
            CVEName_Var.set(SoftwareInfo[1])
            CVSSv3_Var.set(SoftwareInfo[3])
            CVSSv2_Var.set(SoftwareInfo[4])
            URL_Var.set(SoftwareInfo[5])

            # Execute query and use results to fill CVEs Combobox and set the combobox value
            conn = pyodbc.connect(cnxn)
            cursor = conn.cursor()
            SQL = """SELECT
            tbl_AttackTypes.Name,
            tbl_AttackTypes.AttackType_ID
            FROM tbl_AttackTypes
            ORDER BY tbl_AttackTypes.Name ASC"""
            cursor.execute(SQL)
            attackTypes = cursor.fetchall()
            attackTypesList = []
            for record in attackTypes:
                attackTypesList.append(record[0])
            Cmb_AttackType["values"] = attackTypesList
            Cmb_AttackType.set(SoftwareInfo[2])
            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        SoftwareInfo = Selection.get('values')
        SQL = """UPDATE tbl_CVEs
                    SET Deleted = 1
                    WHERE tbl_CVEs.CVE_ID = {0}""".format(SoftwareInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global lbl_Misconfig
            global Entry_Misconfig

            global Btn_Cancel
            global Btn_Save
            global MisconfigList

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                SoftwareInfo = Selection.get('values')

                CVE_Name = entry_CVEName.get()
                CVSSv3 = entry_CVSSv3.get()
                CVSSv2 = entry_CVSSv2.get()
                URL = entry_URL.get()
                for record in attackTypes:
                    if Cmb_AttackType.get() == record[0]:
                        AttackType_ID = record[1]

                SQL = """INSERT INTO tbl_CVEs (Name, CVSSv3, CVSSv2, URL, AttackType_ID)
                        VALUES ('{0}','{1}', '{2}', '{3}', {4})""".format(CVE_Name,
                                                                          CVSSv3, CVSSv2, URL, AttackType_ID,
                                                                          SoftwareInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                CVEsWindow.destroy()
                manageCVEs()
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(CVEsWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_CVE = tk.Label(Frame1, text="CVE Name", bg='white', anchor='w')
            lbl_AttackType = tk.Label(Frame1, text="Attack Type", bg='white', anchor='w')
            lbl_CVSSv3 = tk.Label(Frame1, text="CVSSv3", bg='white', anchor='w')
            lbl_CVSSv2 = tk.Label(Frame1, text="CVSSv2", bg='white', anchor='w')
            lbl_URL = tk.Label(Frame1, text="URL", bg='white', anchor='w')

            entry_CVEName = tk.Entry(Frame1, width=40)
            Cmb_AttackType = ttk.Combobox(Frame1)
            entry_CVSSv3 = tk.Entry(Frame1, width=15)
            entry_CVSSv2 = tk.Entry(Frame1, width=15)
            entry_URL = tk.Entry(Frame1, width=15)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_CVE.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_AttackType.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_CVSSv3.grid(row=4, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_CVSSv2.grid(row=5, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_URL.grid(row=6, column=0, pady=5, padx=5, sticky='w', columnspan=2)

            entry_CVEName.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            Cmb_AttackType.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_CVSSv3.grid(row=4, column=1, pady=5, padx=5, sticky='w')
            entry_CVSSv2.grid(row=5, column=1, pady=5, padx=5, sticky='w')
            entry_URL.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            Btn_Cancel.grid(row=8, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=8, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=7, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=7, column=1, pady=5, padx=5, sticky='w')
            SQL = """SELECT
                        tbl_AttackTypes.Name,
                        tbl_AttackTypes.AttackType_ID
                        FROM tbl_AttackTypes
                        ORDER BY tbl_AttackTypes.Name ASC"""
            cursor.execute(SQL)
            attackTypes = cursor.fetchall()
            attackTypesList = []
            for record in attackTypes:
                attackTypesList.append(record[0])
            Cmb_AttackType["values"] = attackTypesList

            return

    CVEsWindow = tk.Toplevel(otherTablesWindow)
    otherTablesWindow.withdraw()
    CVEsWindow.title("CVEs")
    CVEsWindow.configure(bg='gray15')
    window_width = 950
    window_height = 600
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    CVEsWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(CVEsWindow, width=15, text="Home", bg='white',
                         command=lambda: home(CVEsWindow, homeWindow))
    btn_back = tk.Button(CVEsWindow, width=15, text="Back", bg='white',
                         command=lambda: home(CVEsWindow, homeWindow))
    btn_edit = tk.Button(CVEsWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(CVEsWindow, width=15, text="New Entry", bg='white', comman=lambda: newEntry())
    btn_delete = tk.Button(CVEsWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(CVEsWindow, column=('#1', '#2', '#3', '#4', '#5', '#6'), show='headings')
    tree.column('#1', width=30)
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='CVE')
    tree.heading('#3', text='Attack Type')
    tree.heading('#4', text='CVSSv3')
    tree.heading('#5', text='CVSSv2')
    tree.heading('#6', text='URL')
    tree.column('#2', width=150)
    tree.column('#4', width=100)
    tree.column('#5', width=100)
    tree.column('#6', width=400)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_CVEs.CVE_ID,
                tbl_CVEs.Name,
                tbl_CVEs.CVSSv3,
                tbl_CVEs.CVSSv2,
                tbl_CVEs.URL,
                tbl_AttackTypes.Name

                FROM tbl_CVEs
                INNER JOIN tbl_AttackTypes ON tbl_CVEs.AttackType_ID = tbl_AttackTypes.AttackType_ID
                WHERE tbl_CVEs.Deleted IS NULL""".format()
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count,
                        values=(record[0], record[1], record[5], record[3], record[3], record[4]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)


# ----------------------------------------
#          Begin SIW Assets
# ----------------------------------------

def manageSIWAssets(client_ID):
    global siwAssetsWindow
    global count
    global count2
    Mlist = []
    MoList = []

    def populateListBox(ListBox, Records):
        if ListBox == 'Lbl_AllSIWAssets':
            for i in range(len(Records)):
                record = Records[i]
                Lbl_AllSIWAssets.insert('end', record[1])
            Lbl_AllSIWAssets.select_set(0)
            Lbl_AllSIWAssets.event_generate("<<ListboxSelect>>")

    # function within function allows the global var SIWAsset_ID to be changed - which is passed later on
    def updateListboxSel(event):
        global SIWAsset_ID

        currSelection = Lbl_AllSIWAssets.curselection()

        if currSelection != ():
            SIWAsset = Lbl_AllSIWAssets.get(currSelection)
            indexNum = currSelection[0] + 1
            for i in range(len(siwAssets)):
                record = siwAssets[i]
                if record[1] == SIWAsset:
                    SIWAsset_ID = record[0]
            populateGeneralInfo(SIWAsset_ID)

    def populateGeneralInfo(SIWAsset_ID):
        count = 0
        count2 = 0
        MoList = []
        # create connection to database and execute SQL code
        conn = pyodbc.connect(cnxn)
        cursor = conn.cursor()
        cursor.execute(
            """SELECT
                tbl_SIW_Info.Description, 
                tbl_SIW_Info.DateCollected,
                tbl_Manufacturers.ManufacturerName,
                tbl_Manufacturers.Manufacturer_ID,
                tbl_Models.ModelName,
                tbl_Models.Model_ID,
                tbl_SIW_Hardware_Info.SerialNumber,
                tbl_SIW_Hardware_Info.OSName,
                tbl_SIW_Hardware_Info.NICs,
                tbl_SIW_Vuln_Info.PatchCycle,
                tbl_SIW_Vuln_Info.PatchDays,
                tbl_SIW_Vuln_Info.MemoryLoad,
                tbl_SIW_Vuln_Info.USBDays,
                tbl_SIW_Vuln_Info.MS17010,
                tbl_SIW_Vuln_Info.MimiKatz,
                tbl_SIW_Vuln_Info.VulnerablePorts,
                tbl_AssetTypes.Name
                FROM tbl_SIWAssets as p

                Inner Join tbl_SIW_Info On p.SIWAsset_ID = tbl_SIW_Info.SIWAsset_ID
                Inner Join tbl_SIW_Hardware_Info On p.SIWAsset_ID = tbl_SIW_Hardware_Info.SIWAsset_ID
                Inner Join tbl_SIW_Vuln_Info On p.SIWAsset_ID = tbl_SIW_Vuln_Info.SIWAsset_ID
                Inner Join tbl_Models On tbl_SIW_Hardware_Info.Model_ID = tbl_Models.Model_ID
                Inner Join tbl_Manufacturers On tbl_Models.Manufacturer_ID = tbl_Manufacturers.Manufacturer_ID
                Inner Join tbl_AssetTypes On p.AssetType_ID = tbl_AssetTypes.AssetType_ID

                Where p.SIWAsset_ID = """ + str(SIWAsset_ID))
        SIWInfo = cursor.fetchall()
        for record in SIWInfo:
            description_var.set(record[0])
            dateCollected_var.set(record[1])
            manufacturerName = record[2]
            manufacturer_ID = record[3]
            modelName = record[4]
            model_ID = record[5]
            serialNumber_var.set(record[6])
            OS_var.set(record[7])
            NICs_var.set(record[8])
        # set Manufacturer
        for record in Manufacturers:
            if record[0] == manufacturerName:
                Cmb_manufacturer.current(count)
            count += 1

        # append Models to cmb_Model
        for record in ModelManufacturer:
            if manufacturer_ID == record[2]:
                MoList.append(record[1])
        Cmb_model["values"] = MoList

        # set Model
        for record in MoList:
            if record == modelName:
                Cmb_model.current(count2)
            count2 += 1
        Cmb_AssetType.set(SIWInfo[0][16])
        return

    def updateModelManufacturer(event):
        newModelList = []
        for record in ModelManufacturer:
            if record[3] == Cmb_manufacturer.get():
                newModelList.append(record[1])
        Cmb_model["values"] = newModelList
        Cmb_model.set(newModelList[0])
        return

    def deleteSIWAsset(SIWAsset_ID):

        conn = pyodbc.connect(cnxn)
        cursor = conn.cursor()
        cursor.execute("""UPDATE tbl_SIWAssets
                            SET	Deleted = 1
                            WHERE SIWAsset_ID = """ + str(SIWAsset_ID))
        cursor.commit()
        Lbl_AllSIWAssets.delete(Lbl_AllSIWAssets.curselection())
        return

    def edit():
        global btn_Save
        global btn_Cancel
        btn_Save = tk.Button(Frame1, text="Save", bg='white', command=save)
        btn_Cancel = tk.Button(Frame1, text="Cancel", bg='white', command=cancel)
        btn_Save.place(x=150, y=450, width=120)
        btn_Cancel.place(x=300, y=450, width=120)
        entry_Description['state'] = "normal"
        entry_DateCollected['state'] = "normal"
        Cmb_manufacturer['state'] = "normal"
        Cmb_model['state'] = "normal"
        Cmb_AssetType['state'] = "normal"
        entry_serialNumber['state'] = "normal"
        entry_OS['state'] = "normal"
        NICs_OS['state'] = "normal"

        return

    def cancel():
        entry_Description['state'] = "disabled"
        entry_DateCollected['state'] = "disabled"
        Cmb_manufacturer['state'] = "disabled"
        Cmb_model['state'] = "disabled"
        Cmb_AssetType['state'] = "disabled"
        entry_serialNumber['state'] = "disabled"
        entry_OS['state'] = "disabled"
        NICs_OS['state'] = "disabled"
        btn_Save.place_forget()
        btn_Cancel.place_forget()
        return

    def save():
        global SIWAsset_ID
        for record in ModelManufacturer:
            if record[1] == Cmb_model.get():
                manufacturer_ID = record[2]
                model_ID = record[0]
        for record in assetTypes:
            if record[0] == Cmb_AssetType.get():
                assetType_ID = record[1]

        SQL = """UPDATE tbl_SIW_Hardware_Info
                            SET Model_ID = {0}, SerialNumber = '{1}', OSName = '{2}', NICs = '{3}'
                            WHERE SIWAsset_ID = {4}""".format(model_ID, entry_serialNumber.get(), entry_OS.get(),
                                                              NICs_OS.get(), SIWAsset_ID)
        cursor.execute(SQL)
        cursor.commit()
        SQL = """Update tbl_SIWAssets
                    SET AssetType_ID = {0}
                    WHERE SIWAsset_ID = {1}""".format(assetType_ID, SIWAsset_ID)
        cursor.execute(SQL)
        cursor.commit()
        SQL = """Update tbl_SIW_Info
                            SET Description = '{0}', DateCollected = '{1}'
                            WHERE SIWAsset_ID = {2}""".format(entry_Description.get(), entry_DateCollected.get(),
                                                              SIWAsset_ID)
        cursor.execute(SQL)
        cursor.commit()
        siwAssetsWindow.destroy()
        manageSIWAssets(client_ID)
        return

    # make window
    AssetTreeViewWindow.withdraw()
    siwAssetsWindow = tk.Toplevel(homeWindow)
    siwAssetsWindow.geometry("600x300")
    siwAssetsWindow.title("SIW Assets")
    siwAssetsWindow.configure(bg='gray15')
    window_width = 870
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    siwAssetsWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    Lbl_AllSIWAssets = tk.Listbox(siwAssetsWindow)
    Lbl_AllSIWAssets.place(x=5, y=5, width=150)
    Lbl_AllSIWAssets.bind('<<ListboxSelect>>', updateListboxSel)

    Frame1 = tk.Frame(siwAssetsWindow, padx=5, pady=5, bg='white')
    Frame1.place(x=160, y=5, width=500, height=500)

    label_GeneralInfo = tk.Label(Frame1, text="General Info:", bg='white')
    label_GeneralInfo.config(font=('TkDefaultFont', 15, 'bold'))
    label_GeneralInfo.place(x=185, y=5, width=120)

    label_AssetType = tk.Label(Frame1, text="AssetType", bg='white', anchor='w')
    label_AssetType.config(font=('TkDefaultFont', 10, 'bold'))
    label_AssetType.place(x=5, y=50, width=120)
    Cmb_AssetType = ttk.Combobox(Frame1)
    Cmb_AssetType.bind("<<ComboboxSelected>>", updateModelManufacturer)
    Cmb_AssetType['state'] = "disabled"
    Cmb_AssetType.place(x=130, y=50, width=200)

    label_Description = tk.Label(Frame1, text="Description", bg='white', anchor='w')
    label_Description.config(font=('TkDefaultFont', 10, 'bold'))
    label_Description.place(x=5, y=100, width=120)
    description_var = tk.StringVar()
    entry_Description = tk.Entry(Frame1, textvariable=description_var)
    entry_Description['state'] = "disabled"
    entry_Description.place(x=130, y=100, width=200)

    label_DateCollected = tk.Label(Frame1, text="Date Collected", bg='white', anchor='w')
    label_DateCollected.config(font=('TkDefaultFont', 10, 'bold'))
    label_DateCollected.place(x=5, y=150, width=120)
    dateCollected_var = tk.StringVar()
    entry_DateCollected = tk.Entry(Frame1, textvariable=dateCollected_var)
    entry_DateCollected['state'] = "disabled"
    entry_DateCollected.place(x=130, y=150, width=200)

    label_Manufacturer = tk.Label(Frame1, text="Manufacturer", bg='white', anchor='w')
    label_Manufacturer.config(font=('TkDefaultFont', 10, 'bold'))
    label_Manufacturer.place(x=5, y=200, width=120)
    manufacturer_var = tk.StringVar()
    Cmb_manufacturer = ttk.Combobox(Frame1)
    Cmb_manufacturer.bind("<<ComboboxSelected>>", updateModelManufacturer)
    Cmb_manufacturer['state'] = "disabled"
    Cmb_manufacturer.place(x=130, y=200, width=200)

    label_Model = tk.Label(Frame1, text="Model", bg='white', anchor='w')
    label_Model.config(font=('TkDefaultFont', 10, 'bold'))
    label_Model.place(x=5, y=250, width=120)
    model_var = tk.StringVar()
    Cmb_model = ttk.Combobox(Frame1, textvariable=model_var)
    Cmb_model.bind("<<ComboboxSelected>>", updateModelManufacturer)
    Cmb_model['state'] = "disabled"
    Cmb_model.place(x=130, y=250, width=200)

    label_SerialNumber = tk.Label(Frame1, text="Serial Number", bg='white', anchor='w')
    label_SerialNumber.config(font=('TkDefaultFont', 10, 'bold'))
    label_SerialNumber.place(x=5, y=300, width=120)
    serialNumber_var = tk.StringVar()
    entry_serialNumber = tk.Entry(Frame1, textvariable=serialNumber_var)
    entry_serialNumber['state'] = "disabled"
    entry_serialNumber.place(x=130, y=300, width=200)

    label_OperatingSystem = tk.Label(Frame1, text="Operating System", bg='white', anchor='w')
    label_OperatingSystem.config(font=('TkDefaultFont', 10, 'bold'))
    label_OperatingSystem.place(x=5, y=350, width=120)
    OS_var = tk.StringVar()
    entry_OS = tk.Entry(Frame1, textvariable=OS_var)
    entry_OS['state'] = "disabled"
    entry_OS.place(x=130, y=350, width=200)

    label_NICs = tk.Label(Frame1, text="NICs", bg='white', anchor='w')
    label_NICs.config(font=('TkDefaultFont', 10, 'bold'))
    label_NICs.place(x=5, y=400, width=120)
    NICs_var = tk.StringVar()
    NICs_OS = tk.Entry(Frame1, textvariable=NICs_var)
    NICs_OS['state'] = "disabled"
    NICs_OS.place(x=130, y=400, width=200)

    btn_Edit = tk.Button(Frame1, text="Edit", bg='white', command=edit)
    btn_Edit.place(x=5, y=450, width=120)

    # Create and place buttons on right side of window + Back and Home buttons
    btn_addAsset = tk.Button(siwAssetsWindow, width=5, text="Add Asset", bg='white', command=addSIWAsset)
    btn_deleteAsset = tk.Button(siwAssetsWindow, width=5, text="Delete Asset", bg='white',
                                command=lambda: deleteSIWAsset(SIWAsset_ID))
    btn_networkInfo = tk.Button(siwAssetsWindow, width=15, text="Network Info", bg='white',
                                command=lambda: viewNetworkInfo(SIWAsset_ID))
    btn_OpenPorts = tk.Button(siwAssetsWindow, width=15, text="Open Ports", bg='white',
                              command=lambda: viewOpenPorts(SIWAsset_ID))
    btn_Software = tk.Button(siwAssetsWindow, width=5, text="Software", bg='white',
                             command=lambda: viewSoftware(SIWAsset_ID))
    btn_Updates = tk.Button(siwAssetsWindow, width=5, text="Installed Updates", bg='white',
                            command=lambda: viewUpdates(SIWAsset_ID))
    btn_USBHistory = tk.Button(siwAssetsWindow, width=5, text="USB History", bg='white',
                               command=lambda: viewUSBHistory(SIWAsset_ID))
    btn_Vulnerabilities = tk.Button(siwAssetsWindow, width=5, text="Vulnerabilities", bg='white',
                                    command=lambda: viewVulnerabilities(SIWAsset_ID))
    btn_OtherInfo = tk.Button(siwAssetsWindow, width=5, text="Other Info", bg='white',
                              command=lambda: OtherInfo(SIWAsset_ID))
    btn_home = tk.Button(siwAssetsWindow, width=5, text="Home", bg='white',
                         command=lambda: home(siwAssetsWindow, homeWindow))
    btn_back = tk.Button(siwAssetsWindow, width=5, text="Back", bg='white',
                         command=lambda: home(siwAssetsWindow, AssetTreeViewWindow))

    btn_addAsset.place(x=665, y=5, width=200)
    btn_deleteAsset.place(x=665, y=55, width=200)
    btn_networkInfo.place(x=665, y=155, width=200)
    btn_OpenPorts.place(x=665, y=205, width=200)
    btn_Software.place(x=665, y=255, width=200)
    btn_Updates.place(x=665, y=305, width=200)
    btn_USBHistory.place(x=665, y=355, width=200)
    btn_Vulnerabilities.place(x=665, y=405, width=200)
    btn_OtherInfo.place(x=665, y=455, width=200)
    btn_home.place(x=5, y=260, width=150)
    btn_back.place(x=5, y=310, width=150)

    # create connection to database and execute SQL code
    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    cursor.execute("""SELECT * FROM tbl_SIWAssets
                    INNER JOIN tbl_Buildings ON tbl_SIWAssets.Building_ID = tbl_Buildings.Building_ID
                    INNER JOIN tbl_Client_Info ON tbl_Buildings.Client_ID = tbl_Client_Info.Client_ID
					INNER JOIN tbl_Contacts ON tbl_Client_Info.Contact_ID = tbl_Contacts.Contact_ID
                    WHERE tbl_SIWAssets.Deleted IS NULL AND tbl_Contacts.Company_ID = """ + str(client_ID))

    siwAssets = cursor.fetchall()
    # assign query results to list box

    cursor.execute("""SELECT tbl_Models.Model_ID, tbl_Models.ModelName, tbl_Models.Manufacturer_ID, tbl_Manufacturers.ManufacturerName
                        FROM tbl_Models
                        INNER JOIN tbl_Manufacturers ON tbl_Models.Manufacturer_ID = tbl_Manufacturers.Manufacturer_ID""")
    ModelManufacturer = cursor.fetchall()

    cursor.execute("""SELECT tbl_Manufacturers.ManufacturerName
                            FROM tbl_Manufacturers""")
    Manufacturers = cursor.fetchall()
    for record in Manufacturers:
        Mlist.append(record[0])
    Cmb_manufacturer["values"] = Mlist
    populateListBox('Lbl_AllSIWAssets', siwAssets)

    cursor.execute("""SELECT tbl_AssetTypes.Name, tbl_AssetTypes.AssetType_ID
                            FROM tbl_AssetTypes""")
    assetTypes = cursor.fetchall()
    assetList = []
    for record in assetTypes:
        assetList.append(record[0])
    Cmb_AssetType["values"] = assetList
    return


def viewNetworkInfo(SIWAsset_ID):
    def updateWidgets(event):
        try:
            if entry_Port.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                OpenPortInfo = Selection.get('values')
                Port_Var.set(OpenPortInfo[1])
                State_Var.set(OpenPortInfo[2])
                Cmb_Protocol.set(OpenPortInfo[3])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global entry_MAC
            global entry_IPAddress
            global entry_IPSubnet
            global entry_DefGateway

            global MAC_Var
            global IP_Var
            global Subnet_Var
            global Gateway_Var

            global Btn_Cancel
            global Btn_Save
            global CVEList

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                NetInfo = Selection.get('values')
                SQL = """UPDATE tbl_MACs
                                SET MACAddress = '{0}', IPAddress = '{1}', IPSubnet = '{2}', DefaultGateway = '{3}'
                                WHERE MACs_ID = {4}""".format(entry_MAC.get(), entry_IPAddress.get(),
                                                              entry_IPSubnet.get(), entry_DefGateway.get(), NetInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                networkWindow.destroy()
                viewNetworkInfo(SIWAsset_ID)
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            NetInfo = Selection.get('values')
            # Create Entry Widgets/ Combobox
            MAC_Var = tk.StringVar()
            IP_Var = tk.StringVar()
            Subnet_Var = tk.StringVar()
            Gateway_Var = tk.StringVar()

            Frame1 = tk.Frame(networkWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_MAC = tk.Label(Frame1, text="MAC Address", bg='white', anchor='w')
            lbl_IPAddres = tk.Label(Frame1, text="IP Address", bg='white', anchor='w')
            lbl_IPSubnet = tk.Label(Frame1, text="Subnet", bg='white', anchor='w')
            lbl_DefGateway = tk.Label(Frame1, text="Gateway", bg='white', anchor='w')

            entry_MAC = tk.Entry(Frame1, textvariable=MAC_Var, width=40)
            entry_IPAddress = tk.Entry(Frame1, textvariable=IP_Var, width=30)
            entry_IPSubnet = tk.Entry(Frame1, textvariable=Subnet_Var, width=15)
            entry_DefGateway = tk.Entry(Frame1, textvariable=Gateway_Var, width=15)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_MAC.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_IPAddres.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_IPSubnet.grid(row=4, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_DefGateway.grid(row=5, column=0, pady=5, padx=5, sticky='w', columnspan=2)

            entry_MAC.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_IPAddress.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_IPSubnet.grid(row=4, column=1, pady=5, padx=5, sticky='w')
            entry_DefGateway.grid(row=5, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=8, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=8, column=1, pady=5, padx=5, sticky='w')
            # Set the first two entry widgets values
            MAC_Var.set(NetInfo[1])
            IP_Var.set(NetInfo[2])
            Subnet_Var.set(NetInfo[3])
            Gateway_Var.set(NetInfo[4])

            return

        return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        NetInfo = Selection.get('values')
        SQL = """UPDATE tbl_MACs
                    SET Deleted = 1
                    WHERE MACs_ID = {0}""".format(NetInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry(SIWAsset_ID):

        def save(SIWAsset_ID):
            SQL = """INSERT INTO tbl_MACs (SIWAsset_ID, MACAddress, IPAddress, IPSubnet, DefaultGateway)
                        VALUES ({0},'{1}','{2}','{3}','{4}')""".format(SIWAsset_ID, entry_MAC.get(),
                                                                       entry_IPAddress.get(), entry_IPSubnet.get(),
                                                                       entry_DefGateway.get())
            cursor.execute(SQL)
            cursor.commit()
            networkWindow.destroy()
            viewNetworkInfo(SIWAsset_ID)
            return

        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global entry_MAC
            global entry_IPAddress
            global entry_IPSubnet
            global entry_DefGateway

            global MAC_Var
            global IP_Var
            global Subnet_Var
            global Gateway_Var

            global Btn_Cancel
            global Btn_Save
            global CVEList

            Frame1 = tk.Frame(networkWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_MAC = tk.Label(Frame1, text="MAC Address", bg='white', anchor='w')
            lbl_IPAddres = tk.Label(Frame1, text="IP Address", bg='white', anchor='w')
            lbl_IPSubnet = tk.Label(Frame1, text="Subnet", bg='white', anchor='w')
            lbl_DefGateway = tk.Label(Frame1, text="Gateway", bg='white', anchor='w')

            entry_MAC = tk.Entry(Frame1, width=40)
            entry_IPAddress = tk.Entry(Frame1, width=30)
            entry_IPSubnet = tk.Entry(Frame1, width=15)
            entry_DefGateway = tk.Entry(Frame1, width=15)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=lambda: save(SIWAsset_ID))

            # Place widgets in window
            lbl_MAC.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_IPAddres.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_IPSubnet.grid(row=4, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_DefGateway.grid(row=5, column=0, pady=5, padx=5, sticky='w', columnspan=2)

            entry_MAC.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_IPAddress.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_IPSubnet.grid(row=4, column=1, pady=5, padx=5, sticky='w')
            entry_DefGateway.grid(row=5, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=8, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=8, column=1, pady=5, padx=5, sticky='w')

            return

    networkWindow = tk.Toplevel(siwAssetsWindow)
    siwAssetsWindow.withdraw()
    networkWindow.geometry("600x300")
    networkWindow.title("Network Information")
    networkWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    networkWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(networkWindow, width=15, text="Home", bg='white',
                         command=lambda: home(networkWindow, homeWindow))
    btn_back = tk.Button(networkWindow, width=15, text="Back", bg='white',
                         command=lambda: home(networkWindow, siwAssetsWindow))
    btn_edit = tk.Button(networkWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(networkWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry(SIWAsset_ID))
    btn_delete = tk.Button(networkWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(networkWindow, column=('#1', '#2', '#3', '#4', '#5'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='MAC Address')
    tree.heading('#3', text='IP Address')
    tree.heading('#4', text='IP Subnet')
    tree.heading('#5', text='Default Gateway')
    tree.column('#1', width=30)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
    tbl_MACs.MACs_ID,
    tbl_MACs.MACAddress,
    tbl_MACs.IPAddress,
    tbl_MACs.IPSubnet,
    tbl_MACs.DefaultGateway
    FROM tbl_MACs
    WHERE tbl_MACs.SIWAsset_ID = """ + str(SIWAsset_ID)
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1], record[2], record[3], record[4]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def viewOpenPorts(SIWAsset_ID):
    def updateWidgets(event):
        try:
            if entry_Port.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                OpenPortInfo = Selection.get('values')
                Port_Var.set(OpenPortInfo[1])
                State_Var.set(OpenPortInfo[2])
                Cmb_Protocol.set(OpenPortInfo[3])
        except:
            return
        return

    def deleteWidgets():
        try:
            if entry_Port.winfo_exists():
                entry_Port.grid_forget()
                entry_State.grid_forget()
                Cmb_Protocol.grid_forget()
                Btn_Save.grid_forget()
                Btn_Cancel.grid_forget()
        except:
            return
        return

    def cancel():
        entry_Port.grid_forget()
        entry_State.grid_forget()
        Cmb_Protocol.grid_forget()
        Btn_Cancel.grid_forget()
        Btn_Save.grid_forget()
        return

    def edit():
        global entry_Port
        global entry_State
        global Cmb_Protocol
        global Port_Var
        global State_Var
        global Btn_Cancel
        global Btn_Save
        global pList

        def save():
            curItem = tree.focus()
            Selection = tree.item(curItem)
            OpenPortInfo = Selection.get('values')

            if (str(entry_Port.get()) == str(OpenPortInfo[1])) and (
                    str(entry_State.get()) == str(OpenPortInfo[2])) and (
                    str(Cmb_Protocol.get()) == str(OpenPortInfo[3])):
                messagebox.showinfo("Error", "Nothing to save")
            else:
                count = 1
                protocol_ID = 'NULL'
                for protocol in pList:
                    if Cmb_Protocol.get() == protocol:
                        protocol_ID = count
                    count += 1

                SQL = """UPDATE tbl_OpenPorts
                        SET Port = '{0}', State = '{1}', Protocol_ID = {2}
                        WHERE OpenPorts_ID = {3}""".format(entry_Port.get(), entry_State.get(), protocol_ID,
                                                           OpenPortInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                OpenPortsWindow.destroy()
                viewOpenPorts(SIWAsset_ID)

            return

        # get users current selection in treeview
        curItem = tree.focus()
        Selection = tree.item(curItem)
        OpenPortInfo = Selection.get('values')
        # Create Entry Widgets/ Combobox
        Port_Var = tk.StringVar()
        State_Var = tk.StringVar()

        entry_Port = tk.Entry(OpenPortsWindow, textvariable=Port_Var, width=15)
        entry_State = tk.Entry(OpenPortsWindow, textvariable=State_Var, width=15)
        Cmb_Protocol = ttk.Combobox(OpenPortsWindow, width=15)
        Btn_Cancel = tk.Button(OpenPortsWindow, width=15, text="Cancel", bg='white', command=cancel)
        Btn_Save = tk.Button(OpenPortsWindow, width=15, text="Save", bg='white', command=save)

        # Place widgets in window
        entry_Port.grid(row=2, column=0, pady=5, padx=5, sticky='w')
        entry_State.grid(row=3, column=0, pady=5, padx=5, sticky='w')
        Cmb_Protocol.grid(row=4, column=0, pady=5, padx=5, sticky='w')
        Btn_Cancel.grid(row=5, column=0, pady=5, padx=5, sticky='w')
        Btn_Save.grid(row=5, column=1, pady=5, padx=5, sticky='w')
        # Set the first two entry widgets values
        Port_Var.set(OpenPortInfo[1])
        State_Var.set(OpenPortInfo[2])

        # Execute query and use results to fill Combobox and set the combobox value
        conn = pyodbc.connect(cnxn)
        cursor = conn.cursor()
        SQL = """SELECT
        tbl_Protocols.Name
        FROM tbl_Protocols """
        cursor.execute(SQL)
        protocols = cursor.fetchall()
        pList = []
        for protocol in protocols:
            pList.append(protocol[0])
        Cmb_Protocol["values"] = pList
        Cmb_Protocol.set(OpenPortInfo[3])

        return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        OpenPortInfo = Selection.get('values')
        SQL = """UPDATE tbl_OpenPorts
                    SET Deleted = 1
                    WHERE OpenPorts_ID = {0}""".format(OpenPortInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry(SIWAsset_ID):

        def save(SIWAsset_ID):
            protocol_ID = 'NULL'
            port = 'NULL'
            state = 'NULL'
            count = 1
            for protocol in pList:
                if Cmb_Protocol.get() == protocol:
                    protocol_ID = count
                count += 1
            if entry_Port.get() != '':
                port = entry_Port.get()
            if entry_State.get() != '':
                state = entry_State.get()
            SQL = """INSERT INTO tbl_OpenPorts (SIWAsset_ID, Protocol_ID, Port, State)
                        VALUES ({0},{1},'{2}','{3}')""".format(SIWAsset_ID, protocol_ID, port, state)
            cursor.execute(SQL)
            cursor.commit()
            OpenPortsWindow.destroy()
            viewOpenPorts(SIWAsset_ID)
            return

        deleteWidgets()
        global entry_Port
        global entry_State
        global Cmb_Protocol
        global Port_Var
        global State_Var
        global Btn_Cancel
        global Btn_Save
        entry_Port = tk.Entry(OpenPortsWindow, width=15)
        entry_State = tk.Entry(OpenPortsWindow, width=15)
        Cmb_Protocol = ttk.Combobox(OpenPortsWindow, width=15)
        Btn_Cancel = tk.Button(OpenPortsWindow, width=15, text="Cancel", bg='white', command=cancel)
        Btn_Save = tk.Button(OpenPortsWindow, width=15, text="Save", bg='white', command=lambda: save(SIWAsset_ID))

        # Place widgets in window
        entry_Port.grid(row=2, column=0, pady=5, padx=5, sticky='w')
        entry_State.grid(row=3, column=0, pady=5, padx=5, sticky='w')
        Cmb_Protocol.grid(row=4, column=0, pady=5, padx=5, sticky='w')
        Btn_Cancel.grid(row=5, column=0, pady=5, padx=5, sticky='w')
        Btn_Save.grid(row=5, column=1, pady=5, padx=5, sticky='w')
        cursor = conn.cursor()
        SQL = """SELECT
                tbl_Protocols.Name
                FROM tbl_Protocols """
        cursor.execute(SQL)
        protocols = cursor.fetchall()
        pList = []
        for protocol in protocols:
            pList.append(protocol[0])
        Cmb_Protocol["values"] = pList

        return

    OpenPortsWindow = tk.Toplevel(siwAssetsWindow)
    siwAssetsWindow.withdraw()
    OpenPortsWindow.geometry("600x300")
    OpenPortsWindow.title("Open Ports")
    OpenPortsWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    OpenPortsWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(OpenPortsWindow, width=15, text="Home", bg='white',
                         command=lambda: home(OpenPortsWindow, homeWindow))
    btn_back = tk.Button(OpenPortsWindow, width=15, text="Back", bg='white',
                         command=lambda: home(OpenPortsWindow, siwAssetsWindow))
    btn_edit = tk.Button(OpenPortsWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(OpenPortsWindow, width=15, text="New Entry", bg='white', comman=lambda: newEntry(SIWAsset_ID))
    btn_delete = tk.Button(OpenPortsWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(OpenPortsWindow, column=('#1', '#2', '#3', '#4'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Port')
    tree.heading('#3', text='State')
    tree.heading('#4', text='Protocol')

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
tbl_OpenPorts.Port,
tbl_OpenPorts.State,
tbl_Protocols.Name,
tbl_OpenPorts.OpenPorts_ID

FROM tbl_OpenPorts
INNER JOIN tbl_Protocols ON tbl_OpenPorts.Protocol_ID = tbl_Protocols.Protocol_ID

WHERE tbl_OpenPorts.Deleted IS NULL AND tbl_OpenPorts.SIWAsset_ID = """ + str(SIWAsset_ID)
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[3], record[0], record[1], record[2]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)


def viewSoftware(SIWAsset_ID):
    global NetworkWindow

    def updateWidgets(event):
        try:
            if entry_SoftwareName.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                SoftwareInfo = Selection.get('values')
                Software_Var.set(SoftwareInfo[1])
                Corp_Var.set(SoftwareInfo[2])
                Date_Var.set(SoftwareInfo[3])
                Version_Var.set(SoftwareInfo[4])
                Cmb_CVEs.set(SoftwareInfo[5])
        except:
            return
        return

    def deleteWidgets():
        try:
            if entry_SoftwareName.winfo_exists():
                Frame1.grid_forget()
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global entry_SoftwareName
            global entry_Corp
            global entry_InstallDate
            global entry_Version

            global Cmb_CVEs
            global Cmb_Source

            global Software_Var
            global Corp_Var
            global Date_Var
            global Version_Var

            global Btn_Cancel
            global Btn_Save
            global CVEList

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                SoftwareInfo = Selection.get('values')
                # test if selection has CVE
                if Cmb_CVEs.get() == "None":
                    CVE_ID = 'NULL'
                else:
                    for record in CVEs:
                        if Cmb_CVEs.get() == record[0]:
                            CVE_ID = record[1]

                SQL = """UPDATE tbl_Software_Info
                        SET SoftwareName = '{0}', Corporation = '{1}', InstallDate = '{2}', SoftwareVersion = '{3}', CVE_ID = {4}
                        WHERE SoftwareInfo_ID = {5}""".format(entry_SoftwareName.get(), entry_Corp.get(),
                                                              entry_InstallDate.get(), entry_Version.get(), CVE_ID,
                                                              SoftwareInfo[0])
                cursor.execute(SQL)
                cursor.commit()

                if SoftwareInfo[6] != Cmb_Source.get():
                    for source in sources:
                        if source[0] == Cmb_Source.get():
                            Source_ID = source[1]
                    SQL = """UPDATE tbl_Software
                        SET Source_ID = '{0}'
                        WHERE SoftwareInfo_ID = {1}""".format(Source_ID, SoftwareInfo[0])
                    cursor.execute(SQL)
                    cursor.commit()
                NetworkWindow.destroy()
                viewSoftware(SIWAsset_ID)
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            SoftwareInfo = Selection.get('values')
            # Create Entry Widgets/ Combobox
            Software_Var = tk.StringVar()
            Corp_Var = tk.StringVar()
            Date_Var = tk.StringVar()
            Version_Var = tk.StringVar()

            Frame1 = tk.Frame(NetworkWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_SoftwareName = tk.Label(Frame1, text="Software Name", bg='white', anchor='w')
            lbl_Corporation = tk.Label(Frame1, text="Corporation", bg='white', anchor='w')
            lbl_InstallDate = tk.Label(Frame1, text="Install Date", bg='white', anchor='w')
            lbl_Version = tk.Label(Frame1, text="Version", bg='white', anchor='w')
            lbl_CVE = tk.Label(Frame1, text="CVE", bg='white', anchor='w')
            lbl_Source = tk.Label(Frame1, text="Source", bg='white', anchor='w')

            entry_SoftwareName = tk.Entry(Frame1, textvariable=Software_Var, width=40)
            entry_Corp = tk.Entry(Frame1, textvariable=Corp_Var, width=30)
            entry_InstallDate = tk.Entry(Frame1, textvariable=Date_Var, width=15)
            entry_Version = tk.Entry(Frame1, textvariable=Version_Var, width=15)
            Cmb_CVEs = ttk.Combobox(Frame1, width=15)
            Cmb_Source = ttk.Combobox(Frame1, width=15)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_SoftwareName.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Corporation.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_InstallDate.grid(row=4, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Version.grid(row=5, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_CVE.grid(row=6, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Source.grid(row=7, column=0, pady=5, padx=5, sticky='w', columnspan=2)

            entry_SoftwareName.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_Corp.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_InstallDate.grid(row=4, column=1, pady=5, padx=5, sticky='w')
            entry_Version.grid(row=5, column=1, pady=5, padx=5, sticky='w')
            Cmb_CVEs.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            Cmb_Source.grid(row=7, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=8, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=8, column=1, pady=5, padx=5, sticky='w')
            # Set the first two entry widgets values
            Software_Var.set(SoftwareInfo[1])
            Corp_Var.set(SoftwareInfo[2])
            Date_Var.set(SoftwareInfo[3])
            Version_Var.set(SoftwareInfo[4])

            # Execute query and use results to fill CVEs Combobox and set the combobox value
            conn = pyodbc.connect(cnxn)
            cursor = conn.cursor()
            SQL = """SELECT
            tbl_CVEs.Name, tbl_CVEs.CVE_ID
            FROM tbl_CVEs
            ORDER BY tbl_CVEs.Name ASC"""
            cursor.execute(SQL)
            CVEs = cursor.fetchall()
            CVEList = []
            for CVE in CVEs:
                CVEList.append(CVE[0])
            Cmb_CVEs["values"] = CVEList
            Cmb_CVEs.set(SoftwareInfo[5])

            # Execute query and use results to fill Sources Combobox and set the combobox value
            SQL = """SELECT tbl_SoftwareSources.Name, tbl_SoftwareSources.Source_ID
            FROM tbl_SoftwareSources"""
            cursor.execute(SQL)
            sources = cursor.fetchall()
            SourceList = []
            for source in sources:
                SourceList.append(source[0])
            Cmb_Source["values"] = SourceList
            Cmb_Source.set(SoftwareInfo[6])
            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        SoftwareInfo = Selection.get('values')
        SQL = """UPDATE tbl_Software
                    SET Deleted = 1
                    WHERE SoftwareInfo_ID = {0}""".format(SoftwareInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry(SIWAsset_ID):

        def save(SIWAsset_ID):
            def getSoftwareInfoID():

                SoftwareInfo_ID = ""

                SQL = """SELECT tbl_Software_Info.SoftwareInfo_ID
                FROM tbl_Software_Info
                WHERE tbl_Software_Info.CVE_ID = {0} AND
                tbl_Software_Info.SoftwareName = '{1}' AND
                tbl_Software_Info.Corporation = '{2}' AND
                tbl_Software_Info.InstallDate = '{3}' AND
                tbl_Software_Info.SoftwareVersion = '{4}'""".format(CVE_ID, SoftwareName, Corporation, InstallDate,
                                                                    Version)
                cursor.execute(SQL)
                SoftwareInfo_ID = cursor.fetchall()

                return SoftwareInfo_ID

            SoftwareName = 'NULL'
            Corporation = 'NULL'
            InstallDate = 'NULL'
            Version = 'NULL'
            CVE_ID = 'NULL'
            Source_ID = 'NULL'

            # Set values for
            for cve in CVEList:
                if Cmb_CVEs.get() == cve[0]:
                    CVE_ID = cve[1]
            for source in sources:
                if Cmb_Source.get() == source[0]:
                    Source_ID = source[1]

            if entry_SoftwareName.get() != '':
                SoftwareName = entry_SoftwareName.get()
            if entry_Corp.get() != '':
                Corporation = entry_Corp.get()
            if entry_InstallDate.get() != '':
                InstallDate = entry_InstallDate.get()
            if entry_Version.get() != '':
                Version = entry_Version.get()
            SQL = """SELECT tbl_Software_Info.SoftwareInfo_ID
                        FROM tbl_Software_Info"""
            cursor.execute(SQL)
            records = cursor.fetchall()

            lastInt = records[-1][0]
            SQL = """INSERT INTO tbl_Software_Info (CVE_ID, SoftwareName, Corporation, InstallDate, SoftwareVersion)
                        VALUES ({0},'{1}','{2}','{3}', '{4}')""".format(CVE_ID, SoftwareName, Corporation, InstallDate,
                                                                        Version)
            cursor.execute(SQL)
            cursor.commit()

            SQL = """INSERT INTO tbl_Software (SIWAsset_ID, SoftwareInfo_ID, Source_ID)
                        VALUES ({0}, {1}, {2})""".format(SIWAsset_ID, lastInt, Source_ID)
            cursor.execute(SQL)
            cursor.commit()
            NetworkWindow.destroy()
            viewSoftware(SIWAsset_ID)
            return

        deleteWidgets()
        global entry_SoftwareName
        global entry_Corp
        global entry_InstallDate
        global entry_Version
        global Cmb_CVEs
        global Cmb_Source
        global Frame1
        global Software_Var
        global Corp_Var
        global Btn_Cancel
        global Btn_Save

        Frame1 = tk.Frame(NetworkWindow, padx=5, pady=5, bg='white')
        Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

        lbl_SoftwareName = tk.Label(Frame1, text="Software Name", bg='white', anchor='w')
        lbl_Corporation = tk.Label(Frame1, text="Corporation", bg='white', anchor='w')
        lbl_InstallDate = tk.Label(Frame1, text="Install Date", bg='white', anchor='w')
        lbl_Version = tk.Label(Frame1, text="Version", bg='white', anchor='w')
        lbl_CVE = tk.Label(Frame1, text="CVE", bg='white', anchor='w')
        lbl_Source = tk.Label(Frame1, text="Source", bg='white', anchor='w')

        entry_SoftwareName = tk.Entry(Frame1, width=40)
        entry_Corp = tk.Entry(Frame1, width=30)
        entry_InstallDate = tk.Entry(Frame1, width=15)
        entry_Version = tk.Entry(Frame1, width=15)
        Cmb_CVEs = ttk.Combobox(Frame1, width=15)
        Cmb_Source = ttk.Combobox(Frame1, width=15)

        Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
        Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=lambda: save(SIWAsset_ID))
        Btn_CVEs = tk.Button(Frame1, width=15, text="Manage CVEs", bg='white', command=save)
        Btn_Sources = tk.Button(Frame1, width=15, text="Manage Sources", bg='white', command=save)

        # Place widgets in window
        lbl_SoftwareName.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
        lbl_Corporation.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
        lbl_InstallDate.grid(row=4, column=0, pady=5, padx=5, sticky='w', columnspan=2)
        lbl_Version.grid(row=5, column=0, pady=5, padx=5, sticky='w', columnspan=2)
        lbl_CVE.grid(row=6, column=0, pady=5, padx=5, sticky='w', columnspan=2)
        lbl_Source.grid(row=7, column=0, pady=5, padx=5, sticky='w', columnspan=2)

        entry_SoftwareName.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
        entry_Corp.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
        entry_InstallDate.grid(row=4, column=1, pady=5, padx=5, sticky='w')
        entry_Version.grid(row=5, column=1, pady=5, padx=5, sticky='w')
        Cmb_CVEs.grid(row=6, column=1, pady=5, padx=5, sticky='w')
        Cmb_Source.grid(row=7, column=1, pady=5, padx=5, sticky='w')

        Btn_CVEs.grid(row=6, column=2, pady=5, padx=5, sticky='w')
        Btn_Sources.grid(row=7, column=2, pady=5, padx=5, sticky='w')
        Btn_Cancel.grid(row=8, column=0, pady=5, padx=5, sticky='w')
        Btn_Save.grid(row=8, column=1, pady=5, padx=5, sticky='w')

        cursor = conn.cursor()
        SQL = """SELECT
                tbl_CVEs.Name, tbl_CVEs.CVE_ID
                FROM tbl_CVEs """
        cursor.execute(SQL)
        CVEs = cursor.fetchall()
        CVEList = []
        for CVE in CVEs:
            CVEList.append(CVE[0])
        Cmb_CVEs["values"] = CVEList

        SQL = """SELECT tbl_SoftwareSources.Name, tbl_SoftwareSources.Source_ID
                    FROM tbl_SoftwareSources"""
        cursor.execute(SQL)
        sources = cursor.fetchall()
        SourceList = []
        for source in sources:
            SourceList.append(source[0])
        Cmb_Source["values"] = SourceList

        return

    NetworkWindow = tk.Toplevel(siwAssetsWindow)
    siwAssetsWindow.withdraw()
    NetworkWindow.geometry("600x300")
    NetworkWindow.title("Software")
    NetworkWindow.configure(bg='gray15')
    window_width = 1200
    window_height = 700
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    NetworkWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(NetworkWindow, width=15, text="Home", bg='white',
                         command=lambda: home(NetworkWindow, homeWindow))
    btn_back = tk.Button(NetworkWindow, width=15, text="Back", bg='white',
                         command=lambda: home(NetworkWindow, siwAssetsWindow))
    btn_edit = tk.Button(NetworkWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(NetworkWindow, width=15, text="New Entry", bg='white', comman=lambda: newEntry(SIWAsset_ID))
    btn_delete = tk.Button(NetworkWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')
    tree = ttk.Treeview(NetworkWindow, column=('#1', '#2', '#3', '#4', '#5', '#6', '#7'), show='headings')
    tree.column('#1', width=30)
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Name')
    tree.heading('#3', text='Coprporation')
    tree.heading('#4', text='Install Date')
    tree.heading('#5', text='Version')
    tree.heading('#6', text='CVE')
    tree.heading('#7', text='Source')

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT DISTINCT
                tbl_Software_Info.SoftwareInfo_ID,
                tbl_Software_Info.SoftwareName,
                tbl_Software_Info.Corporation,
                tbl_Software_Info.InstallDate,
                tbl_Software_Info.SoftwareVersion,
                tbl_CVEs.Name,
                tbl_SoftwareSources.Name

                FROM tbl_Software

                FULL JOIN tbl_Software_Info ON tbl_Software.SoftwareInfo_ID = tbl_Software_Info.SoftwareInfo_ID
                FULL JOIN tbl_CVEs ON tbl_Software_Info.CVE_ID = tbl_CVEs.CVE_ID
                FULL JOIN tbl_SoftwareSources ON tbl_Software.Source_ID = tbl_SoftwareSources.Source_ID


                WHERE tbl_Software.SIWAsset_ID = {0} AND tbl_Software.Deleted IS NULL""".format(SIWAsset_ID)
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count,
                        values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)


def viewUpdates(SIWAsset_ID):
    def updateWidgets(event):
        try:
            if entry_UpdateName.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                UpdateInfo = Selection.get('values')
                Update_Var.set(UpdateInfo[1])
                Hotfix_Var.set(UpdateInfo[2])
                Date_Var.set(UpdateInfo[3])
                Cmb_Category.set(UpdateInfo[4])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global entry_UpdateName
            global entry_Hotfix
            global entry_InstallDate

            global Cmb_Category

            global Update_Var
            global Hotfix_Var
            global Date_Var

            global Btn_Cancel
            global Btn_Save
            global CVEList

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                UpdateInfo = Selection.get('values')

                Name = 'NULL'
                Hotfix = 'NULL'
                Date = 'NULL'

                if entry_UpdateName.get() != "":
                    Name = entry_UpdateName.get()
                if entry_Hotfix.get() != "":
                    Hotfix = entry_Hotfix.get()
                if entry_InstallDate.get() != "":
                    Date = entry_InstallDate.get()

                for record in Categories:
                    if Cmb_Category.get() == record[0]:
                        Category_ID = record[1]

                SQL = """UPDATE tbl_Updates
                        SET Category_ID = {0}, Name = '{1}', Hotfix = '{2}', Date = '{3}'
                        WHERE Updates_ID = {4}""".format(Category_ID, Name, Hotfix, Date, UpdateInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                UpdatesWindow.destroy()
                viewUpdates(SIWAsset_ID)
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            UpdateInfo = Selection.get('values')
            # Create Entry Widgets/ Combobox
            Update_Var = tk.StringVar()
            Hotfix_Var = tk.StringVar()
            Date_Var = tk.StringVar()

            Frame1 = tk.Frame(UpdatesWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_UpdateName = tk.Label(Frame1, text="Name", bg='white', anchor='w')
            lbl_Hotfix = tk.Label(Frame1, text="Hotfix", bg='white', anchor='w')
            lbl_InstallDate = tk.Label(Frame1, text="Install Date", bg='white', anchor='w')
            lbl_Category = tk.Label(Frame1, text="Category", bg='white', anchor='w')

            entry_UpdateName = tk.Entry(Frame1, textvariable=Update_Var, width=40)
            entry_Hotfix = tk.Entry(Frame1, textvariable=Hotfix_Var, width=30)
            entry_InstallDate = tk.Entry(Frame1, textvariable=Date_Var, width=15)
            Cmb_Category = ttk.Combobox(Frame1, width=20)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_UpdateName.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Hotfix.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_InstallDate.grid(row=4, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Category.grid(row=5, column=0, pady=5, padx=5, sticky='w', columnspan=2)

            entry_UpdateName.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_Hotfix.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_InstallDate.grid(row=4, column=1, pady=5, padx=5, sticky='w')
            Cmb_Category.grid(row=5, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            # Set the first two entry widgets values
            Update_Var.set(UpdateInfo[1])
            Hotfix_Var.set(UpdateInfo[2])
            Date_Var.set(UpdateInfo[3])

            # Execute query and use results to fill CVEs Combobox and set the combobox value
            conn = pyodbc.connect(cnxn)
            cursor = conn.cursor()
            SQL = """SELECT
            tbl_UpdatesCategories.Name, tbl_UpdatesCategories.Category_ID
            FROM tbl_UpdatesCategories"""
            cursor.execute(SQL)
            Categories = cursor.fetchall()
            CategoryList = []
            for Category in Categories:
                CategoryList.append(Category[0])
            Cmb_Category["values"] = CategoryList
            Cmb_Category.set(UpdateInfo[4])
            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        UpdateInfo = Selection.get('values')
        SQL = """UPDATE tbl_Updates
                    SET Deleted = 1
                    WHERE Updates_ID = {0}""".format(UpdateInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry(SIWAsset_ID):
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global entry_UpdateName
            global entry_Hotfix
            global entry_InstallDate

            global Cmb_Category

            global Software_Var
            global Corp_Var
            global Date_Var
            global Version_Var

            global Btn_Cancel
            global Btn_Save
            global CVEList

            def save(SIWAsset_ID):
                curItem = tree.focus()
                Selection = tree.item(curItem)
                UpdateInfo = Selection.get('values')

                Name = 'NULL'
                Hotfix = 'NULL'
                Date = 'NULL'

                if entry_UpdateName.get() != "":
                    Name = entry_UpdateName.get()
                if entry_Hotfix.get() != "":
                    Hotfix = entry_Hotfix.get()
                if entry_InstallDate.get() != "":
                    Date = entry_InstallDate.get()

                for record in Categories:
                    if Cmb_Category.get() == record[0]:
                        Category_ID = record[1]

                SQL = """INSERT INTO tbl_Updates (SIWAsset_ID, Category_ID, Name, Hotfix, Date)
                            VALUES ({0},{1},'{2}','{3}','{4}')""".format(SIWAsset_ID, Category_ID, Name, Hotfix, Date)
                cursor.execute(SQL)
                cursor.commit()
                UpdatesWindow.destroy()
                viewUpdates(SIWAsset_ID)
                return

            Frame1 = tk.Frame(UpdatesWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_UpdateName = tk.Label(Frame1, text="Name", bg='white', anchor='w')
            lbl_Hotfix = tk.Label(Frame1, text="Hotfix", bg='white', anchor='w')
            lbl_InstallDate = tk.Label(Frame1, text="Install Date", bg='white', anchor='w')
            lbl_Category = tk.Label(Frame1, text="Category", bg='white', anchor='w')

            entry_UpdateName = tk.Entry(Frame1, width=40)
            entry_Hotfix = tk.Entry(Frame1, width=30)
            entry_InstallDate = tk.Entry(Frame1, width=15)
            Cmb_Category = ttk.Combobox(Frame1, width=20)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=lambda: save(SIWAsset_ID))

            # Place widgets in window
            lbl_UpdateName.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Hotfix.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_InstallDate.grid(row=4, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Category.grid(row=5, column=0, pady=5, padx=5, sticky='w', columnspan=2)

            entry_UpdateName.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_Hotfix.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_InstallDate.grid(row=4, column=1, pady=5, padx=5, sticky='w')
            Cmb_Category.grid(row=5, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            # Execute query and use results to fill CVEs Combobox and set the combobox value
            conn = pyodbc.connect(cnxn)
            cursor = conn.cursor()
            SQL = """SELECT
            tbl_UpdatesCategories.Name, tbl_UpdatesCategories.Category_ID
            FROM tbl_UpdatesCategories"""
            cursor.execute(SQL)
            Categories = cursor.fetchall()

            CategoryList = []
            for Category in Categories:
                CategoryList.append(Category[0])
            Cmb_Category["values"] = CategoryList

            return

    UpdatesWindow = tk.Toplevel(siwAssetsWindow)
    siwAssetsWindow.withdraw()
    UpdatesWindow.geometry("600x300")
    UpdatesWindow.title("Installed Updates")
    UpdatesWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    UpdatesWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(UpdatesWindow, width=15, text="Home", bg='white',
                         command=lambda: home(UpdatesWindow, homeWindow))
    btn_back = tk.Button(UpdatesWindow, width=15, text="Back", bg='white',
                         command=lambda: home(UpdatesWindow, siwAssetsWindow))
    btn_edit = tk.Button(UpdatesWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(UpdatesWindow, width=15, text="New Entry", bg='white', comman=lambda: newEntry(SIWAsset_ID))
    btn_delete = tk.Button(UpdatesWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(UpdatesWindow, column=('#1', '#2', '#3', '#4', '#5'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Name')
    tree.heading('#3', text='Hotfix')
    tree.heading('#4', text='Date')
    tree.heading('#5', text='Type')
    tree.column('#1', width=30)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_Updates.Updates_ID,
                tbl_Updates.Name,
                tbl_Updates.Hotfix,
                tbl_Updates.Date,
                tbl_UpdatesCategories.Name

                FROM tbl_Updates
                INNER JOIN tbl_UpdatesCategories ON tbl_Updates.Category_ID = tbl_UpdatesCategories.Category_ID

                WHERE tbl_Updates.SIWAsset_ID = """ + str(SIWAsset_ID)
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1], record[2], record[3], record[4]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def viewUSBHistory(SIWAsset_ID):
    def updateWidgets(event):
        try:
            if entry_UpdateName.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                UpdateInfo = Selection.get('values')
                Update_Var.set(UpdateInfo[1])
                Hotfix_Var.set(UpdateInfo[2])
                Date_Var.set(UpdateInfo[3])
                Cmb_Category.set(UpdateInfo[4])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global entry_Device
            global entry_Diskstamp

            global Cmb_Class

            global Device_Var
            global Stamp_Var

            global Btn_Cancel
            global Btn_Save
            global ClassList

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                USBInfo = Selection.get('values')

                Device = 'NULL'
                DiskStamp = 'NULL'
                Class_ID = 'NULL'

                if entry_Device.get() != "":
                    Device = entry_Device.get()
                if entry_Diskstamp.get() != "":
                    DiskStamp = entry_Diskstamp.get()

                for record in Classes:
                    if Cmb_Class.get() == record[0]:
                        Class_ID = record[1]

                SQL = """UPDATE tbl_USBHistory
                        SET Device = '{0}', DiskStamp = '{1}', Class_ID = {2}
                        WHERE USBHistory_ID = {3}""".format(Device, DiskStamp, Class_ID, USBInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                USBWindow.destroy()
                viewUSBHistory(SIWAsset_ID)
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            USBInfo = Selection.get('values')
            # Create Entry Widgets/ Combobox
            Device_Var = tk.StringVar()
            Stamp_Var = tk.StringVar()

            Frame1 = tk.Frame(USBWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Device = tk.Label(Frame1, text="Name", bg='white', anchor='w')
            lbl_Diskstamp = tk.Label(Frame1, text="Stamp", bg='white', anchor='w')
            lbl_Class = tk.Label(Frame1, text="Class", bg='white', anchor='w')

            entry_Device = tk.Entry(Frame1, textvariable=Device_Var, width=40)
            entry_Diskstamp = tk.Entry(Frame1, textvariable=Stamp_Var, width=30)
            Cmb_Class = ttk.Combobox(Frame1, width=20)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Device.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Diskstamp.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Class.grid(row=4, column=0, pady=5, padx=5, sticky='w', columnspan=2)

            entry_Device.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_Diskstamp.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            Cmb_Class.grid(row=4, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            # Set the first two entry widgets values
            Device_Var.set(USBInfo[1])
            Stamp_Var.set(USBInfo[2])

            # Execute query and use results to fill CVEs Combobox and set the combobox value
            conn = pyodbc.connect(cnxn)
            cursor = conn.cursor()
            SQL = """SELECT
            tbl_USBClasses.Name, tbl_USBClasses.Class_ID
            FROM tbl_USBClasses"""
            cursor.execute(SQL)
            Classes = cursor.fetchall()
            ClassList = []
            for Class in Classes:
                ClassList.append(Class[0])
            Cmb_Class["values"] = ClassList
            Cmb_Class.set(USBInfo[3])
            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        USBInfo = Selection.get('values')
        SQL = """UPDATE tbl_USBHistory
                    SET Deleted = 1
                    WHERE USBHistory_ID = {0}""".format(USBInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry(SIWAsset_ID):
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global entry_Device
            global entry_Diskstamp

            global Cmb_Class

            global Device_Var
            global Stamp_Var

            global Btn_Cancel
            global Btn_Save
            global ClassList

            def save(SIWAsset_ID):
                curItem = tree.focus()
                Selection = tree.item(curItem)
                USBInfo = Selection.get('values')

                Device = 'NULL'
                DiskStamp = 'NULL'
                Class_ID = 'NULL'

                if entry_Device.get() != "":
                    Device = entry_Device.get()
                if entry_Diskstamp.get() != "":
                    DiskStamp = entry_Diskstamp.get()

                for record in Classes:
                    if Cmb_Class.get() == record[0]:
                        Class_ID = record[1]

                SQL = """INSERT INTO tbl_USBHistory (SIWAsset_ID, Device, DiskStamp, Class_ID)
                                VALUES ({0},'{1}','{2}',{3})""".format(SIWAsset_ID, Device, DiskStamp, Class_ID)
                cursor.execute(SQL)
                cursor.commit()
                USBWindow.destroy()
                viewUSBHistory(SIWAsset_ID)
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            USBInfo = Selection.get('values')
            # Create Entry Widgets/ Combobox
            Device_Var = tk.StringVar()
            Stamp_Var = tk.StringVar()

            Frame1 = tk.Frame(USBWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Device = tk.Label(Frame1, text="Name", bg='white', anchor='w')
            lbl_Diskstamp = tk.Label(Frame1, text="Stamp", bg='white', anchor='w')
            lbl_Class = tk.Label(Frame1, text="Class", bg='white', anchor='w')

            entry_Device = tk.Entry(Frame1, width=40)
            entry_Diskstamp = tk.Entry(Frame1, width=30)
            Cmb_Class = ttk.Combobox(Frame1, width=20)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=lambda: save(SIWAsset_ID))

            # Place widgets in window
            lbl_Device.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Diskstamp.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Class.grid(row=4, column=0, pady=5, padx=5, sticky='w', columnspan=2)

            entry_Device.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_Diskstamp.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            Cmb_Class.grid(row=4, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            # Execute query and use results to fill CVEs Combobox and set the combobox value
            conn = pyodbc.connect(cnxn)
            cursor = conn.cursor()
            SQL = """SELECT
            tbl_USBClasses.Name, tbl_USBClasses.Class_ID
            FROM tbl_USBClasses"""
            cursor.execute(SQL)
            Classes = cursor.fetchall()
            ClassList = []
            for Class in Classes:
                ClassList.append(Class[0])
            Cmb_Class["values"] = ClassList

            return

    USBWindow = tk.Toplevel(siwAssetsWindow)
    siwAssetsWindow.withdraw()
    USBWindow.geometry("600x300")
    USBWindow.title("Recent USB History")
    USBWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    USBWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(USBWindow, width=15, text="Home", bg='white',
                         command=lambda: home(USBWindow, homeWindow))
    btn_back = tk.Button(USBWindow, width=15, text="Back", bg='white',
                         command=lambda: home(USBWindow, siwAssetsWindow))
    btn_edit = tk.Button(USBWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(USBWindow, width=15, text="New Entry", bg='white', comman=lambda: newEntry(SIWAsset_ID))
    btn_delete = tk.Button(USBWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(USBWindow, column=('#1', '#2', '#3', '#4'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Device')
    tree.heading('#3', text='Disk Stamp')
    tree.heading('#4', text='Class')
    tree.column('#1', width=30)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_USBHistory.USBHistory_ID,
                tbl_USBHistory.Device,
                tbl_USBHistory.DiskStamp,
                tbl_USBClasses.Name
                FROM tbl_USBHistory
                INNER JOIN tbl_USBClasses ON tbl_USBHistory.Class_ID = tbl_USBClasses.Class_ID
                WHERE tbl_USBHistory.Deleted IS NULL AND tbl_USBHistory.SIWAsset_ID = """ + str(SIWAsset_ID)
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1], record[2], record[3]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def viewVulnerabilities(SIWAsset_ID):
    def updateWidgets(event):
        try:
            if entry_UpdateName.winfo_exists():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                UpdateInfo = Selection.get('values')
                Update_Var.set(UpdateInfo[1])
                Hotfix_Var.set(UpdateInfo[2])
                Date_Var.set(UpdateInfo[3])
                Cmb_Category.set(UpdateInfo[4])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:

            global lbl_Vulnerability
            global Cmb_Vulnerability

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                VulnInfo = Selection.get('values')

                VulnName_ID = 'NULL'

                for record in VulnNames:
                    if Cmb_Vulnerability.get() == record[0]:
                        VulnName_ID = record[1]

                SQL = """UPDATE tbl_Vulnerabilities
                        SET VulnName_ID = {0}
                        WHERE Vulnerabilities_ID = {1}""".format(VulnName_ID, VulnInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                viewVulnerabilities(SIWAsset_ID)
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            VulnInfo = Selection.get('values')

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Vulnerability = tk.Label(Frame1, text="Vulnerability", bg='white', anchor='w')
            Cmb_Vulnerability = ttk.Combobox(Frame1, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_Vulnerability.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Cmb_Vulnerability.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            # Execute query and use results to fill CVEs Combobox and set the combobox value
            conn = pyodbc.connect(cnxn)
            cursor = conn.cursor()
            SQL = """SELECT
            tbl_VulnNames.Name, tbl_VulnNames.VulnName_ID
            FROM tbl_VulnNames"""
            cursor.execute(SQL)
            VulnNames = cursor.fetchall()
            VulnsList = []
            for Vuln in VulnNames:
                VulnsList.append(Vuln[0])
            Cmb_Vulnerability["values"] = VulnsList
            Cmb_Vulnerability.set(VulnInfo[1])
            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        VulnInfo = Selection.get('values')
        SQL = """UPDATE tbl_Vulnerabilities
                    SET Deleted = 1
                    WHERE tbl_Vulnerabilities.Vulnerabilities_ID = {0}""".format(VulnInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry(SIWAsset_ID):
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global lbl_Vulnerability
            global Cmb_Vulnerability

            global Btn_Cancel
            global Btn_Save
            global VulnsList

            def save(SIWAsset_ID):
                VulnName_ID = 'NULL'

                for record in VulnNames:
                    if Cmb_Vulnerability.get() == record[0]:
                        VulnName_ID = record[1]

                SQL = """INSERT INTO tbl_Vulnerabilities (SIWAsset_ID, VulnName_ID)
                                VALUES ({0},{1})""".format(SIWAsset_ID, VulnName_ID)
                cursor.execute(SQL)
                cursor.commit()
                VulnWindow.destroy()
                viewVulnerabilities(SIWAsset_ID)
                return

            # Create Entry Widgets/ Combobox
            Frame1 = tk.Frame(VulnWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_Vulnerability = tk.Label(Frame1, text="Vulnerability", bg='white', anchor='w')
            Cmb_Vulnerability = ttk.Combobox(Frame1, width=50)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=lambda: save(SIWAsset_ID))

            # Place widgets in window
            lbl_Vulnerability.grid(row=0, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            Cmb_Vulnerability.grid(row=0, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=6, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=6, column=1, pady=5, padx=5, sticky='w')

            # Execute query and use results to fill CVEs Combobox and set the combobox value
            conn = pyodbc.connect(cnxn)
            cursor = conn.cursor()
            SQL = """SELECT
                        tbl_VulnNames.Name, tbl_VulnNames.VulnName_ID
                        FROM tbl_VulnNames"""
            cursor.execute(SQL)
            VulnNames = cursor.fetchall()
            VulnsList = []
            for Vuln in VulnNames:
                VulnsList.append(Vuln[0])
            Cmb_Vulnerability["values"] = VulnsList
            return

    VulnWindow = tk.Toplevel(siwAssetsWindow)
    siwAssetsWindow.withdraw()
    VulnWindow.geometry("600x300")
    VulnWindow.title("Vulnerabilities")
    VulnWindow.configure(bg='gray15')
    window_width = 850
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    VulnWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(VulnWindow, width=15, text="Home", bg='white',
                         command=lambda: home(VulnWindow, homeWindow))
    btn_back = tk.Button(VulnWindow, width=15, text="Back", bg='white',
                         command=lambda: home(VulnWindow, siwAssetsWindow))
    btn_edit = tk.Button(VulnWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(VulnWindow, width=15, text="New Entry", bg='white', command=lambda: newEntry(SIWAsset_ID))
    btn_delete = tk.Button(VulnWindow, width=15, text="Delete Entry", bg='white', command=delete)

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(VulnWindow, column=('#1', '#2'), show='headings')
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Vulnerability')
    tree.column('#1', width=40)
    tree.column('#2', width=850)

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT
                tbl_Vulnerabilities.Vulnerabilities_ID,
                tbl_VulnNames.Name
                FROM tbl_Vulnerabilities
                INNER JOIN tbl_VulnNames ON tbl_Vulnerabilities.VulnName_ID = tbl_VulnNames.VulnName_ID
                WHERE tbl_Vulnerabilities.Deleted IS NULL AND tbl_Vulnerabilities.SIWAsset_ID = """ + str(SIWAsset_ID)
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count, values=(record[0], record[1]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)

    return


def OtherInfo(SIWAsset_ID):
    global SwitchPortsWindow
    global SwitchPorts_ID

    def updateWidgets(event):
        global SwitchPorts_ID
        curItem = tree.focus()
        Selection = tree.item(curItem)
        SoftwareInfo = Selection.get('values')
        SwitchPorts_ID = SoftwareInfo[0]
        try:
            if lbl_PatchCycle.winfo_exists():
                PatchCycle_Var.set(SoftwareInfo[1])
                PatchDays_Var.set(SoftwareInfo[2])
                MemoryLoad_Var.set(SoftwareInfo[3])
                USBDays_Var.set(SoftwareInfo[4])
                MS17010_Var.set(SoftwareInfo[5])
                MimiKatz_Var.set(SoftwareInfo[6])
                VulnerablePorts_Var.set(SoftwareInfo[7])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global lbl_PatchCycle
            global PatchCycle_Var
            global PatchDays_Var
            global MemoryLoad_Var
            global USBDays_Var
            global MS17010_Var
            global MimiKatz_Var
            global VulnerablePorts_Var

            global Btn_Cancel
            global Btn_Save

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                otherInfo = Selection.get('values')

                PatchCycle = entry_PatchCycle.get()
                PatchDays = entry_Patchdays.get()
                MemoryLoad = entry_MemoryLoad.get()
                USBDays = entry_USBDays.get()
                MS17010 = entry_MS17010.get()
                MimiKatz = entry_MimiKatz.get()
                VulnerablePorts = entry_VulnPorts.get()

                SQL = """UPDATE tbl_SIW_VUln_Info
                        SET PatchCycle = '{0}', PatchDays = '{1}', MemoryLoad = '{2}', USBDays = '{3}', MS17010 = '{4}', MimiKatz = '{5}', VulnerablePorts = '{6}'
                        WHERE VulnInfo_ID = {7}""".format(PatchCycle, PatchDays, MemoryLoad, USBDays, MS17010, MimiKatz,
                                                          VulnerablePorts, otherInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                OtherInfoWindow.destroy()
                OtherInfo(SIWAsset_ID)
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            OtherInfp = Selection.get('values')

            # Create Entry Widgets/ Combobox
            PatchCycle_Var = tk.StringVar()
            PatchDays_Var = tk.StringVar()
            MemoryLoad_Var = tk.StringVar()
            USBDays_Var = tk.StringVar()
            MS17010_Var = tk.StringVar()
            MimiKatz_Var = tk.StringVar()
            VulnerablePorts_Var = tk.StringVar()

            Frame1 = tk.Frame(OtherInfoWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_PatchCycle = tk.Label(Frame1, text="Patch Cycle", bg='white', anchor='w')
            lbl_Patchdays = tk.Label(Frame1, text="Patch days", bg='white', anchor='w')
            lbl_MemoryLoad = tk.Label(Frame1, text="Memory Load", bg='white', anchor='w')
            lbl_USBDays = tk.Label(Frame1, text="USB Days", bg='white', anchor='w')
            lbl_MS17010 = tk.Label(Frame1, text="MS17010", bg='white', anchor='w')
            lbl_MimiKatz = tk.Label(Frame1, text="MimiKatz", bg='white', anchor='w')
            lbl_VulnPorts = tk.Label(Frame1, text="Vulnerable Ports", bg='white', anchor='w')

            entry_PatchCycle = tk.Entry(Frame1, textvariable=PatchCycle_Var, width=40)
            entry_Patchdays = tk.Entry(Frame1, textvariable=PatchDays_Var, width=30)
            entry_MemoryLoad = tk.Entry(Frame1, textvariable=MemoryLoad_Var, width=15)
            entry_USBDays = tk.Entry(Frame1, textvariable=USBDays_Var, width=15)
            entry_MS17010 = tk.Entry(Frame1, textvariable=MS17010_Var, width=15)
            entry_MimiKatz = tk.Entry(Frame1, textvariable=MimiKatz_Var, width=15)
            entry_VulnPorts = tk.Entry(Frame1, textvariable=VulnerablePorts_Var, width=15)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_PatchCycle.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Patchdays.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_MemoryLoad.grid(row=4, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_USBDays.grid(row=5, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_MS17010.grid(row=6, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_MimiKatz.grid(row=7, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_VulnPorts.grid(row=8, column=0, pady=5, padx=5, sticky='w', columnspan=2)

            entry_PatchCycle.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_Patchdays.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_MemoryLoad.grid(row=4, column=1, pady=5, padx=5, sticky='w')
            entry_USBDays.grid(row=5, column=1, pady=5, padx=5, sticky='w')
            entry_MS17010.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            entry_MimiKatz.grid(row=7, column=1, pady=5, padx=5, sticky='w')
            entry_VulnPorts.grid(row=8, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=9, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=9, column=1, pady=5, padx=5, sticky='w')

            # Set the first two entry widgets values
            PatchCycle_Var.set(OtherInfp[1])
            PatchDays_Var.set(OtherInfp[2])
            MemoryLoad_Var.set(OtherInfp[3])
            USBDays_Var.set(OtherInfp[4])
            MS17010_Var.set(OtherInfp[5])
            MimiKatz_Var.set(OtherInfp[6])
            VulnerablePorts_Var.set(OtherInfp[7])

            return

    OtherInfoWindow = tk.Toplevel(siwAssetsWindow)
    siwAssetsWindow.withdraw()
    OtherInfoWindow.geometry("600x300")
    OtherInfoWindow.title("Switch Ports")
    OtherInfoWindow.configure(bg='gray15')
    window_width = 1200
    window_height = 600
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    OtherInfoWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(OtherInfoWindow, width=15, text="Home", bg='white',
                         command=lambda: home(OtherInfoWindow, homeWindow))
    btn_back = tk.Button(OtherInfoWindow, width=15, text="Back", bg='white',
                         command=lambda: home(OtherInfoWindow, siwAssetsWindow))
    btn_edit = tk.Button(OtherInfoWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(OtherInfoWindow, column=('#1', '#2', '#3', '#4', '#5', '#6', '#7', '#8'), show='headings')
    tree.column('#1', width=30)
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Patch Cycle')
    tree.heading('#3', text='Patch Days')
    tree.heading('#4', text='Memory Load')
    tree.heading('#5', text='USBDays')
    tree.heading('#6', text='MS17010')
    tree.heading('#7', text='MimiKatz')
    tree.heading('#8', text='Vulnerable Ports')
    tree.column('#1', width=40)
    tree.column('#2', width=150)
    tree.column('#3', width=150)
    tree.column('#4', width=150)
    tree.column('#5', width=100)
    tree.column('#6', width=150)
    tree.column('#7', width=150)
    tree.column('#8', width=300)
    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT DISTINCT
                tbl_SIW_Vuln_Info.VulnInfo_ID,
                tbl_SIW_Vuln_Info.PatchCycle,
                tbl_SIW_Vuln_Info.PatchDays,
                tbl_SIW_Vuln_Info.MemoryLoad,
                tbl_SIW_Vuln_Info.USBDays,
                tbl_SIW_Vuln_Info.MS17010,
                tbl_SIW_Vuln_Info.MimiKatz,
                tbl_SIW_Vuln_Info.VulnerablePorts
                FROM tbl_SIW_Vuln_Info
                WHERE tbl_SIW_Vuln_Info.SIWAsset_ID = {0}""".format(SIWAsset_ID)
    cursor.execute(SQL)
    records = cursor.fetchall()

    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count,
                        values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7]))
            count += 1
    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)


# ----------------------------------------
#          Begin Switch Assets
# ----------------------------------------

def manageSwitchAssets(client_ID):
    global switchAssetsWindow
    global count
    global count2
    Mlist = []
    MoList = []

    def populateListBox(ListBox, Records):
        if ListBox == 'Lbl_AllSwitchAssets':
            for i in range(len(Records)):
                record = Records[i]
                Lbl_AllSIWAssets.insert('end', record[1])
            Lbl_AllSIWAssets.select_set(0)
            Lbl_AllSIWAssets.event_generate("<<ListboxSelect>>")

    # function within function allows the global var SIWAsset_ID to be changed - which is passed later on
    def updateListboxSel(event):
        global SwitchAsset_ID

        currSelection = Lbl_AllSIWAssets.curselection()

        if currSelection != ():
            SwitchAsset = Lbl_AllSIWAssets.get(currSelection)
            indexNum = currSelection[0] + 1
            for i in range(len(switchAssets)):
                record = switchAssets[i]
                if record[1] == SwitchAsset:
                    SwitchAsset_ID = record[0]
            populateGeneralInfo(SwitchAsset_ID)

    def populateGeneralInfo(SwitchAsset_ID):
        count = 0
        count2 = 0
        MoList = []
        manufacturerName = ""
        # create connection to database and execute SQL code
        conn = pyodbc.connect(cnxn)
        cursor = conn.cursor()
        cursor.execute(
            """SELECT
                tbl_SwitchInfo.SerialNumber, 
                tbl_SwitchInfo.Firmware,
                tbl_SwitchInfo.BaseMAC,
                tbl_SwitchInfo.DateCollected,
                tbl_Manufacturers.ManufacturerName,
                tbl_Manufacturers.Manufacturer_ID,
                tbl_Models.ModelName,
                tbl_Models.Model_ID,
                tbl_AssetTypes.Name
                FROM tbl_SwitchAssets as p

                Inner Join tbl_SwitchInfo On p.SwitchAsset_ID = tbl_SwitchInfo.SwitchAsset_ID
                Inner Join tbl_Models On tbl_SwitchInfo.Model_ID = tbl_Models.Model_ID
                Inner Join tbl_Manufacturers On tbl_Models.Manufacturer_ID = tbl_Manufacturers.Manufacturer_ID
                Inner Join tbl_AssetTypes On p.AssetType_ID = tbl_AssetTypes.AssetType_ID

                Where p.SwitchAsset_ID = """ + str(SwitchAsset_ID))
        SwitchInfo = cursor.fetchall()
        for record in SwitchInfo:
            serialNumber_var.set(record[0])
            firmware_Var.set(record[1])
            BaseMAC_Var.set(record[2])
            dateCollected_var.set(record[3])
            manufacturerName = record[4]
            manufacturer_ID = record[5]
            modelName = record[6]
            model_ID = record[7]
        # set Manufacturer
        for record in Manufacturers:
            if record[0] == manufacturerName:
                Cmb_manufacturer.current(count)
            count += 1

        # append Models to cmb_Model
        for record in ModelManufacturer:
            if manufacturer_ID == record[2]:
                MoList.append(record[1])
        Cmb_model["values"] = MoList
        # set Model
        Cmb_model.set(SwitchInfo[0][6])
        Cmb_AssetType.set(SwitchInfo[0][8])
        return

    def updateModelManufacturer(event):
        newModelList = []
        for record in ModelManufacturer:
            if record[3] == Cmb_manufacturer.get():
                newModelList.append(record[1])
        Cmb_model["values"] = newModelList
        Cmb_model.set(newModelList[0])
        return

    def deleteSIWAsset(SwitchAsset_ID):

        conn = pyodbc.connect(cnxn)
        cursor = conn.cursor()
        cursor.execute("""UPDATE tbl_SwitchAssets
                            SET	Deleted = 1
                            WHERE SwitchAsset_ID = """ + str(SwitchAsset_ID))
        cursor.commit()
        switchAssetsWindow.destroy()
        AssetTreeViewWindow.destroy()
        showAssetTree(client_ID)
        manageSwitchAssets(client_ID)
        return

    def edit():
        global btn_Save
        global btn_Cancel
        btn_Save = tk.Button(Frame1, text="Save", bg='white', command=save)
        btn_Cancel = tk.Button(Frame1, text="Cancel", bg='white', command=cancel)
        btn_Save.place(x=150, y=400, width=120)
        btn_Cancel.place(x=300, y=400, width=120)
        entry_DateCollected['state'] = "normal"
        Cmb_manufacturer['state'] = "normal"
        Cmb_model['state'] = "normal"
        Cmb_AssetType['state'] = "normal"
        entry_serialNumber['state'] = "normal"
        entry_Firmware['state'] = "normal"
        entry_BaseMAC['state'] = "normal"
        return

    def cancel():
        entry_DateCollected['state'] = "disabled"
        Cmb_manufacturer['state'] = "disabled"
        Cmb_model['state'] = "disabled"
        Cmb_AssetType['state'] = "disabled"
        entry_serialNumber['state'] = "disabled"
        entry_Firmware['state'] = "disabled"
        entry_BaseMAC['state'] = "disabled"
        btn_Save.place_forget()
        btn_Cancel.place_forget()
        return

    def save():
        global SwitchAsset_ID
        for record in ModelManufacturer:
            if record[1] == Cmb_model.get():
                manufacturer_ID = record[2]
                model_ID = record[0]
        for record in assetTypes:
            if record[0] == Cmb_AssetType.get():
                assetType_ID = record[1]

        SQL = """UPDATE tbl_SwitchInfo
                            SET Model_ID = {0}, SerialNumber = '{1}', Firmware = '{2}', BaseMAC = '{3}',DateCollected = '{4}'
                            WHERE SwitchAsset_ID = {5}""".format(model_ID, entry_serialNumber.get(),
                                                                 entry_Firmware.get(), entry_BaseMAC.get(),
                                                                 entry_DateCollected.get(), SwitchAsset_ID)
        cursor.execute(SQL)
        cursor.commit()
        SQL = """Update tbl_SwitchAssets
                    SET AssetType_ID = {0}
                    WHERE SwitchAsset_ID = {1}""".format(assetType_ID, SwitchAsset_ID)
        cursor.execute(SQL)
        cursor.commit()
        return

    # make window
    AssetTreeViewWindow.withdraw()
    switchAssetsWindow = tk.Toplevel(homeWindow)
    switchAssetsWindow.geometry("600x300")
    switchAssetsWindow.title("Switch Assets")
    switchAssetsWindow.configure(bg='gray15')
    window_width = 870
    window_height = 500
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    switchAssetsWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    Lbl_AllSIWAssets = tk.Listbox(switchAssetsWindow)
    Lbl_AllSIWAssets.place(x=5, y=5, width=150)
    Lbl_AllSIWAssets.bind('<<ListboxSelect>>', updateListboxSel)

    Frame1 = tk.Frame(switchAssetsWindow, padx=5, pady=5, bg='white')
    Frame1.place(x=160, y=5, width=500, height=450)

    label_GeneralInfo = tk.Label(Frame1, text="General Info:", bg='white')
    label_GeneralInfo.config(font=('TkDefaultFont', 15, 'bold'))
    label_GeneralInfo.place(x=185, y=5, width=120)

    label_AssetType = tk.Label(Frame1, text="AssetType", bg='white', anchor='w')
    label_AssetType.config(font=('TkDefaultFont', 10, 'bold'))
    label_AssetType.place(x=5, y=50, width=120)
    Cmb_AssetType = ttk.Combobox(Frame1)
    Cmb_AssetType.bind("<<ComboboxSelected>>", updateModelManufacturer)
    Cmb_AssetType['state'] = "disabled"
    Cmb_AssetType.place(x=130, y=50, width=200)

    label_DateCollected = tk.Label(Frame1, text="Date Collected", bg='white', anchor='w')
    label_DateCollected.config(font=('TkDefaultFont', 10, 'bold'))
    label_DateCollected.place(x=5, y=100, width=120)
    dateCollected_var = tk.StringVar()
    entry_DateCollected = tk.Entry(Frame1, textvariable=dateCollected_var)
    entry_DateCollected['state'] = "disabled"
    entry_DateCollected.place(x=130, y=100, width=200)

    label_Manufacturer = tk.Label(Frame1, text="Manufacturer", bg='white', anchor='w')
    label_Manufacturer.config(font=('TkDefaultFont', 10, 'bold'))
    label_Manufacturer.place(x=5, y=150, width=120)
    manufacturer_var = tk.StringVar()
    Cmb_manufacturer = ttk.Combobox(Frame1)
    Cmb_manufacturer.bind("<<ComboboxSelected>>", updateModelManufacturer)
    Cmb_manufacturer['state'] = "disabled"
    Cmb_manufacturer.place(x=130, y=150, width=200)

    label_Model = tk.Label(Frame1, text="Model", bg='white', anchor='w')
    label_Model.config(font=('TkDefaultFont', 10, 'bold'))
    label_Model.place(x=5, y=200, width=120)
    model_var = tk.StringVar()
    Cmb_model = ttk.Combobox(Frame1, textvariable=model_var)
    Cmb_model.bind("<<ComboboxSelected>>", updateModelManufacturer)
    Cmb_model['state'] = "disabled"
    Cmb_model.place(x=130, y=200, width=200)

    label_SerialNumber = tk.Label(Frame1, text="Serial Number", bg='white', anchor='w')
    label_SerialNumber.config(font=('TkDefaultFont', 10, 'bold'))
    label_SerialNumber.place(x=5, y=250, width=120)
    serialNumber_var = tk.StringVar()
    entry_serialNumber = tk.Entry(Frame1, textvariable=serialNumber_var)
    entry_serialNumber['state'] = "disabled"
    entry_serialNumber.place(x=130, y=250, width=200)

    label_Firmware = tk.Label(Frame1, text="Firmware Version", bg='white', anchor='w')
    label_Firmware.config(font=('TkDefaultFont', 10, 'bold'))
    label_Firmware.place(x=5, y=300, width=120)
    firmware_Var = tk.StringVar()
    entry_Firmware = tk.Entry(Frame1, textvariable=firmware_Var)
    entry_Firmware['state'] = "disabled"
    entry_Firmware.place(x=130, y=300, width=200)

    label_BaseMAC = tk.Label(Frame1, text="Base MAC", bg='white', anchor='w')
    label_BaseMAC.config(font=('TkDefaultFont', 10, 'bold'))
    label_BaseMAC.place(x=5, y=350, width=120)
    BaseMAC_Var = tk.StringVar()
    entry_BaseMAC = tk.Entry(Frame1, textvariable=BaseMAC_Var)
    entry_BaseMAC['state'] = "disabled"
    entry_BaseMAC.place(x=130, y=350, width=200)

    # Create and place buttons on right side of window + Back and Home buttons
    btn_addAsset = tk.Button(switchAssetsWindow, width=5, text="Add Asset", bg='white', command=addSwitchAsset)
    btn_deleteAsset = tk.Button(switchAssetsWindow, width=5, text="Delete Asset", bg='white',
                                command=lambda: deleteSIWAsset(SwitchAsset_ID))
    btn_SwitchPorts = tk.Button(switchAssetsWindow, width=15, text="View Switch Ports", bg='white',
                                command=lambda: viewSwitchPorts(SwitchAsset_ID))
    btn_home = tk.Button(switchAssetsWindow, width=5, text="Home", bg='white',
                         command=lambda: home(switchAssetsWindow, homeWindow))
    btn_back = tk.Button(switchAssetsWindow, width=5, text="Back", bg='white',
                         command=lambda: home(switchAssetsWindow, AssetTreeViewWindow))
    btn_Edit = tk.Button(Frame1, text="Edit", bg='white', command=edit)

    btn_addAsset.place(x=665, y=5, width=200)
    btn_deleteAsset.place(x=665, y=55, width=200)
    btn_SwitchPorts.place(x=665, y=155, width=200)
    btn_home.place(x=5, y=260, width=150)
    btn_back.place(x=5, y=310, width=150)
    btn_Edit.place(x=5, y=400, width=120)

    # create connection to database and execute SQL code
    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    cursor.execute("""SELECT * FROM tbl_SwitchAssets
                    INNER JOIN tbl_Buildings ON tbl_SwitchAssets.Building_ID = tbl_Buildings.Building_ID
                    INNER JOIN tbl_Client_Info ON tbl_Buildings.Client_ID = tbl_Client_Info.Client_ID
					INNER JOIN tbl_Contacts ON tbl_Client_Info.Contact_ID = tbl_Contacts.Contact_ID
                    WHERE tbl_SwitchAssets.Deleted IS NULL AND tbl_Contacts.Company_ID = """ + str(client_ID))

    switchAssets = cursor.fetchall()
    # assign query results to list box

    cursor.execute("""SELECT tbl_Models.Model_ID, tbl_Models.ModelName, tbl_Models.Manufacturer_ID, tbl_Manufacturers.ManufacturerName
                        FROM tbl_Models
                        INNER JOIN tbl_Manufacturers ON tbl_Models.Manufacturer_ID = tbl_Manufacturers.Manufacturer_ID""")
    ModelManufacturer = cursor.fetchall()

    cursor.execute("""SELECT tbl_Manufacturers.ManufacturerName
                            FROM tbl_Manufacturers""")
    Manufacturers = cursor.fetchall()
    for record in Manufacturers:
        Mlist.append(record[0])
    Cmb_manufacturer["values"] = Mlist
    populateListBox('Lbl_AllSwitchAssets', switchAssets)

    cursor.execute("""SELECT tbl_AssetTypes.Name, tbl_AssetTypes.AssetType_ID
                            FROM tbl_AssetTypes""")
    assetTypes = cursor.fetchall()
    assetList = []
    for record in assetTypes:
        assetList.append(record[0])
    Cmb_AssetType["values"] = assetList
    return


def addSwitchAsset():
    def populateCombo(comboToPopulate):
        modelList = []
        conn = pyodbc.connect(cnxn)
        cursor = conn.cursor()

        if comboToPopulate == "Ent_Model":
            cursor.execute("""SELECT tbl_Models.ModelName
                              FROM tbl_Models""")
            models = cursor.fetchall()
            for model in models:
                modelList.append(model[0])
            Ent_Model['values'] = modelList

        return

    def addAsset():
        if Ent_HostName.get() != '':
            conn = pyodbc.connect(cnxn)
            cursor = conn.cursor()
            SQL = '''INSERT INTO tbl_SwitchAssets (HostName, AssetType_ID)
                      VALUES ('{0}', {1})'''.format(Ent_HostName.get(), 11)
            cursor.execute(SQL)
            cursor.commit()

            SQL = """SELECT tbl_SwitchAssets.SwitchAsset_ID 
                    FROM tbl_SwitchAssets
                    WHERE HostName = '{0}'""".format(Ent_HostName.get())
            cursor.execute(SQL)
            asset_ID = cursor.fetchall()
            asset_ID = asset_ID[0][-1]

            SQL = """SELECT tbl_Models.Model_ID FROM tbl_Models WHERE ModelName = '{0}'""".format(Ent_Model.get())
            cursor.execute(SQL)
            ModelRecord = cursor.fetchall()
            model_ID = ModelRecord[0][0]

            SQL = '''INSERT INTO tbl_SwitchInfo (SwitchAsset_ID, Model_ID, SerialNumber, Firmware, BaseMAC, DateCollected)
                      VALUES ({0},{1},'{2}','{3}','{4}','{5}')'''.format(asset_ID, model_ID, Ent_SerialNumber.get(),
                                                                         Ent_Firmware.get(), Ent_BaseMAC.get(),
                                                                         Ent_DateCollected.get())
            cursor.execute(SQL)
            cursor.commit()

    global addSwitchAssetWindow
    switchAssetsWindow.withdraw()
    addSwitchAssetWindow = tk.Toplevel(AssetTreeViewWindow)
    addSwitchAssetWindow.title("Add a New Switch Asset")
    addSwitchAssetWindow.configure(bg='gray15')
    window_width = 600
    window_height = 450
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    addSwitchAssetWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    Lbl_HostName = tk.Label(addSwitchAssetWindow, width=5, text="HostName")
    Lbl_Model = tk.Label(addSwitchAssetWindow, width=5, text="Model")
    Lbl_SerialNumber = tk.Label(addSwitchAssetWindow, width=5, text="Serial Number")
    Lbl_Firmware = tk.Label(addSwitchAssetWindow, width=5, text="Firmware")
    Lbl_BaseMAC = tk.Label(addSwitchAssetWindow, width=5, text="Base MAC")
    Lbl_DateCollected = tk.Label(addSwitchAssetWindow, width=5, text="Date Collected")

    Ent_HostName = tk.Entry(addSwitchAssetWindow, width=10)
    Ent_Model = ttk.Combobox(addSwitchAssetWindow, width=8)
    Ent_SerialNumber = tk.Entry(addSwitchAssetWindow, width=10)
    Ent_Firmware = tk.Entry(addSwitchAssetWindow, width=10)
    Ent_BaseMAC = tk.Entry(addSwitchAssetWindow, width=10)
    Ent_DateCollected = tk.Entry(addSwitchAssetWindow, width=10)

    btn_back = tk.Button(addSwitchAssetWindow, width=5, text="Back", bg='azure4',
                         command=lambda: home(addSwitchAssetWindow, switchAssetsWindow))
    btn_Home = tk.Button(addSwitchAssetWindow, width=5, text="Home", bg='azure4',
                         command=lambda: home(AssetTreeViewWindow, homeWindow))
    btn_Save = tk.Button(addSwitchAssetWindow, width=5, text="Save", bg='azure4', command=addAsset)

    Lbl_HostName.grid(row=0, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    Lbl_Model.grid(row=1, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    Lbl_SerialNumber.grid(row=2, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    Lbl_Firmware.grid(row=3, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    Lbl_BaseMAC.grid(row=4, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    Lbl_DateCollected.grid(row=5, column=0, pady=5, padx=5, ipadx=50, sticky='n')

    Ent_HostName.grid(row=0, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    Ent_Model.grid(row=1, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    Ent_SerialNumber.grid(row=2, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    Ent_Firmware.grid(row=3, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    Ent_BaseMAC.grid(row=4, column=1, pady=5, padx=5, ipadx=50, sticky='n')
    Ent_DateCollected.grid(row=5, column=1, pady=5, padx=5, ipadx=50, sticky='n')

    btn_back.grid(row=6, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    btn_Home.grid(row=7, column=0, pady=5, padx=5, ipadx=50, sticky='n')
    btn_Save.grid(row=6, column=1, pady=5, padx=5, ipadx=50, sticky='n')

    populateCombo("Ent_Model")

    return


def viewSwitchPorts(SwitchAsset_ID):
    global SwitchPortsWindow
    global SwitchPorts_ID

    def updateWidgets(event):
        global SwitchPorts_ID
        curItem = tree.focus()
        Selection = tree.item(curItem)
        SoftwareInfo = Selection.get('values')
        SwitchPorts_ID = SoftwareInfo[0]
        try:
            if entry_SoftwareName.winfo_exists():
                SwitchPorts_ID = SoftwareInfo[0]
                Software_Var.set(SoftwareInfo[1])
                Corp_Var.set(SoftwareInfo[2])
                Date_Var.set(SoftwareInfo[3])
                Version_Var.set(SoftwareInfo[4])
                Cmb_CVEs.set(SoftwareInfo[5])
        except:
            return
        return

    def cancel():
        Frame1.grid_forget()
        return

    def edit():
        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global entry_PortNumber
            global entry_Description
            global entry_AdminStatus
            global entry_PortType
            global entry_NativeVLAN
            global entry_VLANTag

            global PortNumber_Var
            global Description_Var
            global AdminStatus_Var
            global PortType_Var
            global NativeVLAN_Var
            global VLANTag_Var

            global Btn_Cancel
            global Btn_Save

            def save():
                curItem = tree.focus()
                Selection = tree.item(curItem)
                SwitchPortsInfo = Selection.get('values')

                SwitchPorts_ID = SwitchPortsInfo[0]
                PortNumber = 'NULL'
                Description = 'NULL'
                AdminStatus = 'NULL'
                PortType = 'NULL'
                NativeVLAN = 'NULL'
                VLANTag = 'NULL'

                if entry_PortNumber.get() != "":
                    PortNumber = entry_PortNumber.get()
                if entry_Description.get() != "":
                    Description = entry_Description.get()
                if entry_AdminStatus.get() != "":
                    AdminStatus = entry_AdminStatus.get()
                if entry_PortType.get() != "":
                    PortType = entry_PortType.get()
                if entry_NativeVLAN.get() != "":
                    NativeVLAN = entry_NativeVLAN.get()
                if entry_VLANTag.get() != "":
                    VLANTag = entry_VLANTag.get()

                SQL = """UPDATE tbl_SwitchPorts
                        SET PortNumber = '{0}', Description = '{1}', AdminStatus = '{2}', PortType = '{3}', NativeVLAN = '{4}', VLANTag = '{5}'
                        WHERE SwitchPorts_ID = {6}""".format(PortNumber, Description, AdminStatus, PortType, NativeVLAN,
                                                             VLANTag, SwitchPortsInfo[0])
                cursor.execute(SQL)
                cursor.commit()
                return

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            SwitchPortInfo = Selection.get('values')

            # Create Entry Widgets/ Combobox
            PortNumber_Var = tk.StringVar()
            Description_Var = tk.StringVar()
            AdminStatus_Var = tk.StringVar()
            PortType_Var = tk.StringVar()
            NativeVLAN_Var = tk.StringVar()
            VLANTag_Var = tk.StringVar()

            Frame1 = tk.Frame(SwitchPortsWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_PortNumber = tk.Label(Frame1, text="Port Number", bg='white', anchor='w')
            lbl_Description = tk.Label(Frame1, text="Description", bg='white', anchor='w')
            lbl_AdminStatus = tk.Label(Frame1, text="Admin Status", bg='white', anchor='w')
            lbl_PortType = tk.Label(Frame1, text="Port Type", bg='white', anchor='w')
            lbl_NativeVLAN = tk.Label(Frame1, text="Native VLAN", bg='white', anchor='w')
            lbl_VLANTag = tk.Label(Frame1, text="VLAN Tag", bg='white', anchor='w')

            entry_PortNumber = tk.Entry(Frame1, textvariable=PortNumber_Var, width=40)
            entry_Description = tk.Entry(Frame1, textvariable=Description_Var, width=30)
            entry_AdminStatus = tk.Entry(Frame1, textvariable=AdminStatus_Var, width=15)
            entry_PortType = tk.Entry(Frame1, textvariable=PortType_Var, width=15)
            entry_NativeVLAN = tk.Entry(Frame1, textvariable=NativeVLAN_Var, width=15)
            entry_VLANTag = tk.Entry(Frame1, textvariable=VLANTag_Var, width=15)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_PortNumber.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Description.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_AdminStatus.grid(row=4, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_PortType.grid(row=5, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_NativeVLAN.grid(row=6, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_VLANTag.grid(row=7, column=0, pady=5, padx=5, sticky='w', columnspan=2)

            entry_PortNumber.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_Description.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_AdminStatus.grid(row=4, column=1, pady=5, padx=5, sticky='w')
            entry_PortType.grid(row=5, column=1, pady=5, padx=5, sticky='w')
            entry_NativeVLAN.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            entry_VLANTag.grid(row=7, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=8, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=8, column=1, pady=5, padx=5, sticky='w')

            # Set the first two entry widgets values
            PortNumber_Var.set(SwitchPortInfo[1])
            Description_Var.set(SwitchPortInfo[2])
            AdminStatus_Var.set(SwitchPortInfo[3])
            PortType_Var.set(SwitchPortInfo[4])
            NativeVLAN_Var.set(SwitchPortInfo[5])
            VLANTag_Var.set(SwitchPortInfo[6])

            return

    def delete():
        curItem = tree.focus()
        Selection = tree.item(curItem)
        MACsInfo = Selection.get('values')
        SQL = """UPDATE tbl_ConnectedMACs
                    SET Deleted = 1
                    WHERE ConnectedMACs_ID = {0}""".format(MACsInfo[0])
        cursor.execute(SQL)
        cursor.commit()
        tree.delete(tree.selection())
        return

    def newEntry(SwitchAsset_ID):

        def save(SwitchAsset_ID):
            curItem = tree.focus()
            Selection = tree.item(curItem)
            SwitchPortsInfo = Selection.get('values')

            SwitchPorts_ID = SwitchPortsInfo[0]
            PortNumber = 'NULL'
            Description = 'NULL'
            AdminStatus = 'NULL'
            PortType = 'NULL'
            NativeVLAN = 'NULL'
            VLANTag = 'NULL'

            if entry_PortNumber.get() != "":
                PortNumber = entry_PortNumber.get()
            if entry_Description.get() != "":
                Description = entry_Description.get()
            if entry_AdminStatus.get() != "":
                AdminStatus = entry_AdminStatus.get()
            if entry_PortType.get() != "":
                PortType = entry_PortType.get()
            if entry_NativeVLAN.get() != "":
                NativeVLAN = entry_NativeVLAN.get()
            if entry_VLANTag.get() != "":
                VLANTag = entry_VLANTag.get()

            SQL = """INSERT INTO tbl_SwitchPorts (SwitchAsset_ID, PortNumber, Description, AdminStatus, PortType, NativeVLAN, VLANTag)
                        VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}')""".format(SwitchAsset_ID, PortNumber,
                                                                                   Description, AdminStatus, PortType,
                                                                                   NativeVLAN, VLANTag)
            cursor.execute(SQL)
            cursor.commit()
            return

        global Frame1
        try:
            Frame1.grid_forget()
        finally:
            global entry_PortNumber
            global entry_Description
            global entry_AdminStatus
            global entry_PortType
            global entry_NativeVLAN
            global entry_VLANTag

            global Btn_Cancel
            global Btn_Save

            # get users current selection in treeview
            curItem = tree.focus()
            Selection = tree.item(curItem)
            SwitchPortInfo = Selection.get('values')

            Frame1 = tk.Frame(SwitchPortsWindow, padx=5, pady=5, bg='white')
            Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

            lbl_PortNumber = tk.Label(Frame1, text="Port Number", bg='white', anchor='w')
            lbl_Description = tk.Label(Frame1, text="Description", bg='white', anchor='w')
            lbl_AdminStatus = tk.Label(Frame1, text="Admin Status", bg='white', anchor='w')
            lbl_PortType = tk.Label(Frame1, text="Port Type", bg='white', anchor='w')
            lbl_NativeVLAN = tk.Label(Frame1, text="Native VLAN", bg='white', anchor='w')
            lbl_VLANTag = tk.Label(Frame1, text="VLAN Tag", bg='white', anchor='w')

            entry_PortNumber = tk.Entry(Frame1, width=40)
            entry_Description = tk.Entry(Frame1, width=30)
            entry_AdminStatus = tk.Entry(Frame1, width=15)
            entry_PortType = tk.Entry(Frame1, width=15)
            entry_NativeVLAN = tk.Entry(Frame1, width=15)
            entry_VLANTag = tk.Entry(Frame1, width=15)

            Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
            Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

            # Place widgets in window
            lbl_PortNumber.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_Description.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_AdminStatus.grid(row=4, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_PortType.grid(row=5, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_NativeVLAN.grid(row=6, column=0, pady=5, padx=5, sticky='w', columnspan=2)
            lbl_VLANTag.grid(row=7, column=0, pady=5, padx=5, sticky='w', columnspan=2)

            entry_PortNumber.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_Description.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
            entry_AdminStatus.grid(row=4, column=1, pady=5, padx=5, sticky='w')
            entry_PortType.grid(row=5, column=1, pady=5, padx=5, sticky='w')
            entry_NativeVLAN.grid(row=6, column=1, pady=5, padx=5, sticky='w')
            entry_VLANTag.grid(row=7, column=1, pady=5, padx=5, sticky='w')

            Btn_Cancel.grid(row=8, column=0, pady=5, padx=5, sticky='w')
            Btn_Save.grid(row=8, column=1, pady=5, padx=5, sticky='w')

        return

    def connectedMACs(SwitchAsset_ID, SwitchPorts_ID):
        global connectedMACsWindow

        def updateWidgets(event):
            try:
                if entry_PortNumber.winfo_exists():
                    curItem = tree.focus()
                    Selection = tree.item(curItem)
                    ConnectedMACs = Selection.get('values')
                    entry_PortNumber.set(ConnectedMACs[1])
                    entry_MAC.set(ConnectedMACs[2])
            except:
                return
            return

        def cancel():
            Frame1.grid_forget()
            return

        def edit():
            global Frame1
            try:
                Frame1.grid_forget()
            finally:
                global entry_PortNumber
                global entry_PortNumber

                global PortNumber_Var
                global MAC_Var

                global Btn_Cancel
                global Btn_Save

                def save():
                    curItem = tree.focus()
                    Selection = tree.item(curItem)
                    MACsInfo = Selection.get('values')

                    SwitchPorts_ID = MACsInfo[0]
                    MACAddress = 'NULL'

                    if entry_MAC.get() != "":
                        MACAddress = entry_MAC.get()

                    SQL = """UPDATE tbl_ConnectedMACs
                            SET MACAddress = '{0}'
                            WHERE ConnectedMACs_ID = {1}""".format(MACAddress, MACsInfo[0])
                    cursor.execute(SQL)
                    cursor.commit()
                    return

                # get users current selection in treeview
                curItem = tree.focus()
                Selection = tree.item(curItem)
                MACInfo = Selection.get('values')

                # Create Entry Widgets/ Combobox
                PortNumber_Var = tk.StringVar()
                MAC_Var = tk.StringVar()

                Frame1 = tk.Frame(connectedMACsWindow, padx=5, pady=5, bg='white')
                Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

                lbl_PortNumber = tk.Label(Frame1, text="Port Number", bg='white', anchor='w')
                lbl_MAC = tk.Label(Frame1, text="MACAddress", bg='white', anchor='w')
                entry_PortNumber = tk.Entry(Frame1, textvariable=PortNumber_Var, width=40)
                entry_MAC = tk.Entry(Frame1, textvariable=MAC_Var, width=30)
                Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
                Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

                entry_PortNumber['state'] = "disabled"

                # Place widgets in window
                lbl_PortNumber.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
                lbl_MAC.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
                entry_PortNumber.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
                entry_MAC.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
                Btn_Cancel.grid(row=8, column=0, pady=5, padx=5, sticky='w')
                Btn_Save.grid(row=8, column=1, pady=5, padx=5, sticky='w')

                # Set the first two entry widgets values
                PortNumber_Var.set(MACInfo[1])
                MAC_Var.set(MACInfo[2])
                return

        def delete():
            curItem = tree.focus()
            Selection = tree.item(curItem)
            SwitchPortsInfo = Selection.get('values')
            SQL = """UPDATE tbl_ConnectedMACs
                        SET Deleted = 1
                        WHERE ConnectedMACs_ID = {0}""".format(SwitchPortsInfo[0])
            cursor.execute(SQL)
            cursor.commit()
            tree.delete(tree.selection())
            return

        def newEntry(SwitchPorts_ID):

            def save(SwitchPorts_ID):
                curItem = tree.focus()
                Selection = tree.item(curItem)
                MACsInfo = Selection.get('values')

                MACAddress = 'NULL'

                if entry_MAC.get() != "":
                    MACAddress = entry_MAC.get()

                SQL = """INSERT INTO tbl_ConnectedMACs (SwitchPorts_ID, MACAddress)
                            VALUES ({0},'{1}')""".format(SwitchPorts_ID, MACAddress)
                cursor.execute(SQL)
                cursor.commit()
                return

            global Frame1
            try:
                Frame1.grid_forget()
            finally:
                global entry_PortNumber
                global entry_PortNumber

                global Btn_Cancel
                global Btn_Save

                PortNumber_Var = tk.StringVar()

                Frame1 = tk.Frame(connectedMACsWindow, padx=5, pady=5, bg='white')
                Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

                lbl_PortNumber = tk.Label(Frame1, text="Port Number", bg='white', anchor='w')
                lbl_MAC = tk.Label(Frame1, text="MACAddress", bg='white', anchor='w')
                entry_PortNumber = tk.Entry(Frame1, textvariable=PortNumber_Var, width=40)
                entry_MAC = tk.Entry(Frame1, width=30)
                Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
                Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=lambda: save(SwitchPorts_ID))

                entry_PortNumber['state'] = "disabled"

                # Place widgets in window
                lbl_PortNumber.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
                lbl_MAC.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
                entry_PortNumber.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
                entry_MAC.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
                Btn_Cancel.grid(row=8, column=0, pady=5, padx=5, sticky='w')
                Btn_Save.grid(row=8, column=1, pady=5, padx=5, sticky='w')
                PortNumber_Var.set(records[0][1])

                return

            return

        connectedMACsWindow = tk.Toplevel(switchAssetsWindow)
        SwitchPortsWindow.withdraw()
        connectedMACsWindow.title("Connected MACs")
        connectedMACsWindow.configure(bg='gray15')
        window_width = 850
        window_height = 450
        screen_width = homeWindow.winfo_screenwidth()
        screen_height = homeWindow.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        connectedMACsWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

        btn_home = tk.Button(connectedMACsWindow, width=15, text="Home", bg='white',
                             command=lambda: home(connectedMACsWindow, homeWindow))
        btn_back = tk.Button(connectedMACsWindow, width=15, text="Back", bg='white',
                             command=lambda: home(connectedMACsWindow, SwitchPortsWindow))
        btn_edit = tk.Button(connectedMACsWindow, width=15, text="Edit Entry", bg='white',
                             command=lambda: edit())
        btn_new = tk.Button(connectedMACsWindow, width=15, text="New Entry", bg='white',
                            comman=lambda: newEntry(SwitchPorts_ID))
        btn_delete = tk.Button(connectedMACsWindow, width=15, text="Delete Entry", bg='white', command=delete)

        btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
        btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
        btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
        btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
        btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

        tree = ttk.Treeview(connectedMACsWindow, column=('#1', '#2', '#3'), show='headings')
        tree.column('#1', width=30)
        tree.bind("<<TreeviewSelect>>", updateWidgets)
        tree.heading('#1', text='ID')
        tree.heading('#2', text='Port Number')
        tree.heading('#3', text='MAC Address')

        conn = pyodbc.connect(cnxn)
        cursor = conn.cursor()
        SQL = """SELECT DISTINCT
                    tbl_SwitchPorts.SwitchPorts_ID,
                    tbl_SwitchPorts.PortNumber,
                    tbl_ConnectedMACs.ConnectedMACs_ID,
                    tbl_ConnectedMACs.MACAddress
                    FROM tbl_SwitchPorts
                    INNER JOIN tbl_ConnectedMACs ON tbl_SwitchPorts.SwitchPorts_ID = tbl_ConnectedMACs.SwitchPorts_ID
                    WHERE tbl_SwitchPorts.SwitchPorts_ID = {0} AND tbl_ConnectedMACs.Deleted IS NULL""".format(
            SwitchAsset_ID)
        cursor.execute(SQL)
        records = cursor.fetchall()

        if len(records) == 0:
            messagebox.showinfo("Oops!", "No Records Found")
        else:
            count = 0
            for record in records:
                tree.insert('', tk.END, iid=count,
                            values=(record[2], record[1], record[3]))
                count += 1

        tree.grid(row=0, column=0, columnspan=41)
        if len(records) > 0:
            tree.selection_set(0)
            tree.focus(0)
        return

    def misconfigs(SwitchAsset_ID, SwitchPorts_ID):

        global connectedMACsWindow

        def updateWidgets(event):
            try:
                if entry_PortNumber.winfo_exists():
                    curItem = tree.focus()
                    Selection = tree.item(curItem)
                    ConnectedMACs = Selection.get('values')
                    entry_PortNumber.set(ConnectedMACs[1])
                    entry_MAC.set(ConnectedMACs[2])
            except:
                return
            return

        def cancel():
            Frame1.grid_forget()
            return

        def edit():
            global Frame1
            try:
                Frame1.grid_forget()
            finally:
                global entry_PortNumber
                global Cmb_Misconfig

                global PortNumber_Var

                global Btn_Cancel
                global Btn_Save

                def save():
                    curItem = tree.focus()
                    Selection = tree.item(curItem)
                    MisInfo = Selection.get('values')

                    PortMisconfigurations_ID = MisInfo[0]

                    for record in Misconfigurations:
                        if record[0] == Cmb_Misconfig.get():
                            Error_ID = record[1]

                    SQL = """UPDATE tbl_PortMisconfigurations
                            SET Error_ID = {0}
                            WHERE PortMisconfigurations_ID = {1}""".format(Error_ID, PortMisconfigurations_ID)
                    cursor.execute(SQL)
                    cursor.commit()
                    connectedMACsWindow.destroy()
                    misconfigs(SwitchAsset_ID, SwitchPorts_ID)
                    return

                # get users current selection in treeview
                curItem = tree.focus()
                Selection = tree.item(curItem)
                MisInfo = Selection.get('values')

                # Create Entry Widgets/ Combobox
                PortNumber_Var = tk.StringVar()

                Frame1 = tk.Frame(connectedMACsWindow, padx=5, pady=5, bg='white')
                Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

                lbl_PortNumber = tk.Label(Frame1, text="Port Number", bg='white', anchor='w')
                lbl_Misconfig = tk.Label(Frame1, text="Misconfiguration", bg='white', anchor='w')
                entry_PortNumber = tk.Entry(Frame1, textvariable=PortNumber_Var, width=20)
                Cmb_Misconfig = ttk.Combobox(Frame1, width=50)
                Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
                Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=save)

                entry_PortNumber['state'] = "disabled"

                # Place widgets in window
                lbl_PortNumber.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
                lbl_Misconfig.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
                entry_PortNumber.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
                Cmb_Misconfig.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
                Btn_Cancel.grid(row=8, column=0, pady=5, padx=5, sticky='w')
                Btn_Save.grid(row=8, column=1, pady=5, padx=5, sticky='w')

                SQL = """SELECT
                            tbl_Misconfigurations.Name, tbl_Misconfigurations.Error_ID
                            FROM tbl_Misconfigurations"""
                cursor.execute(SQL)
                Misconfigurations = cursor.fetchall()
                MisconfigList = []
                for Misconfig in Misconfigurations:
                    MisconfigList.append(Misconfig[0])
                Cmb_Misconfig["values"] = MisconfigList
                Cmb_Misconfig.set(MisInfo[2])

                # Set the first two entry widgets values
                PortNumber_Var.set(MisInfo[1])
                return

        def delete():
            curItem = tree.focus()
            Selection = tree.item(curItem)
            MisconfigInfo = Selection.get('values')
            SQL = """UPDATE tbl_PortMisconfigurations
                        SET Deleted = 1
                        WHERE tbl_PortMisconfigurations.PortMisconfigurations_ID = {0}""".format(MisconfigInfo[0])
            cursor.execute(SQL)
            cursor.commit()
            tree.delete(tree.selection())
            return

        def newEntry(SwitchPorts_ID):

            def save(SwitchPorts_ID):
                for record in Misconfigurations:
                    if Cmb_Misconfig.get() == record[0]:
                        Error_ID = record[1]

                SQL = """INSERT INTO tbl_PortMisconfigurations(Error_ID, SwitchPorts_ID)
                            VALUES ({0},{1})""".format(Error_ID, SwitchPorts_ID)
                cursor.execute(SQL)
                cursor.commit()
                connectedMACsWindow.destroy()
                misconfigs(SwitchAsset_ID, SwitchPorts_ID)
                return

            global Frame1
            try:
                Frame1.grid_forget()
            finally:
                global entry_PortNumber
                global Cmb_Misconfig

                global PortNumber_Var

                global Btn_Cancel
                global Btn_Save

                # get users current selection in treeview
                curItem = tree.focus()
                Selection = tree.item(curItem)
                MisInfo = Selection.get('values')

                # Create Entry Widgets/ Combobox
                PortNumber_Var = tk.StringVar()

                Frame1 = tk.Frame(connectedMACsWindow, padx=5, pady=5, bg='white')
                Frame1.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=41)

                lbl_PortNumber = tk.Label(Frame1, text="Port Number", bg='white', anchor='w')
                lbl_Misconfig = tk.Label(Frame1, text="Misconfiguration", bg='white', anchor='w')
                entry_PortNumber = tk.Entry(Frame1, textvariable=PortNumber_Var, width=20)
                Cmb_Misconfig = ttk.Combobox(Frame1, width=50)
                Btn_Cancel = tk.Button(Frame1, width=15, text="Cancel", bg='white', command=cancel)
                Btn_Save = tk.Button(Frame1, width=15, text="Save", bg='white', command=lambda: save(SwitchPorts_ID))

                entry_PortNumber['state'] = "disabled"

                # Place widgets in window
                lbl_PortNumber.grid(row=2, column=0, pady=5, padx=5, sticky='w', columnspan=2)
                lbl_Misconfig.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
                entry_PortNumber.grid(row=2, column=1, pady=5, padx=5, sticky='w', columnspan=2)
                Cmb_Misconfig.grid(row=3, column=1, pady=5, padx=5, sticky='w', columnspan=2)
                Btn_Cancel.grid(row=8, column=0, pady=5, padx=5, sticky='w')
                Btn_Save.grid(row=8, column=1, pady=5, padx=5, sticky='w')

                SQL = """SELECT
                                tbl_Misconfigurations.Name, tbl_Misconfigurations.Error_ID
                                FROM tbl_Misconfigurations"""
                cursor.execute(SQL)
                Misconfigurations = cursor.fetchall()
                MisconfigList = []
                for Misconfig in Misconfigurations:
                    MisconfigList.append(Misconfig[0])
                Cmb_Misconfig["values"] = MisconfigList

                PortNumber_Var.set(records[0][1])
                return

            return

        connectedMACsWindow = tk.Toplevel(switchAssetsWindow)
        SwitchPortsWindow.withdraw()
        connectedMACsWindow.title("Misconfigurations")
        connectedMACsWindow.configure(bg='gray15')
        window_width = 850
        window_height = 450
        screen_width = homeWindow.winfo_screenwidth()
        screen_height = homeWindow.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        connectedMACsWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

        btn_home = tk.Button(connectedMACsWindow, width=15, text="Home", bg='white',
                             command=lambda: home(connectedMACsWindow, homeWindow))
        btn_back = tk.Button(connectedMACsWindow, width=15, text="Back", bg='white',
                             command=lambda: home(connectedMACsWindow, SwitchPortsWindow))
        btn_edit = tk.Button(connectedMACsWindow, width=15, text="Edit Entry", bg='white',
                             command=lambda: edit())
        btn_new = tk.Button(connectedMACsWindow, width=15, text="New Entry", bg='white',
                            comman=lambda: newEntry(SwitchPorts_ID))
        btn_delete = tk.Button(connectedMACsWindow, width=15, text="Delete Entry", bg='white', command=delete)

        btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
        btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
        btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
        btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
        btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')

        tree = ttk.Treeview(connectedMACsWindow, column=('#1', '#2', '#3'), show='headings')
        tree.column('#1', width=30)
        tree.bind("<<TreeviewSelect>>", updateWidgets)
        tree.heading('#1', text='ID')
        tree.heading('#2', text='Port Number')
        tree.heading('#3', text='Misconfiguration')
        tree.column('#3', width=400)

        conn = pyodbc.connect(cnxn)
        cursor = conn.cursor()
        SQL = """SELECT DISTINCT
                    tbl_SwitchPorts.SwitchPorts_ID,
                    tbl_SwitchPorts.PortNumber,
                    tbl_PortMisconfigurations.PortMisconfigurations_ID,
                    tbl_Misconfigurations.Name
                    FROM tbl_SwitchPorts
                    INNER JOIN tbl_PortMisconfigurations ON tbl_SwitchPorts.SwitchPorts_ID = tbl_PortMisconfigurations.SwitchPorts_ID
                    INNER JOIN tbl_Misconfigurations ON tbl_PortMisconfigurations.Error_ID = tbl_Misconfigurations.Error_ID
                    WHERE tbl_SwitchPorts.SwitchPorts_ID = {0} AND tbl_PortMisconfigurations.Deleted IS NULL""".format(
            SwitchAsset_ID)
        cursor.execute(SQL)
        records = cursor.fetchall()

        if len(records) == 0:
            messagebox.showinfo("Oops!", "No Records Found")
        else:
            count = 0
            for record in records:
                tree.insert('', tk.END, iid=count,
                            values=(record[2], record[1], record[3]))
                count += 1

        tree.grid(row=0, column=0, columnspan=41)
        if len(records) > 0:
            tree.selection_set(0)
            tree.focus(0)
        return

    SwitchPortsWindow = tk.Toplevel(switchAssetsWindow)
    switchAssetsWindow.withdraw()
    SwitchPortsWindow.geometry("600x300")
    SwitchPortsWindow.title("Switch Ports")
    SwitchPortsWindow.configure(bg='gray15')
    window_width = 1200
    window_height = 600
    screen_width = homeWindow.winfo_screenwidth()
    screen_height = homeWindow.winfo_screenheight()
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)
    SwitchPortsWindow.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    btn_home = tk.Button(SwitchPortsWindow, width=15, text="Home", bg='white',
                         command=lambda: home(SwitchPortsWindow, homeWindow))
    btn_back = tk.Button(SwitchPortsWindow, width=15, text="Back", bg='white',
                         command=lambda: home(SwitchPortsWindow, switchAssetsWindow))
    btn_edit = tk.Button(SwitchPortsWindow, width=15, text="Edit Entry", bg='white',
                         command=lambda: edit())
    btn_new = tk.Button(SwitchPortsWindow, width=15, text="New Entry", bg='white', comman=lambda: newEntry(SIWAsset_ID))
    btn_delete = tk.Button(SwitchPortsWindow, width=15, text="Delete Entry", bg='white', command=delete)
    btn_connectedMACs = tk.Button(SwitchPortsWindow, width=15, text="View Conn.", bg='white',
                                  command=lambda: connectedMACs(SwitchAsset_ID, SwitchPorts_ID))
    btn_misconfigs = tk.Button(SwitchPortsWindow, width=15, text="View Misconfigurations", bg='white',
                               command=lambda: misconfigs(SwitchAsset_ID, SwitchPorts_ID))

    btn_home.grid(row=1, column=3, pady=5, padx=5, sticky='w')
    btn_back.grid(row=1, column=4, pady=5, padx=5, sticky='w')
    btn_new.grid(row=1, column=0, pady=5, padx=5, sticky='w')
    btn_edit.grid(row=1, column=1, pady=5, padx=5, sticky='w')
    btn_delete.grid(row=1, column=2, pady=5, padx=5, sticky='w')
    btn_connectedMACs.grid(row=1, column=5, pady=5, padx=5, sticky='w')
    btn_misconfigs.grid(row=1, column=6, pady=5, padx=5, sticky='w')

    tree = ttk.Treeview(SwitchPortsWindow, column=('#1', '#2', '#3', '#4', '#5', '#6', '#7'), show='headings')
    tree.column('#1', width=30)
    tree.bind("<<TreeviewSelect>>", updateWidgets)
    tree.heading('#1', text='ID')
    tree.heading('#2', text='Port Number')
    tree.heading('#3', text='Description')
    tree.heading('#4', text='Admin Status')
    tree.heading('#5', text='Port Type')
    tree.heading('#6', text='Native VLAN')
    tree.heading('#7', text='VLANTag')

    conn = pyodbc.connect(cnxn)
    cursor = conn.cursor()
    SQL = """SELECT DISTINCT
                tbl_SwitchPorts.SwitchPorts_ID,
                tbl_SwitchPorts.PortNumber,
                tbl_SwitchPorts.Description,
                tbl_SwitchPorts.AdminStatus,
                tbl_SwitchPorts.PortType,
                tbl_SwitchPorts.NativeVLAN,
                tbl_SwitchPorts.VLANTag
                FROM tbl_SwitchPorts
                WHERE tbl_SwitchPorts.SwitchAsset_ID = {0} AND tbl_SwitchPorts.Deleted IS NULL""".format(SwitchAsset_ID)
    cursor.execute(SQL)
    records = cursor.fetchall()
    if len(records) == 0:
        messagebox.showinfo("Oops!", "No Records Found")
    else:
        count = 0
        for record in records:
            tree.insert('', tk.END, iid=count,
                        values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6]))
            count += 1

    tree.grid(row=0, column=0, columnspan=41)
    if len(records) > 0:
        tree.selection_set(0)
        tree.focus(0)


# -----------------------------------START---------------------------------------------

SignIn()