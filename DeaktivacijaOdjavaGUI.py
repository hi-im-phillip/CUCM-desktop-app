from tkinter import *
from tkinter import messagebox
import ssl
import urllib
import requests
import suds
import urllib3
from suds.client import Client
import xml.etree.cElementTree as ET
from tkinter import ttk
import pyperclip
import pyautogui as pya
import time

root = Tk()
root.title("CUCM Functions")
root.resizable(0, 0)

style = ttk.Style()
style.configure("BW.TLabel", foreground="black", background="white")

style.map("C.TButton",
          foreground=[('pressed', 'red'), ('active', 'blue')],
          background=[('pressed', '!disabled', 'black'), ('active', 'white')]
          )

ssl._create_default_https_context = ssl._create_unverified_context


username = 'username'
password = 'password'
cucm = 'http://0.0.0.0//axl/'
wsdl = 'file:///D:/axlsqltoolkit/schema/11.5/AXLAPI.wsdl'


# vraca sve brojeve koji se ne koriste u DirectoryNumber-u
def difference_list(extension_range, clean_list_numbers):
    return [item for item in extension_range if item not in clean_list_numbers]


# vraca sve brojeve koji nisu upisani u listu shared_line_extenstion
def shared_line(free_agent_extensions, shared_line_exstension):
    return [item for item in free_agent_extensions if item not in shared_line_exstension]


# vraca sve brojeve koji nisu upisani u listu reserved_line
def reserved_line(shared_line_extensions, reserved_line_extensions):
    return [item for item in shared_line_extensions if item not in reserved_line_extensions]


# CCM


def get_ccm_version(client):
    return client.service.getCCMVersion


# DIRECTORY NUMBER


def get_directory_number_by_number(client, pattern, routePartitionName):
    return client.service.getLine(pattern=pattern, routePartitionName=routePartitionName)


def get_directory_number_by_description(client, description, routePartitionName):
    return client.service.GetLineReq(description=description, routePartitionName=routePartitionName)


def remove_directory_number_by_number(client, pattern, routePartitionName):
    return client.service.removeLine(pattern=pattern, routePartitionName=routePartitionName)


def list_directory_number(client, pattern):
    return client.service.listLine(pattern=pattern)


# END USER PROFILE


def get_end_user_by_name(client, name):
    return client.service.getUser(userid=name)


def remove_end_user_by_userid(client, name):
    return client.service.removeUser(userid=name)


# DEVICE PROFILE


def get_device_by_name(client, name):
    return client.service.getDeviceProfile(name=name)


def remove_device_by_name(client, name):
    return client.service.removeDeviceProfile(name=name)


def list_listDeviceProfile(client, name):
    return client.service.listDeviceProfile(name=name)


# PHONE PROFILE


def get_phone_by_name(client, name):
    return client.service.getPhone(name=name)


def remove_phone_by_name(client, name):
    return client.service.removePhone(name=name)


def update_phone_by_name(client, name, description):
    return client.service.updatePhone(**{'name': name, 'description': description})


def list_phones(client):
    return client.service.listPhone


def ask_user():
    user_input = str(input("Opcije: \n 1. Get device name \n 2. Get phone \n"))
    return user_input


def directory_number_pull_ph():
    login_url = 'https://0.0.0.0/axl/'
    cucm_version_actions = 'CUCM:DB ver=11.5 listLine'
    username = 'username'
    password = r'password'
    soap_data = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" ' \
                'xmlns:ns="http://www.cisco.com/AXL/API/8.5"><soapenv:Header/><soapenv:Body>' \
                '<ns:listLine><searchCriteria><description>%%</description></searchCriteria><returnedTags>' \
                '<pattern></pattern><description></description></returnedTags></ns' \
                ':listLine></soapenv:Body></soapenv:Envelope> '
    soap_headers = {'Content-type': 'text/xml', 'SOAPAction': cucm_version_actions}

    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    requests.packages.urllib3.util.ssl_.DEFAULT_CIPHERS += 'HIGH:!DH:!aNULL'

    try:
        requests.packages.urllib3.contrib.pyopenssl.DEFAULT_SSL_CIPHERS_LIST += 'HIGH:!DH:!aNULL'
    except AttributeError:
        # nije pyopenssl support koristena / potrebna / moguca
        pass

    try:
        axl_request = requests.post(login_url, data=soap_data, headers=soap_headers, verify=False,
                                    auth=(username, password))
    except ConnectionError:
        print("Za prikaz slobodnih ekstenzija potreban je VPN")
    plain_txt = axl_request.text
    root = ET.fromstring(plain_txt)

    list_numbers = []
    extension_range_zagreb = list(range(8500, 8599)) + list(range(8700, 8800)) + [4111, 4801, 4804]
    shared_line_exstension_zagreb = []
    reserved_line_exstenion_zagreb = [8570, 8579, 8581, 8583, 8590, 8595, 8598, 8700, 8740, 8750, 8760, 8770, 8780,
                                      8790]

    extension_range_split = list(range(8600, 8700))
    shared_line_exstension_split = [8603, 8607, 8615]

    for device in root.iter('line'):
        list_numbers.append(device.find('pattern').text)
        # Za prikaz kome broj pripada
        # list_description.append(device.find('description').text)

    list_without_quotes = []

    for line in list_numbers:
        line = line.replace("'", "").replace("*", "")
        list_without_quotes.append(line)

    clean_list_numbers = [int(i) for i in list_without_quotes]

    # nema rezerviranih brojeva u splitu
    free_agent_extensions_split = difference_list(extension_range_split, clean_list_numbers)
    shared_line_numbers_split = shared_line(free_agent_extensions_split, shared_line_exstension_split)

    free_agent_extensions_zagreb = difference_list(extension_range_zagreb, clean_list_numbers)
    shared_line_numbers_zagreb = shared_line(free_agent_extensions_zagreb, shared_line_exstension_zagreb)
    reserver_line_numbers_zagreb = reserved_line(shared_line_numbers_zagreb, reserved_line_exstenion_zagreb)

    return reserver_line_numbers_zagreb + ' Split: ' + shared_line_exstension_split


def directory_number_pull_isk():
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    login_url = 'https://0.0.0.0/axl/'
    cucm_version_actions = 'CUCM:DB ver=11.5 listLine'
    username = 'username'
    password = r'password'
    soap_data = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" ' \
                'xmlns:ns="http://www.cisco.com/AXL/API/11.5"><soapenv:Header/><soapenv:Body>' \
                '<ns:listLine><searchCriteria><description>%%</description></searchCriteria><returnedTags>' \
                '<pattern></pattern><description></description></returnedTags></ns' \
                ':listLine></soapenv:Body></soapenv:Envelope> '
    soap_headers = {'Content-type': 'text/xml', 'SOAPAction': cucm_version_actions}

    axl_request = requests.post(login_url, data=soap_data, headers=soap_headers, verify=False,
                                auth=(username, password))
    plain_txt = axl_request.text
    root = ET.fromstring(plain_txt)

    list_numbers = []
    extension_range = list(range(100, 599)) + [701, 802, 839, 847, 857, 876, 877, 878, 881, 927, 967]
    shared_line_exstension = [130, 170]
    reserved_line_exstenion = [108, 150, 205, 259, 300, 320, 330, 334, 394, 398, 399, 480, 481, 482, 483, 484, 485,
                               486, 487, 488, 489, 570, 571, 573, 574, 575, 577, 578, 579, 580, 590, 702]

    for device in root.iter('line'):
        list_numbers.append(device.find('pattern').text)
        # Za prikaz kome broj pripada
        # list_description.append(device.find('description').text)

    list_without_quotes = []

    for line in list_numbers:
        line = line.replace("'", "").replace("*", "")
        list_without_quotes.append(line)

    clean_list_numbers = [int(i) for i in list_without_quotes]

    free_agent_extensions = difference_list(extension_range, clean_list_numbers)
    shared_line_numbers = shared_line(free_agent_extensions, shared_line_exstension)
    reserver_line_numbers = reserved_line(shared_line_numbers, reserved_line_exstenion)

    return reserver_line_numbers


def deactivation_agent_isk_GUI_2(first_last_name, username):
    while True:
        try:
            route_partition_name_isk= 'PT_isk_Internal'
            device_profile_isk = ' Device Profile 6941'
            cucm_server = Client(wsdl, location=cucm, username=username, password=password)
            device_profile = get_device_by_name(cucm_server, first_last_name + device_profile_isk)
            if device_profile is None:
                print("Ne postoji korisnik pod imenom " + first_last_name)
            device_number = (int(device_profile['return']['deviceProfile']['lines']['line'][0]['dirn']['pattern']))
            if device_number is None:
                print("Ne postoji broj od " + first_last_name)
            remove_device_by_name(cucm_server, first_last_name + device_profile_isk)
            remove_end_user_by_userid(cucm_server, username)
            remove_directory_number_by_number(cucm_server, device_number, route_partition_name_isk)
        except suds.transport.TransportError:
            print("Korisnik " + first_last_name + " nije pronaden")
        except urllib.error.HTTPError:
            print("HTTP error")
        except suds.WebFault:
            print("Korinsik " + first_last_name + " ne postoji")


def deactivation_agent_ph_GUI(first_last_name, username):
    # while True:
    # try:
    route_partition_name_ph = 'PT_Internal'
    device_profile_ph = ' Device Profile'
    cucm_server = Client(wsdl, location=cucm, username=username, password=password)
    device_profile = get_device_by_name(cucm_server, first_last_name + device_profile_ph)
    if device_profile is None:
        messagebox.showwarning("Obavijest", "Ne postoji korisnik pod imenom" + first_last_name)
        # break
    device_number = (int(device_profile['return']['deviceProfile']['lines']['line'][0]['dirn']['pattern']))
    if device_number is None:
        messagebox.showwarning("Obavijest", "Ne postoji broj od " + first_last_name)
        # break
    remove_device_by_name(cucm_server, first_last_name + device_profile_ph)
    remove_end_user_by_userid(cucm_server, username)
    remove_directory_number_by_number(cucm_server, device_number, route_partition_name_ph)


# except suds.transport.TransportError:
# messagebox.showwarning("Obavijest", "Korisnik " + str(first_last_name) + " nije pronaden")
# break
# except urllib.error.HTTPError:
# messagebox.showwarning("Obavijest", "HTTP error")
# break
# except suds.WebFault:
#    messagebox.showwarning("Obavijest", "Korisnik " + str(first_last_name) + " ne postoji")
#    break


def deactivation_agent_isk_GUI(first_last_name, username):
    # while True:
    # try:
    route_partition_name_isk = 'PT_Isk_Internal'
    device_profile_isk = ' Device Profile 6941'
    cucm_server = Client(wsdl, location=cucm, username=username,
                         password=password)
    device_profile = get_device_by_name(cucm_server, first_last_name + device_profile_isk)
    if device_profile is None:
        messagebox.showwarning("Obavijest", "Ne postoji korisnik pod imenom " + first_last_name)
    #    break
    device_number = (int(device_profile['return']['deviceProfile']['lines']['line'][0]['dirn']['pattern']))
    if device_number is None:
        messagebox.showwarning("Obavijest", "Ne postoji broj od " + first_last_name)
    #    break
    remove_device_by_name(cucm_server, first_last_name + device_profile_isk)
    remove_end_user_by_userid(cucm_server, username)
    remove_directory_number_by_number(cucm_server, device_number, route_partition_name_isk)


# except suds.transport.TransportError:
#    messagebox.showwarning("Obavijest", "Korisnik " + str(first_last_name) + " nije pronaden")
#   break
# except urllib.error.HTTPError:
#    messagebox.showwarning("Obavijest", "HTTP error")
#   break
# except suds.WebFault:
#   messagebox.showwarning("Obavijest", "Korinsik " + str(first_last_name) + " ne postoji")
#   break


def deactivation_agent_ph_old():
    while True:
        try:
            route_partition_name_ph = 'PT_Internal'
            device_profile_ph = ' Device Profile'
            cucm_server = Client(wsdl, location=cucm, username=username, password=password)
            user_first_last_name = str(input("Unesi ime i prezime agenta: "))
            user_username = str(input("Unesi korisničko ime agenta (username): "))

            device = get_device_by_name(cucm_server, user_first_last_name + device_profile_ph)
            device_number = (int(device['return']['deviceProfile']['lines']['line'][0]['dirn']['pattern']))
            if device_number is None:
                print("Ne postoji broj kod traženog agenta!")
            remove_device_by_name(cucm_server, user_first_last_name + device_profile_ph)
            remove_end_user_by_userid(cucm_server, user_username)
            remove_directory_number_by_number(cucm_server, device_number, route_partition_name_ph)
            print("Uspješno odjavljen agent " + user_first_last_name)
        except suds.transport.TransportError:
            print("Agent " + user_first_last_name + " nije pronađen")
        except urllib.error.HTTPError:
            print("HTTP error")
        except suds.WebFault:
            print("Agent " + user_first_last_name + " ne postoji")


def show_isk_extensions():
    list1.delete(0.0, END)
    try:
        d = directory_number_pull_isk()
    except ConnectionError:
        messagebox.showwarning("Obavijest", "Problem kod konekcije, provjeri VPN")
    list1.insert(END, d)


def show_ph_extensions():
    list1.delete(0.0, END)
    try:
        t = directory_number_pull_ph()
    except ConnectionError:
        messagebox.showwarning("Obavijest", "Problem kod konekcije, provjeri VPN")
    list1.insert(END, t)


def deactivation_agent_isk():
    if len(first_last_name_text.get()) >= 5 and len(username_text.get()) >= 3:
        first_last_entry = str(first_last_name_text.get())
        username_entry = str(username_text.get())
        deactivation_agent_isk_GUI(first_last_entry, username_entry)
        messagebox.showwarning("Obavijest", "Uspješno odjavljen korisnik " + first_last_entry)
        e1.delete(0, END)
        e2.delete(0, END)
    else:
        messagebox.showwarning("Obavijest", "Naziv agenta ili Username nije upisan!")


def deactivation_agent_ph():
    if len(first_last_name_text.get()) >= 5 and len(username_text.get()) >= 3:
        first_last_entry = str(first_last_name_text.get())
        username_entry = str(username_text.get())
        deactivation_agent_ph_GUI(first_last_entry, username_entry)
        messagebox.showinfo("Obavijest", "Uspješno odjavljen korisnik " + first_last_entry)
        e1.delete(0, END)
        e2.delete(0, END)
    else:
        messagebox.showwarning("Obavijest", "Naziv agenta ili Username nije upisan!")


# radi!
def radi_li_textbox():
    entry_test = first_last_name_text.get()
    print(entry_test)
    e1.delete(0, END)


def leftClick(event):
    print("Left")


def middleClick(event):
    print("Middle")


#def copy_text_to_clipboard(event):
 #   field_value = event.widget.get("1.0", 'end-1c')  # get field value from event, but remove line return at end
  #  root.clipboard_clear()  # clear clipboard contents
   # root.clipboard_append(field_value)  # append new value to clipbaord


def rightClick(event):
    print("Right")


def popup(event):
    try:
        popup_menu.tk_popup(event.x_root, event.y_root, 0)
    finally:
        popup_menu.grab_release()


def copy_clipboard():
    pya.hotkey('ctrl', 'c')
    time.sleep(.01)  # ctrl-c is usually very fast but your program may execute faster
    return pyperclip.paste()


pya.doubleClick(pya.position())


def copy_highlighted():
    list = []
    var = copy_clipboard()
    list.append(var)
    return list


popup_menu = Menu(tearoff=0)
popup_menu.add_command(label="Copy", command=copy_highlighted)

# root.bind("<Button-1>", leftClick)
# root.bind("<Button-2>", middleClick)
root.bind("<Button-3>", popup)
root.grid()

lbl1 = ttk.Label(root, text="Ime i prezime:")
lbl1.grid(row=0, sticky=W)

lbl2 = ttk.Label(root, text="Korisničko ime:")
lbl2.grid(row=1, sticky=W)

first_last_name_text = StringVar()
e1 = Entry(root, textvariable=first_last_name_text)
e1.grid(row=0, column=1, ipadx=37, sticky=W)
e1.focus()

username_text = StringVar()
e2 = Entry(root, textvariable=username_text)
e2.grid(row=1, column=1, ipadx=37, sticky=W)

list1 = Text(root, height=6, width=35)
list1.grid(row=2, column=0, rowspan=8, columnspan=2)

sb1 = Scrollbar(root)
sb1.grid(row=2, column=2, rowspan=6)

list1.configure(yscrollcommand=sb1.set)
sb1.configure(command=list1.yview)

b1 = ttk.Button(root, text="Odjaviti isk agenta", width=25, style="C.TButton", command=deactivation_agent_isk)
b1.grid(row=2, column=3)

b2 = ttk.Button(root, text="Odjaviti ph agenta", width=25, style="C.TButton", command=deactivation_agent_ph)
b2.grid(row=3, column=3)

b3 = ttk.Button(root, text="Slobodne ekstenzije isk", width=25, style="C.TButton", command=show_isk_extensions)
b3.grid(row=4, column=3)

b4 = ttk.Button(root, text="Slobodne ekstenzije ph", width=25, style="C.TButton", command=show_ph_extensions)
b4.grid(row=5, column=3)

b5 = ttk.Button(root, text="Zatvoriti", width=25, style="C.TButton", command=root.destroy)
b5.grid(row=6, column=3)

root.mainloop()
