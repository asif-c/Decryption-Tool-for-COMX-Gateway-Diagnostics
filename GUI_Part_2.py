import tkinter as tk
from tkinter import messagebox, filedialog
import tkSimpleDialog
import os
from subprocess import Popen,PIPE
import sys
import subprocess
import re
import os
import os.path
import datetime
# following module is needed to move (not copy) the decrypted file to the source location
import shutil
from shutil import copy2

#***********global variables*********
file_name_r=""
complete_output_path=""

#*************Class and functions related to Menubar*********

'''
    In the following link, the code of tkSimpledialog can be found:
    http://effbot.org/tkinterbook/tkinter-dialog-windows.htm
    To create dialogbox, it is suggested to create a class that inherits the Dialog class and
    overwrite the body and apply method to suit the specific need

'''

class MyDialog(tkSimpleDialog.Dialog):

    def body(self,root):
        # tk.Label(root,text="Set minutes\n(only integer, no negative or fraction)").grid(row=0,column=0)

        self.e1=tk.Entry(root)
        self.e1.grid(row=0,column=1)

        return self.e1

def about_this_app():

    tk.messagebox.showinfo("About this app","This simple desktop app decrypts the \nencrypted diagnostic file from a COM'X \nnot needing user to type\ncomplex and finicky commandline arguments.\n\nNote:\nThis app runs only on the Windows machine.")

def command_the_creator_of_this_app():

    tk.messagebox.showinfo("Contact the creator of the app","Asif Choudhury\n\nEmail:\n(Work)\tasif.choudhury@schneider-electric.com,\n(Personal) asifikchoudhury@gmail.com\nLocation:\tVictoria, BC, Canada\n\nPlease contact the creator of the app, if you have any feedback.")


#***********functions related to buttons***********

def open_decrypted_file_location():
    global complete_output_path

    # Now, to convert Python/Linux like directory path (C:/home/asif) to Windows-like (C:\home\asif)
    # you need os.path.normpath functions
    complete_output_path_windows=os.path.normpath(complete_output_path)
    # print(complete_output_path_windows)
    theproc2=Popen(r'explorer /select, {0}'.format(complete_output_path_windows))
    theproc2.communicate()

def browse_select_file():
    global file_name_r
    file_name_r=filedialog.askopenfilename()
    decrypt_button.config(state=tk.NORMAL)

def decryption():

    global complete_output_path

    decrypt_button.config(state=tk.DISABLED)
    decrypted_file_location.config(state=tk.DISABLED)
    try:
        file_name = os.path.basename(file_name_r)
        # The output of following code will be C:/Users/sesa196358/PycharmProjects/Decrypt_COMX_Diagnostics_2
        # Note the forward slash
        file_path=os.path.dirname(file_name_r)
        # The output of the following code will be C:\Users\sesa196358\PycharmProjects\Decrypt_COMX_Diagnostics_2
        #  Note the backslash in the directory path
        python_code_path=os.getcwd()
        # The following is there only for testing
        # file_name_r_windows = os.path.normpath(file_name_r)
        # print("file name={0},\nfile path={1},\npython code path={2}".format(file_name,file_path,python_code_path))
        # moving the OneEsp file to python code location
        # Instead just the fil name (OneEsp..), input the entire thing: file_name_r
        file_path_windows=os.path.normpath(file_path)

        # The following code checks whether the source and destination directories are same or not
        # Move the OneEsp file only if the source and destination folders are different

        if (file_path_windows != python_code_path):
            shutil.copy(file_name_r,python_code_path)

        # regex pattern
        if file_name.startswith("OneESP_"):
            '''
            Here goes the code that sucks out any space from the OneESP file. Why?
            OneESP_diags_DN13266SE000200_20181018_1130_diagnosticfile WILL WORK
            OneESP_diags_DN13266SE000200_20181018_1130_ diagnostic file WILL **NOT** WORK
            So, basically you are renaming the copied file removing any space from it
            '''
            if re.search(r'\s',file_name):
                file_name_no_space=file_name.replace(" ","")
            # first remove/delete a file if it already exists
            #     os.remove(file_name_no_space)
            # Now rename the copied file
                os.rename(file_name,file_name_no_space)
            else:
                file_name_no_space=file_name
            pattern = re.compile(r"(?=D)[a-zA-Z0-9]+")
            # do some regex here to extract the DN part from the encrypted OneESP file name

            dn_part_list = re.findall(pattern, file_name_no_space)
            # recall re.findall() outputs a list like this ['DN18222SE000022']
            # The following line will transform ['DN18222SE000022'] to 'DN18222SE000022'
            dn_part = "".join(x for x in dn_part_list)

            # %Y%h%d_%H%M%S= %Y=>2018, %h=>Dec,Nov, %d=Date, %H=hour of the day in 24-hour format,
            # %M=minute of the hour, %S=Second in the 60 second
            # No milisecond added in the output file name to make the name uncluttered
            current_time = datetime.datetime.now().strftime("%Y%h%d_%H%M%S")
            destination_zipFile_name = "Decrypted_diagnostics_{0}_{1}.zip".format(dn_part, current_time)

            # print(dn_part)
            complete_output_path="{0}/{1}".format(file_path,destination_zipFile_name)
            # The following line is there only for testing
            # print(complete_output_path)
            cmd_command = "openssl enc -d -aes-128-cbc -nosalt -pass pass:{0} -in {1} -out {2}".format(dn_part, file_name_no_space,destination_zipFile_name)

            theproc1 = Popen(cmd_command, shell=True,stdout=PIPE, stderr=PIPE,stdin=PIPE)
            # theproc1=subprocess.Popen('explorer /select, {0}'.format(cmd_command))
            child_stdout,child_stderr=theproc1.communicate()

        else:
            tk.messagebox.showinfo("Error Message","Please select the correct diagnostic file.\nIt has the following format:\nOneESP_diags_DNxxxxxSExxxxxx_xxxxxxxx_xxxx.octet-stream")
            os.remove(file_name)
            # file_name=""

    except Exception as e:
        tk.messagebox.showinfo("Error Message", e)

    else:
        '''
        The following piece of code will only kick in if the decryption succeeds.
        In the following piece of code the decrypted file moves from software's base directory to
        the folder of origin
        '''
        slash="\\"
        new_src=slash.join((python_code_path,destination_zipFile_name))
        # Only if the source and destination directories are different the following codes will kick in
        if (file_path_windows != python_code_path):
            shutil.move(new_src, file_path)
        decrypted_file_location.config(state=tk.NORMAL)
    finally:
        # Only if the source and destination directories are different the following codes will kick in
        if (file_path_windows != python_code_path):
            os.remove(file_name_no_space)
        file_name=""
        decrypt_button.config(state=tk.DISABLED)

root=tk.Tk()

#*******Menubar*********************
menubar=tk.Menu(root)
root.config(menu=menubar)
# helpmenu
helpmenu=tk.Menu(menubar,tearoff=0)
helpmenu.add_cascade(label="About this app",command=about_this_app)
helpmenu.add_separator()
helpmenu.add_cascade(label="Contact the creator",command=command_the_creator_of_this_app)
helpmenu.add_separator()
helpmenu.add_cascade(label="Exit",command=root.quit)

# menubar.add_cascade(label="File",menu=filemenu)
menubar.add_cascade(label="Help",menu=helpmenu)

#Main display area and buttons
lbl1=tk.Label(root,text="Diagnostic file location",font=("Helvetica",20),fg='black',padx=10,pady=10)
lbl1.grid(row=0,column=0)

# Browse Button
browse_button=tk.Button(root,text="Browse",font=("Helvetica",20),fg='black',padx=10,pady=10,command=browse_select_file )
browse_button.grid(row=0,column=1)

#Decrypt button
decrypt_button=tk.Button(root,text="Decrypt",font=("Helvetica",20),fg='black',command=decryption)
decrypt_button.grid(row=1,columnspan=2)
decrypt_button.config(state=tk.DISABLED)

#Decrypted File Location button
decrypted_file_location=tk.Button(root, text="Decrypted file location",font=("Helvetica",20),fg='black', command=open_decrypted_file_location )
decrypted_file_location.grid(row=2,columnspan=2)
decrypted_file_location.config(state=tk.DISABLED)


#******************Window Title Bar**************
root.title("Decrypting COM'X Diagnostics 2.0.1")

#******************Window Title Bar Icon*********
# root.call('wm', 'iconbitmap', root._w, '-default', 'pmodoro_icon.ico')

root.mainloop()