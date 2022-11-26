from PyQt5 import QtGui
from PyQt5.QtGui import *
from PyQt5.QtGui import QPixmap
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QDialog, qApp, QFileDialog 
from PyQt5.QtCore import *
from PyQt5.uic import loadUi
import MySQLdb
from PIL import Image
from PIL.ImageQt import ImageQt
import base64
import io
import sys
import datetime
from xlsxwriter import *
from xlrd import *
import xlwt
from xlwt import *
from email_sender import send_mail_with_excel
from dateutil import relativedelta

class FirstScreen(QDialog):
    def __init__(self):
        super(FirstScreen, self).__init__()
        loadUi("firstpage.ui", self)

        self.Handel_All_Buttons()

    def Handel_All_Buttons(self):
        self.signin.clicked.connect(self.Main_Page)
        self.exitprogram.clicked.connect(self.Exit_Program_Func)

    def Main_Page(self):
        login_page = MainScreen()
        widget.setFixedHeight(728)
        widget.setFixedWidth(1210)
        widget.addWidget(login_page)
        widget.setCurrentIndex(widget.currentIndex()+1)

    def Exit_Program_Func(self):
        sys.exit()

class MainScreen(QDialog):
    def __init__(self):
        global today_date
        super(MainScreen, self).__init__()
        loadUi("index2.ui", self)

        self.Db_Connect()
        self.Hide_Password()
        self.Show_Tab_Changes()
        self.Handell_All_Buttons()
        self.Today_Date()
        self.Show_User_Name()

        self.Active_Client_Type()

        self.Login_Tab()
        self.Show_Call_Type()
        self.Show_Calls_History()
        self.Show_Technician_History()
        self.Tech_Combo_Activ()

        self.Activate_Combo_Dele()

        self.Show_Machines_Type()
        self.Activate_Combo()
        self.Show_Client()
        self.Show_Contact()
        self.Show_Branch()
        self.Show_Machine()

        self.Show_Monthly_Followup()
        self.Clear_Search_Tech_Monthly()
        self.Clear_Search_Monthly_Todelete()
        self.Show_Clname_Spare()
        self.Show_Clname_Cleaners()

    def Db_Connect(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='password',  port=3306, db='db-name')
        self.cur = self.db.cursor()

    def Show_Tab_Changes(self):
        self.tabWidget.tabBar().setVisible(False)

    def Handell_All_Buttons(self):
        self.loginbutton.clicked.connect(self.Login_User)
        self.createaccount.clicked.connect(self.Create_User)
        self.loginshowpass.clicked.connect(self.Show_Password_Checkbox)
        self.logoutbutton.clicked.connect(self.Login_Tab)

        self.createaccountprofile.clicked.connect(self.Add_Account)
        self.cancelcreatebutton.clicked.connect(self.Login_Tab)
        self.cancelresetpassword.clicked.connect(self.Cancel_Reset_Password)

        self.openresetpassword.clicked.connect(self.Open_Reset_Password)
        self.savenewpassword.clicked.connect(self.Save_New_Password)

        self.saveprofilebutton.clicked.connect(self.Save_User_Profile)
        self.cancelprofbutton.clicked.connect(self.Clear_Profile)
        self.upimagebutton.clicked.connect(self.Upload_Image)
        self.searchprofile.clicked.connect(self.Edit_Profile)
        self.addpermissionsprofile.clicked.connect(self.Add_User_Permission)
        self.deleteprofile.clicked.connect(self.Delete_User_Profile)

        self.executecallbutton.clicked.connect(self.Handel_Calls)
        self.cancelcall.clicked.connect(self.Clear_Calls_Fields)

        self.searchhistory.clicked.connect(self.Search_Calls_History)
        self.clearsearchhistory.clicked.connect(self.Clear_Search_History)
        self.deleteexecutedcall.clicked.connect(self.Delete_Executed_Call)
        self.exportdailycalls.clicked.connect(self.Export_Dailly_Calls)

        self.canceldelcall.clicked.connect(self.Edit_Del_Calls)
        self.deleteexecutedcall_2.clicked.connect(self.Empty_Labels)
        self.searchbutdel.clicked.connect(self.Fill_Fields)

        self.dailymovebutton.clicked.connect(self.Daily_Move_Tab)
        self.editdelcalls.clicked.connect(self.Edit_Del_Calls)
        self.techbutton.clicked.connect(self.Technician_Tab)
        self.clientbutton.clicked.connect(self.Clients_Tab)
        self.monthlyfollowupbutton.clicked.connect(self.Clients_Details_Tab)
        self.createuserbutt.clicked.connect(self.Create_User)
        self.openprofuser.clicked.connect(self.Open_Profile)

        self.sendwahts.clicked.connect(self.Whats_Message)

        self.sendemails.clicked.connect(self.Send_Email_Page)
        self.sendemail.clicked.connect(self.Email_To_Send)

        self.searchtechn.clicked.connect(self.Search_Technician_Jobs)
        self.clearseatechist.clicked.connect(self.Clear_Technician_Jobs)
        self.exporttechcallsbutt.clicked.connect(self.Export_Tech_Calls)

        self.newclientbutton.clicked.connect(self.Add_New_Client)
        self.searchclientbutton.clicked.connect(self.Search_Client_Name)
        self.clearsearchbutton.clicked.connect(self.Clear_Search_Client)
        self.editeclientbutton.clicked.connect(self.Edit_Client_Information)
        self.deleteclientbutton.clicked.connect(self.Delete_Client)

        self.addcontactbutton.clicked.connect(self.Add_Client_Contact)
        self.searchcontactbutton.clicked.connect(self.Search_Contact)
        self.clearseaconbutton.clicked.connect(self.Clear_Search_Contact)
        self.editcontactbutton.clicked.connect(self.Eidt_Contact_Name)
        self.deletecontactbutton.clicked.connect(self.Delete_Contact)

        self.addbranchbutton.clicked.connect(self.Add_New_Branch)
        self.searchbrabutton.clicked.connect(self.Search_Branch)
        self.editebrabutton.clicked.connect(self.Edit_Client_Branch)
        self.clearseabrabutton.clicked.connect(self.Clear_Search_Branch)
        self.deletebranchbutton.clicked.connect(self.Delete_Branch)

        self.addmachinebutton.clicked.connect(self.Add_Client_Machine)
        self.searchmachinebutton.clicked.connect(self.Search_Bytype_Machine)
        self.editmachinebutton.clicked.connect(self.Edit_Client_Machine)
        self.clearseaedimabutton_2.clicked.connect(self.Clear_Search_Machine)
        self.deletemacbutton.clicked.connect(self.Delete_Machine)

        self.addbranchtolist.clicked.connect(self.Add_Monthly_Followup)
        self.cancelmonthlybutt.clicked.connect(self.Clear_Selection_Monthly)
        self.searchtechn_2.clicked.connect(self.Search_Tech_Monthly)
        self.clearseatechist_2.clicked.connect(self.Clear_Search_Tech_Monthly)
        self.exportmonthlyrecord.clicked.connect(self.Export_Monthly_Record)
        self.searchtechn_3.clicked.connect(self.Show_Monthly_Record_Todelete)
        self.clearseatechist_3.clicked.connect(self.Clear_Search_Monthly_Todelete)
        self.deletemonthlyrecord.clicked.connect(self.Delete_Monthly_Follow)

        self.logoutbutton.clicked.connect(self.Login_Tab)
        self.exitprogbutton.clicked.connect(self.Exit_Program_Func)

    def Login_User(self):
        user = self.loginuser.text()
        log_password = self.loginpassword.text()
        login_date = datetime.date.today()
        now = datetime.datetime.now()
        login_time = datetime.time(now.hour, now.minute, now.second)

        if len(user)== 0 or len(log_password)==0:
            self.loginerror.setText("Invalid user name or password")
        else:
            self.loginerror.clear()
            self.cur.execute('''SELECT user_name, password FROM users WHERE user_code=user_code''')
            all_data = self.cur.fetchall()
            if all_data==():
                self.loginerror.setText("Invalid user name or password")
                self.loginuser.clear()
                self.loginpassword.clear()
            else:
                for row in all_data:
                    if row[0]!=user or row[1]!=log_password:
                        self.loginerror.setText("Invalid user name or password")
                    else:
                        self.cur.execute('''
                            INSERT INTO login_history(user_name, login_date, login_time)
                            VALUES(%s, %s, %s)
                        ''', (user, login_date, login_time))
                        self.db.commit()
                        self.loginerror.clear()
                        self.Show_User_Login()
                        self.Daily_Move_Tab()
                        self.Clear_Fields()

    def Hide_Password(self):
        self.loginpassword.setEchoMode(QtWidgets.QLineEdit.Password)

    def Show_Password_Checkbox(self):
        if self.loginshowpass.isChecked()==True:
            self.loginpassword.setEchoMode(QtWidgets.QLineEdit.Normal)
        else:
            self.loginpassword.setEchoMode(QtWidgets.QLineEdit.Password)
    
    def Clear_Fields(self):
        self.loginuser.clear()
        self.loginpassword.clear()

    def Add_Account(self):
        cr_user = self.createusername.text()
        cr_code = self.createusercode.text()
        cr_proff = self.createcoboxprof.currentText()
        cr_password = self.createpassword.text()
        crf_password = self.createcfpassword.text()
        self.cur.execute('''
            SELECT user_name FROM users
        ''')
        all_names = self.cur.fetchall()
        name_list = []
        for names in all_names:
            for name in names:
                name_list.append(name)
        self.cur.execute('''
            SELECT user_code FROM users
        ''')
        all_codes = self.cur.fetchall()
        result = []
        for code in all_codes:
            for i in code:
                result.append(i)
        if len(cr_user)==0 or len(cr_code)==0 or len(cr_password)==0:
            self.createerror.setText("Please fill in all Fields!")
        elif cr_user in name_list:
            self.createerror.setText("User Name Already exists!")
        elif int(cr_code) in result:
            self.createerror.setText("User Code Already exists!")
        elif cr_password!=crf_password:
            self.createerror.setText("Passwords do not match!")
        else:
            self.cur.execute('''
                INSERT INTO users (user_name, user_code, user_professional, password)
                VALUES(%s, %s, %s, %s)
            ''', (cr_user, cr_code, cr_proff, cr_password))
            self.db.commit()
            self.Login_Tab()
            self.Clear_Adduser_Fields()
    
    def Clear_Adduser_Fields(self):
        self.createusername.clear()
        self.createusercode.clear()
        self.createpassword.clear()
        self.createcfpassword.clear()

    def Save_New_Password(self):
        user_name=self.usernameresetpass.text()
        user_password = self.currentpasswordresetpass.text()
        new_pass = self.newpasswordreset.text()
        conf_newpass = self.confirmnewpasswordreset.text()
        if len(user_name)==0 and len(user_password)==0:
            self.errormessageresetpass.setText("Invalid User Name or Current Password!")
        else:
            self.cur.execute('''
                SELECT user_name, password FROM users
            ''')
            names = self.cur.fetchall()
            use_names = []
            use_pass = []
            for user_info in names:
                use_names.append(user_info[0])
                use_pass.append(user_info[1])
            if user_name not in use_names or user_password not in use_pass:
                self.errormessageresetpass.setText("User Name or Current Password Does Not Existe!")
            elif new_pass!=conf_newpass:
                self.errormessageresetpass.setText("New Password and Confirm Password Do Not Match!")
            else:
                self.cur.execute('''
                    UPDATE users SET password=%s WHERE user_name=%s
                        ''', (new_pass, user_name))
                self.db.commit()
                QMessageBox.information(self, 'Success', 'Password has been changed successfully')
            self.Clear_Reset_Password_Fields()

    def Clear_Reset_Password_Fields(self):
        self.usernameresetpass.clear()
        self.currentpasswordresetpass.clear()
        self.newpasswordreset.clear()
        self.confirmnewpasswordreset.clear()
    
    def Cancel_Reset_Password(self):
        self.Daily_Move_Tab()

    def Show_User_Name(self):
        self.cur.execute('''
            SELECT user_name FROM users
        ''')
        all_names = self.cur.fetchall()
        dash = ''
        self.userproname.addItem(dash)
        for name in all_names:
            self.userproname.addItems(name)
            self.userproname.activated.connect(self.Show_User_Code)
    
    def Show_User_Code(self):
        user_name = self.userproname.currentText()
        if user_name=='':
            self.userprocode.setText('')
        else:
            self.label_192.clear()
            self.cur.execute('''
                SELECT user_code FROM users WHERE user_name=%s
            ''', (user_name,))
            user_code = self.cur.fetchone()
            self.userprocode.setText(str(user_code[0]))

    def Upload_Image(self):
        global filefi
        vbox = QVBoxLayout()
        self.setLayout(vbox)
        fname = QFileDialog.getOpenFileName(self, 'Open file', 'C:', 'Images(*.png, *.jpg)')
        imagePath = fname[0]
        pixmap = QPixmap(imagePath)
        self.label_156.setPixmap(pixmap)
        if imagePath=='':
            self.label_156.setText('')
        else:
            file = open(imagePath,'rb').read()
            filefi = base64.b64encode(file)

    def Save_User_Profile(self):
        user_namepro = self.userproname.currentText()
        user_codepro = self.userprocode.text()
        user_professional = self.usercoboxprofis.currentText()
        user_ginder = self.usercoboxginder.currentText()
        full_name = self.fname.text()
        user_address = self.useraddress.toPlainText()
        user_mobile = self.usermobile.text()
        user_smobile = self.usersmobile.text()
        self.cur.execute('''
            SELECT name FROM usersprofile
        ''')
        all_names = self.cur.fetchall()
        names_list = []
        for names in all_names:
            names_list.append(names[0])
        if full_name in names_list:
            self.label_193.setText("User Profile Already Exist")
        elif len(user_namepro)==0 or len(user_codepro)==0 or len(user_professional)==0 or len(user_ginder)==0 or len(full_name)==0 or len(user_address)==0 or len(user_mobile)==0:
            self.label_193.setText("Fill in all fields!")
        else:
            self.cur.execute('''
                INSERT INTO usersprofile(user_name, user_code, user_professional, ginder, name, address, mobile, smobile, image)
                VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s)
            ''', (user_namepro, user_codepro, user_professional, user_ginder, full_name, user_address, user_mobile, user_smobile, filefi))
            self.db.commit()
            QMessageBox.information(self, 'success', 'New User Profile Has Been Added')
            self.Clear_Profile()
            self.Login_Tab()
    
    def Clear_Profile(self):
        self.userproname.clear()
        self.userprocode.clear()
        self.label_156.clear()
        self.usercoboxprofis.setCurrentIndex(0)
        self.usercoboxginder.setCurrentIndex(0)
        self.fname.clear()
        self.useraddress.clear()
        self.usermobile.clear()
        self.usersmobile.clear()
        self.label_12.clear()
        self.Show_User_Name()
        self.Login_Tab()
    
    def Clear_Profile_2(self):
        self.label_156.clear()
        self.usercoboxprofis.setCurrentIndex(0)
        self.usercoboxginder.setCurrentIndex(0)
        self.fname.clear()
        self.useraddress.clear()
        self.usermobile.clear()
        self.usersmobile.clear()
        self.label_12.clear()

    def Edit_Profile(self):
        user_name = self.userproname.currentText()
        user_code = self.userprocode.text()
        self.cur.execute('''
                SELECT user_professional, ginder, name, address, mobile, smobile, image FROM usersprofile WHERE user_name=%s
            ''', (user_name,))
        profile_data = self.cur.fetchall()
        if len(user_name)==0 or user_code==0:
            self.label_192.setText("Invalid User Name or User Code!")
        try: 
            self.usercoboxprofis.setCurrentText(str(profile_data[0][0]))
            self.usercoboxginder.setCurrentText(str(profile_data[0][1]))
            self.fname.setText(str(profile_data[0][2]))
            self.useraddress.setText(str(profile_data[0][3]))
            self.usermobile.setText(str(profile_data[0][4]))
            self.usersmobile.setText(str(profile_data[0][5]))
            self.Retreive_Blob()
        except IndexError:
            self.label_192.setText("This user has no profile!")
            self.Clear_Profile_2()

    def Retreive_Blob(self):
        user_name = self.userproname.currentText()
        sql = '''SELECT image FROM usersprofile WHERE user_name=%s'''
        self.cur.execute(sql, ([user_name]))
        all_data = self.cur.fetchone()
        img_fi = all_data[0]
        binary_data = base64.b64decode(img_fi)
        image = Image.open(io.BytesIO(binary_data))
        qimage = ImageQt(image)
        pixmap = QtGui.QPixmap.fromImage(qimage)
        self.label_156.setPixmap(QPixmap(pixmap))
    
    def Add_User_Permission(self):
        user_name = self.userproname.currentText()
        create_tab = 0
        profile_tab = 0
        if self.userpermissioncreate.isChecked()==True:
            create_tab = 1
        if self.userpermissionprofile.isChecked()==True:
            profile_tab = 1
        
        self.cur.execute('''
            INSERT INTO user_permission(user_name, create_tab, profile_tab)
            VALUES(%s, %s, %s)
        ''', (user_name, create_tab, profile_tab))
        self.db.commit()

    def Update_Exist_Profile(self):
        user_namepro = self.userproname.currentText()
        user_codepro = self.userprocode.text()
        user_professional = self.usercoboxprofis.currentText()
        user_ginder = self.usercoboxginder.currentText()
        full_name = self.fname.text()
        user_address = self.useraddress.toPlainText()
        user_mobile = self.usermobile.text()
        user_smobile = self.usersmobile.text()
        image = self.label_156.pixmap()

        if len(user_professional)==0 or len(user_ginder)==0 or len(full_name)==0 or len(user_address)==0 or len(user_mobile)==0:
            self.label_193.setText("Fill in all fields!")
        else:
            self.cur.execute('''
                UPDATE usersprofile SET user_professional=%s, ginder=%s, name=%s, address=%s, mobile=%s, smobile%s WHERE user_name=%s
            ''', (user_professional, user_ginder, full_name, user_address, user_mobile, user_smobile, user_namepro))
            self.db.commit()
            QMessageBox.information(self, 'success', 'User Profile Has Been Updated')
            self.Clear_Profile()
            self.Login_Tab()
    
    def Delete_User_Profile(self):
        question_mark = QtWidgets.QMessageBox
        choice_del = question_mark.question(self, 'warning', 'Are You Sure Delete This User', question_mark.Yes | question_mark.No)
        if choice_del == question_mark.Yes:
            user_code = int(self.userprocode.text())
            self.cur.execute('''
                DELETE FROM usersprofile WHERE user_code=%s
            ''', (user_code,))
            self.db.commit()
            self.cur.execute('''
                DELETE FROM users WHERE user_code=%s
            ''', (user_code,))
            self.db.commit()
            self.Clear_Profile()

    def Login_Tab(self):
        self.tabWidget.setCurrentIndex(0)
        self.groupBox_4.setEnabled(False)

    def Create_User(self):
        self.tabWidget.setCurrentIndex(1)
        self.groupBox_4.setEnabled(False)

    def Open_Reset_Password(self):
        self.tabWidget.setCurrentIndex(2)
        self.groupBox_4.setEnabled(False)

    def Open_Profile(self):
        self.tabWidget.setCurrentIndex(3)
        self.groupBox_4.setEnabled(False)

    def Daily_Move_Tab(self):
        self.tabWidget.setCurrentIndex(4)
        self.groupBox_4.setEnabled(True)
        self.Auto_Incr_Calls()
    
    def Edit_Del_Calls(self):
        self.tabWidget.setCurrentIndex(5)
        
    def Delete_Executed_Call(self):
        self.tabWidget.setCurrentIndex(6)
    
    def Send_Email_Page(self):
        self.tabWidget.setCurrentIndex(7)

    def Clients_Tab(self):
        self.tabWidget.setCurrentIndex(8)
        self.tabWidget_2.setCurrentIndex(0)
        self.Client_Code_AutIncreament()
    
    def Clients_Details_Tab(self):
        self.tabWidget.setCurrentIndex(9)
        self.tabWidget_3.setCurrentIndex(0)
    
    def Technician_Tab(self):
        self.tabWidget.setCurrentIndex(10)

    def Show_User_Login(self):
        log_password = self.loginpassword.text()
        self.cur.execute('''
            SELECT user_name FROM users WHERE password=%s
        ''', (log_password,))
        login_user = self.cur.fetchall()
        if login_user==():
            self.callsusername.setText(())
        else:
            self.callsusername.setText(login_user[0][0])

    def Auto_Incr_Calls(self):
        self.cur.execute('''
            SELECT call_number From callsinformation
        ''')
        callnum = self.cur.fetchall()
        if callnum == ():
            num_zero = 1
            self.callnumber.setText(str(num_zero))
        else:
            for i in callnum:
                cal = (i[0]) + 1
                self.callnumber.setText(str(cal))

    def Active_Client_Type(self):
        self.clienttype.clear()
        self.clienttype.addItems(['', 'Exist Client', 'New Client'])
        self.clienttype.activated.connect(self.Client_Type)

    def Client_Type(self):
        currenttype = self.clienttype.currentText()
        if currenttype=='Exist Client':
            self.Active_Clname_Callspage()
        elif currenttype=='New Client':
            self.Clear_Calls_Fields_2()
            self.Autoincrement_ClBrCon_Codes()
            self.clientcoboxname.setEditable(True)
            self.branchcoboxname.setEditable(True)
            self.comboBox.setEditable(True)

    def Autoincrement_ClBrCon_Codes(self):
        self.cur.execute('''
            SELECT client_code FROM new_client
        ''')
        all_clcode = self.cur.fetchall()
        for i in all_clcode:
            new = (i[0]) + 1
            self.clientcode.setText(str(new))
        cl_code = self.clientcode.text()
        branch_code = str(cl_code) + str(101)
        self.branchcode.setText(branch_code)
        contact_code = str(cl_code) + str(201)
        self.contactcodecalpage.setText(contact_code)

    def Today_Date(self):
        self.dateEdit_7.setDate(datetime.date.today())
        self.dateEdit_10.setDate(datetime.date.today())
        self.dateEdit_18.setDate(datetime.date.today())
        self.dateEditmonth.setDate(datetime.date.today())
        self.dateEdit_6.setDate(datetime.date.today())
        self.dateEdit_5.setDate(datetime.date.today())
        self.dateEdit_4.setDate(datetime.date.today())
        self.dateEdit_3.setDate(datetime.date.today())
        self.dateEdit_11.setDate(datetime.date.today())
        self.dateEdit_19.setDate(datetime.date.today())
        self.dateEdit_20.setDate(datetime.date.today())
        self.dateEdit_21.setDate(datetime.date.today())

    def Handel_Calls(self):
        call_number = self.callnumber.text()
        call_date = self.dateEdit_7.text()
        user_name = self.callsusername.text()
        call_type = self.callcoboxtype.currentText()
        client_name = self.clientcoboxname.currentText()
        client_code = self.clientcode.text()
        main_address = self.mainaddresscallpage.toPlainText()
        branch_name = self.branchcoboxname.currentText()
        branch_code = self.branchcode.text()
        branch_address = self.branchaddress.toPlainText()
        callby = self.comboBox.currentText()
        mobile = self.mobilecallpage.text()
        smobile = self.smobilecallpage.text()
        machine_type = self.machinetypecalls.currentText()
        machine_model = self.machinemodelcallpage.text()
        machine_serial = self.machineserialcallpage.text()
        group_number = self.groupsnumcallpage.currentText()
        client_complain = self.clientcomplain.toPlainText()
        technician_name = self.textEdit_4.toPlainText()

        currenttype = self.clienttype.currentText()
        if len(currenttype)==0:
            self.label_37.setText('Please Select Client Type!')
        elif currenttype=='Exist Client':
            if len(call_type)==0 or len(client_name)==0 or len(machine_type)==0 or len(client_complain)==0 or len(technician_name)==0:
                self.label_37.setText('Please Select All Neccessary Fields (*)!')
            elif len(mobile)==0:
                self.label_37.setText('Please Select Contact Name!')
            else:
                self.cur.execute('''
                    INSERT INTO callsinformation(call_number, recieve_date, user_name, call_type, client_name, client_code, main_address, branch_name, branch_code, branch_address, call_by, mobile, smobile, machine_type, machine_model, machine_serial, group_number, client_complain, technician_name)
                    VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ''', (call_number, call_date, user_name, call_type, client_name, client_code, main_address, branch_name, branch_code, branch_address, callby, mobile, smobile, machine_type, machine_model, machine_serial, group_number, client_complain, technician_name))
                self.db.commit()
                QMessageBox.information(self, 'success', 'New Call Has Been Added')
                self.Auto_Incr_Calls()
                self.Record_Calls_History()
                self.Show_Calls_History()
                self.Show_Technician_History()
                self.Clear_Calls_Fields()
        else:
            if len(call_type)==0 or len(client_name)==0 or len(machine_type)==0 or len(client_complain)==0 or len(technician_name)==0 or len(branch_address)==0:
                self.label_37.setText('Please Select All Neccessary Fields (*)!')
            else:
                self.cur.execute('''
                    INSERT INTO callsinformation(call_number, recieve_date, user_name, call_type, client_name, client_code, main_address, branch_name, branch_code, branch_address, call_by, mobile, smobile, machine_type, machine_model, machine_serial, group_number, client_complain, technician_name)
                    VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ''', (call_number, call_date, user_name, call_type, client_name, client_code, main_address, branch_name, branch_code, branch_address, callby, mobile, smobile, machine_type, machine_model, machine_serial, group_number, client_complain, technician_name))
                self.db.commit()
                QMessageBox.information(self, 'success', 'New Call Has Been Added')
            
                self.Auto_Incr_Calls()
                self.Record_Calls_History()
                self.Show_Calls_History()
                self.Show_Technician_History()

                self.Add_Newcl_CallsP()
                self.Add_Newbr_CallsP()
                self.Add_Newcon_CallsP()
                self.Add_Newmach_CallsP()

                self.Clear_Calls_Fields()

    def Add_Newcl_CallsP(self):
        client_name = self.clientcoboxname.currentText()
        client_code = self.clientcode.text()
        main_address = self.mainaddresscallpage.toPlainText()
        join_date = self.dateEdit_7.text()
        edit_date = self.dateEdit_7.text()
        self.cur.execute('''
            INSERT INTO new_client (client_code, client_name, main_address, join_date, edit_date)
            VALUES(%s, %s, %s, %s, %s)
        ''', (client_code, client_name, main_address, join_date, edit_date))
        self.db.commit()

    def Add_Newbr_CallsP(self):
        client_name = self.clientcoboxname.currentText()
        branch_name = self.branchcoboxname.currentText()
        branch_code = self.branchcode.text()
        branch_address = self.branchaddress.toPlainText()
        join_date = self.dateEdit_7.text()
        self.cur.execute('''
            INSERT INTO client_branch (client_name, branch_name, branch_code, branch_address, join_date)
            VALUES(%s, %s, %s, %s, %s)
        ''', (client_name, branch_name, branch_code, branch_address, join_date))
        self.db.commit()

    def Add_Newcon_CallsP(self):
        client_name = self.clientcoboxname.currentText()
        contact_name = self.comboBox.currentText()
        contact_code = self.contactcodecalpage.text()
        mobile = self.mobilecallpage.text()
        smobile = self.smobilecallpage.text()
        join_date = self.dateEdit_7.text()
        self.cur.execute('''
            INSERT INTO client_contact (client_name, contact_name, contact_code, mobile, second_mobile, join_date)
            VALUES(%s, %s, %s, %s, %s, %s)
        ''', (client_name, contact_name, contact_code, mobile, smobile,  join_date))
        self.db.commit()

    def Add_Newmach_CallsP(self):
        client_name = self.clientcoboxname.currentText()
        branch_name = self.branchcoboxname.currentText()
        machine_type = self.machinetypecalls.currentText()
        machine_model = self.machinemodelcallpage.text()
        group_number = self.groupsnumcallpage.currentText()
        machine_serial = self.machineserialcallpage.text()
        join_date = self.dateEdit_7.text()
        self.cur.execute('''
            INSERT INTO client_machine (client_name, branch_name, machine_type, machine_model, machine_serial, machine_group, join_date)
            VALUES(%s, %s, %s, %s, %s, %s, %s)
        ''', (client_name, branch_name, machine_type, machine_model, group_number, machine_serial, join_date))
        self.db.commit()

    def Clear_Calls_Fields(self):
        self.comboBox.setCurrentIndex(0)
        self.techname.setCurrentIndex(0)
        self.contactcodecalpage.clear()
        self.mobilecallpage.clear()
        self.smobilecallpage.clear()
        self.clientcoboxname.setCurrentIndex(0)
        self.mainaddresscallpage.clear()
        self.branchcoboxname.setCurrentIndex(0)
        self.clientcode.clear()
        self.branchcode.clear()
        self.branchaddress.clear()
        self.machinemodelcallpage.clear()
        self.machineserialcallpage.clear()
        self.clientcomplain.clear()
        self.textEdit_4.clear()
        self.label_37.clear()
        self.Active_Client_Type()
        self.Show_Call_Type()
        self.Show_Machines_Type()

    def Active_Clname_Callspage(self):
        self.clientcoboxname.clear()
        self.clientcode.clear()
        dash = ''
        self.clientcoboxname.addItem(dash)
        self.cur.execute('''SELECT client_name FROM new_client ORDER BY client_name''')
        clientname = self.cur.fetchall()
        for items in clientname:
            self.clientcoboxname.addItems(items)
            self.clientcoboxname.activated.connect(self.Show_Related_Clcode_Claddress)
            self.clientcoboxname.activated.connect(self.Show_Related_Brname)
            self.clientcoboxname.activated.connect(self.Show_Contact_Data)

    def Activate_Combo(self):
        self.clientcoboxnamehis.clear()
        self.clientcombname.clear()
        self.contactcombclname.clear()
        self.contactcombedclname.clear()
        self.branchcoboxclname.clear()
        self.editbracoboxclname.clear()
        self.machincoboxclname.clear()
        self.machincoboxclname_2.clear()
        self.clnamemonth.clear()
        self.clnamemonth_2.clear()
        self.clnamemonth_3.clear()

        dash = ''
        self.clientcoboxnamehis.addItem(dash)
        self.clientcombname.addItem(dash)
        self.contactcombclname.addItem(dash)
        self.contactcombedclname.addItem(dash)
        self.branchcoboxclname.addItem(dash)
        self.editbracoboxclname.addItem(dash)
        self.machincoboxclname.addItem(dash)
        self.machincoboxclname_2.addItem(dash)
        self.clnamemonth.addItem(dash)
        self.clnamemonth_2.addItem(dash)
        self.clnamemonth_3.addItem(dash)


        self.cur.execute('''SELECT client_name FROM new_client ORDER BY client_name''')
        clientname = self.cur.fetchall()
        for item in clientname:
            for clname in item:
                self.clientcoboxnamehis.addItem(clname)
                self.clientcoboxnamehis.activated.connect(self.Show_Related_Brname_ExpDel)

                self.clientcombname.addItem(clname)
                self.clientcombname.activated.connect(self.Show_Neclient_Code)
                self.contactcombclname.addItem(clname)
                self.contactcombclname.activated.connect(self.Show_Conclient_Code)
                self.contactcombclname.activated.connect(self.New_Contact_Code)

                self.contactcombedclname.addItem(clname)
                self.contactcombedclname.activated.connect(self.Activate_Contact_Name)
                self.contactcombedclname.activated.connect(self.Show_All_Contact)

                self.branchcoboxclname.addItem(clname)
                self.branchcoboxclname.activated.connect(self.Show_Clclient_code)
                self.branchcoboxclname.activated.connect(self.New_Branch_Code)

                self.editbracoboxclname.addItem(clname)
                self.editbracoboxclname.activated.connect(self.Show_Branch_Name)
                self.editbracoboxclname.activated.connect(self.Show_Client_Branshes)

                self.machincoboxclname.addItem(clname)
                self.machincoboxclname.activated.connect(self.Show_Related_Brname_AddNewMach)
                self.machincoboxclname_2.addItem(clname)
                self.machincoboxclname_2.activated.connect(self.Show_Related_Brname_EditMach)
                self.machincoboxclname_2.activated.connect(self.Show_Client_CoMaEd)

                self.clnamemonth.addItem(clname)
                self.clnamemonth.activated.connect(self.Show_Brname_Month)
                self.clnamemonth_2.addItem(clname)
                self.clnamemonth_2.activated.connect(self.Show_Branch_TecMon)
                self.clnamemonth_3.addItem(clname)
                self.clnamemonth_3.activated.connect(self.Show_Branch_DelMon)

    def Show_Related_Brname(self):
        client_name = self.clientcoboxname.currentText()
        self.branchcoboxname.clear()
        dash=''
        self.branchcoboxname.addItem(dash)
        self.cur.execute('''
            SELECT branch_name FROM client_branch WHERE client_name=%s ORDER BY branch_name
        ''', (client_name,))
        branchesname = self.cur.fetchall()
        for item in branchesname:
            for brname in item:
                self.branchcoboxname.addItem(brname)
                self.branchcoboxname.activated.connect(self.Show_Related_Data)
                self.branchcoboxname.activated.connect(self.Machine_Type_Combobox)

    def Show_Related_Brname_ExpDel(self):
        client_name_his = self.clientcoboxnamehis.currentText()
        self.branchcoboxnamehis.clear()
        dash = ''
        self.branchcoboxnamehis.addItem(dash)
        self.cur.execute('''
            SELECT branch_name FROM client_branch WHERE client_name=%s ORDER BY branch_name
        ''', (client_name_his,))
        branches_name = self.cur.fetchall()
        for item in branches_name:
            for brname in item:
                self.branchcoboxnamehis.addItem(brname)

    def Show_Related_Brname_AddNewMach(self):
        client_name_mach = self.machincoboxclname.currentText()
        self.machinecoboxbrname.clear()
        dash = ''
        self.machinecoboxbrname.addItem(dash)
        self.cur.execute('''
            SELECT branch_name FROM client_branch WHERE client_name=%s ORDER BY branch_name
        ''', (client_name_mach,))
        branchesname = self.cur.fetchall()
        for item in branchesname:
            for brname in item:
                self.machinecoboxbrname.addItem(brname)

    def Show_Related_Brname_EditMach(self):
        client_name_mach_2 = self.machincoboxclname_2.currentText()
        self.machinecombrname_2.clear()
        dash = ''
        self.machinecombrname_2.addItem(dash)
        self.cur.execute('''
            SELECT branch_name FROM client_branch WHERE client_name=%s ORDER BY branch_name
        ''', (client_name_mach_2,))
        all_branches = self.cur.fetchall()
        for branches in all_branches:
            self.machinecombrname_2.addItem(branches[0])
            self.machinecombrname_2.activated.connect(self.Show_Branch_COMaEd)

    def Show_Related_Clcode_Claddress(self):
        client_name = self.clientcoboxname.currentText()
        try:
            self.label_53.clear()
            self.cur.execute('''
                SELECT client_code, main_address FROM new_client WHERE client_name=%s
            ''', (client_name,))
            client_data = self.cur.fetchall()
            self.clientcode.setText(str(client_data[0][0]))
            self.mainaddresscallpage.setText(str(client_data[0][1]))
        except IndexError:
            self.clientcode.clear()
            self.mainaddresscallpage.clear()
            self.label_53.setText(" ")

    def Show_Related_Data(self):
        self.machinetypecalls.setCurrentIndex(0)
        self.machinemodelcallpage.clear()
        self.machineserialcallpage.clear()
        self.groupsnumcallpage.setCurrentIndex(0)
        try:
            branch_name = self.branchcoboxname.currentText()
            self.cur.execute('''
                SELECT branch_code, branch_address FROM client_branch WHERE branch_name=%s
            ''', (branch_name,))
            bra_coadd = self.cur.fetchall()
            self.branchcode.setText(str(bra_coadd[0][0]))
            self.branchaddress.setText(str(bra_coadd[0][1]))
        except IndexError:
            self.branchcode.clear()
            self.branchaddress.clear()
            self.label_53.setText(" ")
    
    def Show_Contact_Data(self):
        client_name = self.clientcoboxname.currentText()
        try:
            self.cur.execute('''
                    SELECT contact_name FROM client_contact WHERE client_name=%s ORDER BY contact_name
            ''', (client_name,))
            all_data = self.cur.fetchall()
            self.comboBox.clear()
            dash = ''
            self.comboBox.addItem(dash)
            for any_data in all_data:
                for single_data in any_data:
                    self.comboBox.addItem(single_data)
                    self.comboBox.activated.connect(self.Select_Code_Mobile)
        except IndexError:
            self.mobilecallpage.clear()
            self.smobilecallpage.clear()
            self.label_53.setText(" ")

    def Machine_Type_Combobox(self):
        self.machinetypecalls.activated.connect(self.Show_Machine_Model_Serial)

    def Show_Machine_Model_Serial(self):
        branch_name = self.branchcoboxname.currentText()
        try:
            machine_type = self.machinetypecalls.currentText()
            if machine_type == "Coffe Machine":
                self.cur.execute('''
                    SELECT machine_model, machine_serial, number_of_groups FROM client_information WHERE branch_name=%s
                ''', (branch_name,))
                machine_data = self.cur.fetchall()
                self.machinemodelcallpage.setText(str(machine_data[0][0]))
                self.machineserialcallpage.setText(str(machine_data[0][1]))
                group_quantity = str(machine_data[0][2])
                if group_quantity=="1GR":
                    self.groupsnumcallpage.setCurrentIndex(1)
                elif group_quantity=="2GR":
                    self.groupsnumcallpage.setCurrentIndex(2)
                elif group_quantity=="3GR":
                    self.groupsnumcallpage.setCurrentIndex(3)
            elif machine_type == "Grinder":
                self.groupsnumcallpage.setCurrentIndex(0)
                self.cur.execute('''
                    SELECT grinder_model, grinder_serial FROM client_information WHERE branch_name=%s
                ''', (branch_name,))
                machine_data = self.cur.fetchall()
                self.machinemodelcallpage.setText(str(machine_data[0][0]))
                self.machineserialcallpage.setText(str(machine_data[0][1]))
            else:
                self.label_37.clear()
        except IndexError:
            self.label_37.setText("Please Select Machine Type!")

    def Select_Code_Mobile(self):
        contact_name = self.comboBox.currentText()
        if len(contact_name)==0:
            self.contactcodecalpage.setText('')
            self.mobilecallpage.clear()
            self.smobilecallpage.clear()
        else:
            self.cur.execute('''
                SELECT contact_code FROM client_contact WHERE contact_name=%s
            ''', (contact_name,))
            all_data1 = self.cur.fetchone()
            self.contactcodecalpage.setText(str(all_data1[0]))

            self.cur.execute('''
                SELECT mobile, second_mobile FROM client_contact WHERE contact_name=%s
            ''', (contact_name,))
            all_data2 = self.cur.fetchall()
            self.mobilecallpage.setText(str(all_data2[0][0]))
            self.smobilecallpage.setText(str(all_data2[0][1]))

    def Show_Call_Type(self):
        self.callcoboxtype.setCurrentIndex(0)

    def Clear_Calls_Fields_2(self):
        self.callcoboxtype.setCurrentIndex(0)
        self.machinetypecalls.setCurrentIndex(0)
        self.groupsnumcallpage.setCurrentIndex(0)
        self.comboBox.setCurrentIndex(0)
        self.techname.setCurrentIndex(0)
        self.contactcodecalpage.clear()
        self.mobilecallpage.clear()
        self.clientcoboxname.clear()
        self.branchcoboxname.clear()
        self.clientcode.clear()
        self.branchcode.clear()
        self.branchaddress.clear()
        self.clientcomplain.clear()
        self.textEdit_4.clear()

    def Tech_Combo_Activ(self):
        self.techname.setCurrentIndex(0)
        self.techname.activated.connect(self.Show_Techs_Names)

    def Show_Techs_Names(self):
        tech_name = self.techname.currentText()
        self.textEdit_4.setText(f'{self.textEdit_4.toPlainText()}\n{tech_name}'.strip())

    def Show_Machines_Type(self):
        self.machinetypecalls.setCurrentIndex(0)
        self.groupsnumcallpage.setCurrentIndex(0)

    def Show_Calls_History(self):
        self.tableWidget_10.setRowCount(0)
        self.tableWidget_10.insertRow(0)
        self.cur.execute('''
            SELECT user_name, client_name, branch_name, machine_type, client_complain, call_by, mobile, technician_name, call_number, recieve_date FROM callsinformation
        ''')
        allcalls = self.cur.fetchall()
        for row, form in enumerate(allcalls):
            for col, item in enumerate(form):
                self.tableWidget_10.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_10.rowCount()
            self.tableWidget_10.insertRow(row_position)

    def Record_Calls_History(self):
        user_name = self.callsusername.text()
        client_name = self.clientcoboxname.currentText()
        branch_name = self.branchcoboxname.currentText()
        machine_type = self.machinetypecalls.currentText()
        client_complain = self.clientcomplain.toPlainText()
        callby = self.comboBox.currentText()
        technician_name = self.textEdit_4.toPlainText()
        call_number = self.callnumber.text()
        call_type = self.callcoboxtype.currentText()
        call_date = self.dateEdit_7.text()
        
        self.cur.execute('''
            INSERT INTO callshistory(user_name, client_name, branch_name, machine_type, client_complain, call_by, technician_name, call_number, recieve_date)
            VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s)
        ''', (user_name, client_name, branch_name, machine_type, client_complain, callby, technician_name, call_number, call_date))
        self.db.commit()

        self.cur.execute('''
            INSERT INTO technician_calls(call_number, user_name,technician_name, client_name, branch_name, call_type, machine_type, recieve_date)
            VALUES(%s, %s, %s, %s, %s, %s, %s, %s)
        ''', (call_number, user_name, technician_name, client_name, branch_name, call_type, machine_type, call_date))
        self.db.commit()
    
    def Show_Brname_Calledit(self):
        self.branamecomcalledit.clear()
        client_name = self.callscoboxclnamedit.currentText()
        self.cur.execute('''SELECT branch_name FROM client_branch WHERE client_name=%s''', (client_name,))
        branchesname = self.cur.fetchall()
        for item in branchesname:
            for brname in item:
                self.branamecomcalledit.addItem(brname)

    def Search_Calls_History(self):
        branch_name = self.branchcoboxnamehis.currentText()
        if  len(branch_name)==0:
            self.Search_Clientn_History()
        elif len(branch_name)!=0:
            self.Search_Clientn_Branchn_History()
    
    def Search_Clientn_History(self):
        client_name = self.clientcoboxnamehis.currentText()
        self.tableWidget_10.setRowCount(0)
        self.tableWidget_10.insertRow(0)
        self.cur.execute('''
            SELECT user_name, client_name, branch_name, machine_type, client_complain, call_by, mobile, technician_name, call_number, recieve_date FROM callsinformation WHERE client_name=%s
        ''', (client_name,))
        allcalls = self.cur.fetchall()
        for row, form in enumerate(allcalls):
            for col, item in enumerate(form):
                self.tableWidget_10.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_10.rowCount()
            self.tableWidget_10.insertRow(row_position)

    def Search_Clientn_Branchn_History(self):
        client_name = self.clientcoboxnamehis.currentText()
        branch_name = self.branchcoboxnamehis.currentText()
        self.tableWidget_10.setRowCount(0)
        self.tableWidget_10.insertRow(0)
        self.cur.execute('''
            SELECT user_name, client_name, branch_name, machine_type, client_complain, call_by, mobile, technician_name, call_number, recieve_date FROM callsinformation WHERE client_name=%s AND branch_name=%s
        ''', (client_name, branch_name))
        allcalls = self.cur.fetchall()
        for row, form in enumerate(allcalls):
            for col, item in enumerate(form):
                self.tableWidget_10.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_10.rowCount()
            self.tableWidget_10.insertRow(row_position)

    def Clear_Search_History(self):
        self.Show_Calls_History()

    def Export_Dailly_Calls(self):
        branch_name = self.branchcoboxnamehis.currentText()
        if len(branch_name)==0:
            self.Export_By_Client()
        elif len(branch_name)!=0:
            self.Export_By_Clientn_Branchn()
        else:
            self.cur.execute('''
                SELECT user_name, client_name, branch_name, machine_type, client_complain, call_by, mobile, technician_name, call_number, recieve_date FROM callsinformation
            ''')
            all_calls = self.cur.fetchall()
            filename = QFileDialog.getSaveFileName(self, 'Save File', 'C:', ".xls(*.xls)")
            my_file = filename[0]
            excel_file = xlwt.Workbook()
            self.sheet1 = excel_file.add_sheet('Dailly Calls', cell_overwrite_ok=True)
            
            self.sheet1.write(0, 0, 'User Name')
            self.sheet1.write(0, 1, 'Client Name')
            self.sheet1.write(0, 2, 'Branch Name')
            self.sheet1.write(0, 3, 'Machine Type')
            self.sheet1.write(0, 4, 'Client Complain')
            self.sheet1.write(0, 5, 'Call By')
            self.sheet1.write(0, 6, 'Mobile')
            self.sheet1.write(0, 7, 'Technician Name')
            self.sheet1.write(0, 8, 'Call Number')
            self.sheet1.write(0, 9, 'Date')

            row_number = 0
            for row in all_calls:
                col = 0
                for item in row:
                    self.sheet1.write(row_number, col, str(item))
                    col += 1
                row_number += 1
            if my_file=='':
                self.Edit_Del_Calls()
            else:
                excel_file.save(my_file)

    def Export_By_Client(self):
        client_name = self.clientcoboxnamehis.currentText()
        self.cur.execute('''
            SELECT user_name, client_name, branch_name, machine_type, client_complain, call_by, mobile, technician_name, call_number, recieve_date FROM callsinformation WHERE client_name=%s
        ''', (client_name,))
        all_calls = self.cur.fetchall() 
        filename = QFileDialog.getSaveFileName(self, 'Save File', 'C:', ".xls(*.xls)")
        my_file = filename[0]
        excel_file = xlwt.Workbook()
        self.sheet1 = excel_file.add_sheet('Dailly Calls', cell_overwrite_ok=True)
        self.sheet1.write(0, 0, 'User Name')
        self.sheet1.write(0, 1, 'Client Name')
        self.sheet1.write(0, 2, 'Branch Name')
        self.sheet1.write(0, 3, 'Machine Type')
        self.sheet1.write(0, 4, 'Client Complain')
        self.sheet1.write(0, 5, 'Call By')
        self.sheet1.write(0, 6, 'Mobile')
        self.sheet1.write(0, 7, 'Technician Name')
        self.sheet1.write(0, 8, 'Call Number')
        self.sheet1.write(0, 9, 'Date')
        row_number = 1
        for row in all_calls:
            col = 0
            for item in row:
                self.sheet1.write(row_number, col, str(item))
                col += 1
            row_number += 1
        if my_file=='':
            self.Edit_Del_Calls()
        else:
            excel_file.save(my_file)

    def Export_By_Clientn_Branchn(self):
        client_name = self.clientcoboxnamehis.currentText()
        branch_name = self.branchcoboxnamehis.currentText()
        self.cur.execute('''
            SELECT user_name, client_name, branch_name, machine_type, client_complain, call_by, mobile, technician_name, call_number, recieve_date FROM callsinformation WHERE client_name=%s AND branch_name=%s
        ''', (client_name, branch_name))
        all_calls = self.cur.fetchall()
        filename = QFileDialog.getSaveFileName(self, 'Save File', 'C:', ".xls(*.xls)")
        my_file = filename[0]
        excel_file = xlwt.Workbook()
        self.sheet1 = excel_file.add_sheet('Dailly Calls', cell_overwrite_ok=True)
        self.sheet1.write(0, 0, 'User Name')
        self.sheet1.write(0, 1, 'Client Name')
        self.sheet1.write(0, 2, 'Branch Name')
        self.sheet1.write(0, 3, 'Machine Type')
        self.sheet1.write(0, 4, 'Client Complain')
        self.sheet1.write(0, 5, 'Call By')
        self.sheet1.write(0, 6, 'Mobile')
        self.sheet1.write(0, 7, 'Technician Name')
        self.sheet1.write(0, 8, 'Call Number')
        self.sheet1.write(0, 9, 'Date')
        row_number = 1
        for row in all_calls:
            col = 0
            for item in row:
                self.sheet1.write(row_number, col, str(item))
                col += 1
            row_number += 1
        if my_file=='':
            self.Edit_Del_Calls()
        else:
            excel_file.save(my_file)

    def Activate_Combo_Dele(self):
        self.clientcoboxnamedel.clear()
        dash = ''
        self.clientcoboxnamedel.addItem(dash)
        self.cur.execute('''SELECT client_name FROM new_client''')
        clientname = self.cur.fetchall()
        for item in clientname:
            for clname in item:
                self.clientcoboxnamedel.addItem(clname)
                self.clientcoboxnamedel.activated.connect(self.Show_Related_Brname_Dele)

    def Show_Related_Brname_Dele(self):
        self.branchcoboxnamedel.clear()
        dash = ''
        self.branchcoboxnamedel.addItem(dash)
        client_name = self.clientcoboxnamedel.currentText()
        self.cur.execute('''SELECT branch_name FROM callsinformation WHERE client_name=%s''', (client_name,))
        branchesname = self.cur.fetchall()
        for item in branchesname:
            for brname in item:
                self.branchcoboxnamedel.addItem(brname)
                self.branchcoboxnamedel.activated.connect(self.Show_Related_Call_Number)       

    def Show_Related_Call_Number(self):
        self.callnumdel.clear()
        branch_name = self.branchcoboxnamedel.currentText()
        self.cur.execute('''SELECT call_number FROM callsinformation WHERE branch_name=%s''', (branch_name,))
        callnumber = self.cur.fetchone()
        self.callnumdel.addItem(str(callnumber[0]))

    def Fill_Fields(self):
        call_numb = self.callnumdel.currentText()
        self.cur.execute('''SELECT call_type, call_by, client_complain, technician_name FROM callsinformation WHERE call_number=%s''', (call_numb,))
        all_data = self.cur.fetchall()
        self.lineEdit_19.setText(all_data[0][0])
        self.lineEdit_18.setText(all_data[0][1])
        self.textEdit_5.setPlainText(all_data[0][2])
        self.textEdit_3.setPlainText(all_data[0][3])

    def Empty_Labels(self):
        call_type = self.lineEdit_19.text()
        callby = self.lineEdit_18.text()
        client_comp = self.textEdit_5.toPlainText()
        tech_assigned = self.textEdit_3.toPlainText()
        if len(call_type)==0 or len(callby)==0 or len(client_comp)==0 or len(tech_assigned)==0:
            self.label_6.setText('Empty Fields Not Accepted, Please Try Again')
        else:
            self.Delete_Calls()

    def Delete_Calls(self):
        question_mark = QtWidgets.QMessageBox
        choice_del = question_mark.question(self, 'warning', 'Are You Sure Delete This Call', question_mark.Yes | question_mark.No)
        if choice_del == question_mark.Yes:
            call_numb = self.callnumdel.currentText()
            self.cur.execute('''
                DELETE FROM callsinformation WHERE call_number=%s
            ''', (call_numb,))
            self.db.commit()
            self.cur.execute('''
                DELETE FROM callshistory WHERE call_number=%s
            ''', (call_numb,))
            self.db.commit()
            self.Edit_Del_Calls()
            self.Show_Calls_History()

    def Client_Code_AutIncreament(self):
        self.cur.execute('''
            SELECT client_code FROM new_client
        ''')
        client_codes = self.cur.fetchall()
        if client_codes == ():
            start_counter = '101'
            self.lineEdit_86.setText(start_counter)
        else:
            for i in client_codes:
                cal = (i[0]) + 1
                self.lineEdit_86.setText(str(cal))

    def Add_New_Client(self):
        client_code = self.lineEdit_86.text()
        client_name = self.lineEdit_88.text()
        client_address = self.textEdit_13.toPlainText()
        join_date = self.dateEdit_6.text()

        if len(client_name)==0 or len(client_address)==0:
            self.label_56.setText('All Fields (*) Are Neccessary!')
        else:
            self.cur.execute('''
                INSERT INTO new_client(client_code, client_name, main_address, join_date)
                VALUES(%s, %s, %s, %s)
            ''', (client_code, client_name, client_address, join_date))
            self.db.commit()

            self.Show_Client()
            self.Activate_Combo()
            QMessageBox.information(self, 'success', 'New Client Has Been Added')
            self.Clear_New_Client()
    
    def Clear_New_Client(self):
        self.lineEdit_88.clear()
        self.textEdit_13.clear()
        self.Client_Code_AutIncreament()

    def Show_Conclient_Code(self):
        client_name = self.contactcombclname.currentText()
        if len(client_name)==0:
            self.lineEdit_contactcl.setText('')
        else:
            self.cur.execute('''
                SELECT client_code FROM new_client WHERE client_name=%s
            ''', (client_name,))
            client_code = self.cur.fetchone()
            self.lineEdit_contactcl.setText(str(client_code[0]))

    def New_Contact_Code(self):
        client_name = self.contactcombclname.currentText()
        cient_code = self.lineEdit_contactcl.text()
        if len(client_name)==0:
            self.lineEdit_85.setText('')
        else:
            self.cur.execute('''
                SELECT contact_code FROM client_contact WHERE client_name=%s
            ''', (client_name,))
            contact_codes = self.cur.fetchall()
            if contact_codes == ():
                start_counter = cient_code + ('201')
                self.lineEdit_85.setText(str(start_counter))
            else:
                for i in contact_codes:
                    cal = (i[0]) + 1
                    self.lineEdit_85.setText(str(cal))

    def Add_Client_Contact(self):
        client_name = self.contactcombclname.currentText()
        contact_name = self.lineEdit_76.text()
        contact_code = self.lineEdit_85.text()
        mobile = self.lineEdit_77.text()
        second_mobile = self.lineEdit_78.text()
        join_date = self.dateEdit_5.text()
        if len(client_name)==0 or len(contact_name)==0 or len(mobile)==0:
            self.label_20.setText('All Fields (*) Are Neccessary!')
        else:
            self.cur.execute('''
                INSERT INTO client_contact(client_name, contact_name, contact_code, mobile, second_mobile, join_date)
                VALUES(%s, %s, %s, %s, %s, %s)
            ''', (client_name, contact_name, contact_code, mobile, second_mobile, join_date))
            self.db.commit()

            QMessageBox.information(self, 'success', 'Client Contact Has Been Added')
            self.label_20.clear()
            self.label_26.clear()
            self.Show_Contact()
            self.Clear_Client_Contact()

    def Clear_Client_Contact(self):
        self.contactcombclname.setCurrentIndex(0)
        self.lineEdit_contactcl.clear()
        self.lineEdit_76.clear()
        self.lineEdit_85.clear()
        self.lineEdit_77.clear()
        self.lineEdit_78.clear()

    def Show_Clclient_code(self):
        client_name = self.branchcoboxclname.currentText()
        if len(client_name)==0:
            self.lineEdit_89.setText('')
        else:
            self.cur.execute('''
                SELECT client_code FROM new_client WHERE client_name=%s
            ''', (client_name,))
            client_code = self.cur.fetchone()
            self.lineEdit_89.setText(str(client_code[0]))

    def New_Branch_Code(self):
        client_name = self.branchcoboxclname.currentText()
        cient_code = self.lineEdit_89.text()
        if len(client_name)==0:
            self.lineEdit_90.setText('')
        else:
            self.cur.execute('''
                SELECT branch_code FROM client_branch WHERE client_name=%s
            ''', (client_name,))
            branch_codes = self.cur.fetchall()
            if branch_codes == ():
                start_counter = cient_code + ('001')
                self.lineEdit_90.setText(str(start_counter))
            else:
                for i in branch_codes:
                    cal = (i[0]) + 1
                    self.lineEdit_90.setText(str(cal))

    def Add_New_Branch(self):
        client_name = self.branchcoboxclname.currentText()
        branch_name = self.lineEdit_81.text()
        branch_code = self.lineEdit_90.text()
        branch_address = self.textEdit_12.toPlainText()
        join_date = self.dateEdit_3.text()
        
        if len(branch_name)==0 or len(branch_address)==0:
            self.label_27.setText('All Fields (*) Are Neccessary!')
        else:
            self.cur.execute('''
                INSERT INTO client_branch(client_name, branch_name, branch_code, branch_address, join_date)
                VALUES(%s, %s, %s, %s, %s)
            ''', (client_name, branch_name, branch_code, branch_address, join_date))
            self.db.commit()
            self.Show_Branch()
            QMessageBox.information(self, 'success', 'Client Branch Has Been Added')
            self.Clear_New_Branch()
    
    def Clear_New_Branch(self):
        self.branchcoboxclname.setCurrentIndex(0)
        self.lineEdit_89.clear()
        self.lineEdit_81.clear()
        self.lineEdit_90.clear()
        self.textEdit_12.clear()
        self.label_27.clear()

    def Add_Client_Machine(self):
        client_name = self.machincoboxclname.currentText()
        branch_name = self.machinecoboxbrname.currentText()
        machine_type = self.machinetype.currentText()
        machine_model = self.lineEdit_84.text()
        machine_serial = self.lineEdit_83.text()
        machine_groups = self.groupsnum.currentText()
        join_date = self.dateEdit_3.text()

        if len(branch_name)==0 or len(machine_type)==0 or len(machine_model)==0:
            self.label_36.setText('All Fields (*) Are Neccessary!')
        elif machine_type=='Graneta Machine' and machine_type=='Coffe Machine':
            self.label_36.setText('Please Select Group Quantity!')
        else:
            self.cur.execute('''
                INSERT INTO client_machine(client_name, branch_name, machine_type, machine_model, machine_serial, machine_group, join_date)
                VALUES(%s, %s, %s, %s, %s, %s, %s)
            ''', (client_name, branch_name, machine_type, machine_model, machine_serial, machine_groups, join_date))
            self.db.commit()
            self.Show_Machine()
            QMessageBox.information(self, 'success', 'Client Machine Has Been Added')
            self.Clear_Client_Machine()

    def Clear_Client_Machine(self):
        self.machincoboxclname.clear()
        self.machinecoboxbrname.clear()
        dash = ''
        self.machinecoboxbrname.addItem(dash)
        self.lineEdit_84.clear()
        self.lineEdit_83.clear()
        self.label_36.clear()
        self.machinetype.setCurrentIndex(0)
        self.groupsnum.setCurrentIndex(0)
        self.Activate_Combo()

    def Show_Neclient_Code(self):
        client_name = self.clientcombname.currentText()
        self.cur.execute('''
            SELECT client_code FROM new_client WHERE client_name=%s
        ''', (client_name,))
        client_code = self.cur.fetchone()
        self.lineEdit_clcode.setText(str(client_code[0]))

    def Table4_Column_Width(self):
        self.tableWidget_4.setColumnWidth(0, 100)
        self.tableWidget_4.setColumnWidth(1, 150)
        self.tableWidget_4.setColumnWidth(2, 250)
        self.tableWidget_4.setColumnWidth(3, 100)
        self.tableWidget_4.setColumnWidth(4, 100)

    def Show_Client(self):
        self.Table4_Column_Width()
        self.tableWidget_4.setRowCount(0)
        self.tableWidget_4.insertRow(0)
        self.cur.execute('''
            SELECT client_code, client_name, main_address, join_date, edit_date FROM new_client
        ''')
        client_information = self.cur.fetchall()
        for row, form in enumerate(client_information):
            for col, item in enumerate(form):
                self.tableWidget_4.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_4.rowCount()
            self.tableWidget_4.insertRow(row_position)

    def Search_Client_Name(self):
        client_name = self.clientcombname.currentText()
        self.cur.execute('''
            SELECT main_address, join_date FROM new_client WHERE client_name=%s
        ''', (client_name,))
        all_info = self.cur.fetchall()
        for info in all_info:
            self.textEditnewcladdress.setText(info[0])
            self.dateEdit_23.setDate(datetime.datetime.strptime(info[1], '%d-%m-%Y'))

        self.tableWidget_4.setRowCount(0)
        self.tableWidget_4.insertRow(0)
        self.cur.execute('''
            SELECT client_code, client_name, main_address, join_date, edit_date FROM new_client WHERE client_name=%s
        ''', (client_name,))
        client_information = self.cur.fetchall()
        for row, form in enumerate(client_information):
            for col, item in enumerate(form):
                self.tableWidget_4.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_4.rowCount()
            self.tableWidget_4.insertRow(row_position)
    
    def Clear_Search_Client(self):
        self.clientcombname.clear()
        self.lineEdit_clcode.clear()
        self.textEditnewcladdress.clear()
        d2 = '01-01-2000'
        self.dateEdit_23.setDate(datetime.datetime.strptime(d2, '%d-%m-%Y'))
        self.Activate_Combo()
        self.Show_Client()

    def Edit_Client_Information(self):
        add_newadd = self.textEditnewcladdress.toPlainText()
        if len(add_newadd)==0:
            self.label_6.setText('Empty Field Not Acceptable!')
        else:
            client_code = int(self.lineEdit_clcode.text())
            new_address = self.textEditnewcladdress.toPlainText()
            edit_date= self.dateEdit_11.text()
            self.cur.execute('''
                UPDATE new_client SET main_address=%s, edit_date=%s WHERE client_code=%s
            ''', (new_address, edit_date, client_code))
            self.db.commit()
            QMessageBox.information(self, 'success', 'New Address Has Been Edited')
            self.Search_Client_Name()
            self.Activate_Combo()
            self.textEditnewcladdress.clear()
            self.label_6.clear()

    def Delete_Client(self):
        client_name = self.clientcombname.currentText()
        question_mark = QtWidgets.QMessageBox
        choice_del = question_mark.question(self, 'warning', 'Are You Sure Delete This Client', question_mark.Yes | question_mark.No)
        if choice_del == question_mark.Yes:
            client_code = int(self.lineEdit_clcode.text())
            self.cur.execute('''
                DELETE FROM new_client WHERE client_code=%s
            ''', (client_code,))
            self.cur.execute('''
                DELETE FROM client_contact WHERE client_name=%s
            ''', (client_name,))
            self.cur.execute('''
                DELETE FROM client_branch WHERE client_name=%s
            ''', (client_name,))
            self.cur.execute('''
                DELETE FROM client_machine WHERE client_name=%s
            ''', (client_name,))
            self.db.commit()
        self.Show_Client()
        self.Show_Contact()
        self.Clear_Search_Client()

    def Activate_Contact_Name(self):
        self.concombedcona.clear()
        dash = ''
        self.concombedcona.addItem(dash)
        
        client_name = self.contactcombedclname.currentText()
        self.cur.execute('''
            SELECT contact_name FROM client_contact WHERE client_name=%s
        ''', (client_name,))
        all_contact = self.cur.fetchall()
        for item in all_contact:
            for conname in item:
                self.concombedcona.addItem(conname)
                self.concombedcona.activated.connect(self.Show_Contact_Code)

    def Table7_Column_Width(self):
        self.tableWidget_7.setColumnWidth(0, 150)
        self.tableWidget_7.setColumnWidth(1, 200)
        self.tableWidget_7.setColumnWidth(2, 130)
        self.tableWidget_7.setColumnWidth(3, 150)
        self.tableWidget_7.setColumnWidth(4, 150)
        self.tableWidget_7.setColumnWidth(5, 100)
        self.tableWidget_7.setColumnWidth(6, 100)

    def Show_Contact(self):
        self.Table7_Column_Width()
        self.tableWidget_7.setRowCount(0)
        self.tableWidget_7.insertRow(0)
        self.cur.execute('''
                SELECT client_name, contact_name, contact_code,  mobile, second_mobile, join_date, edit_date FROM client_contact 
            ''')
        contact_client = self.cur.fetchall()
        for row, form in enumerate(contact_client):
            for col, item in enumerate(form):
                self.tableWidget_7.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_7.rowCount()
            self.tableWidget_7.insertRow(row_position)
    
    def Show_Contact_Code(self):
        contact_name = self.concombedcona.currentText()
        if len(contact_name)==0:
            self.lineEdit_concode.setText('')
            self.lineEdit_79.clear()
            self.lineEdit_82.clear()
        else:
            self.cur.execute('''
                SELECT contact_code FROM client_contact WHERE contact_name=%s
            ''', (contact_name,))
            contact_code = self.cur.fetchone()
            self.lineEdit_concode.setText(str(contact_code[0]))

    def Search_Contact(self):
        contact_code = self.lineEdit_concode.text()
        self.cur.execute('''
            SELECT mobile, second_mobile FROM client_contact WHERE contact_code=%s
        ''', (contact_code,))
        all_mobiles = self.cur.fetchall()
        for mobiles in all_mobiles:
            self.lineEdit_79.setText(mobiles[0])
            self.lineEdit_82.setText(mobiles[1])

        self.tableWidget_7.setRowCount(0)
        self.tableWidget_7.insertRow(0)
        self.cur.execute('''
            SELECT client_name, contact_name, contact_code, mobile, second_mobile, join_date, edit_date FROM client_contact WHERE contact_code=%s
        ''', (contact_code,))
        client_information = self.cur.fetchall()
        for row, form in enumerate(client_information):
            for col, item in enumerate(form):
                self.tableWidget_7.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_7.rowCount()
            self.tableWidget_7.insertRow(row_position)

    def Show_All_Contact(self):
        client_name = self.contactcombedclname.currentText()
        self.tableWidget_7.setRowCount(0)
        self.tableWidget_7.insertRow(0)
        self.cur.execute('''
            SELECT client_name, contact_name, contact_code, mobile, second_mobile, join_date, edit_date FROM client_contact WHERE client_name=%s
        ''', (client_name,))
        client_information = self.cur.fetchall()
        for row, form in enumerate(client_information):
            for col, item in enumerate(form):
                self.tableWidget_7.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_7.rowCount()
            self.tableWidget_7.insertRow(row_position)

    def Clear_Search_Contact(self):
        self.contactcombedclname.clear()
        self.concombedcona.clear()
        self.lineEdit_concode.clear()
        self.lineEdit_79.clear()
        self.lineEdit_82.clear()
        self.label_17.clear()
        self.Show_Contact()
        self.Activate_Combo()
        self.Activate_Contact_Name()

    def Eidt_Contact_Name(self):
        mobile = self.lineEdit_79.text()
        second_mobile = self.lineEdit_82.text()
        contact_code = self.lineEdit_concode.text()
        edit_date = self.dateEdit_19.text()
        if len(mobile)!=0 and len(second_mobile)==0:
            self.cur.execute('''
                UPDATE client_contact SET mobile=%s, edit_date=%s WHERE contact_code=%s
            ''', (mobile, edit_date, contact_code))
            self.db.commit()
            QMessageBox.information(self, 'success', 'Client Contact Has Been Edited')
            self.Search_Contact()
            self.lineEdit_79.clear()
            self.lineEdit_82.clear()
            self.label_17.clear()
        elif len(mobile)==0 and len(second_mobile)!=0:
            self.cur.execute('''
                UPDATE client_contact SET second_mobile=%s, edit_date=%s WHERE contact_code=%s
            ''', (second_mobile, edit_date, contact_code))
            self.db.commit()
            QMessageBox.information(self, 'success', 'Client Contact Has Been Edited')
            self.Search_Contact()
            self.lineEdit_79.clear()
            self.lineEdit_82.clear()
            self.label_17.clear()
        elif len(mobile)==0 or len(second_mobile)==0:
            self.label_17.setText('Please Fillin Necessary Field/Fields')
        else:
            self.cur.execute('''
                UPDATE client_contact SET mobile=%s, second_mobile=%s, edit_date=%s WHERE contact_code=%s
            ''', (mobile, second_mobile, edit_date, contact_code))
            self.db.commit()
            QMessageBox.information(self, 'success', 'Client Contact Has Been Edited')
            self.Search_Contact()
            self.lineEdit_79.clear()
            self.lineEdit_82.clear()
            self.label_17.clear()

    def Delete_Contact(self):
        self.concombedcona.clear()
        client_name = self.contactcombedclname.currentText()
        self.cur.execute('''
            SELECT contact_name FROM client_contact WHERE client_name=%s
        ''', (client_name,))
        all_contact = self.cur.fetchall()
        for item in all_contact:
            for conname in item:
                self.concombedcona.addItem(conname)
        count = self.concombedcona.count()
        if count==1:
            self.label_17.setText('Can Not Delete Default Contact')
        else:
            contact_code = self.lineEdit_concode.text()
            question_mark = QtWidgets.QMessageBox
            choice_del = question_mark.question(self, 'warning', 'Are You Sure Delete This Client Contact', question_mark.Yes | question_mark.No)
            if choice_del == question_mark.Yes:
                self.cur.execute('''
                    DELETE FROM client_contact WHERE contact_code=%s
                ''', (contact_code,))
                self.db.commit()
                self.contactcombedclname.clear()
                self.concombedcona.clear()
                self.lineEdit_concode.clear()
                self.label_17.clear()
                self.Show_Contact()
                QMessageBox.information(self, 'success', 'Client Contact Has Been Deleted')

    def Table8_Column_Width(self):
        self.tableWidget_8.setColumnWidth(0, 150)
        self.tableWidget_8.setColumnWidth(1, 200)
        self.tableWidget_8.setColumnWidth(2, 130)
        self.tableWidget_8.setColumnWidth(3, 250)
        self.tableWidget_8.setColumnWidth(4, 100)
        self.tableWidget_8.setColumnWidth(5, 100)

    def Show_Branch(self):
        self.Table8_Column_Width()
        self.tableWidget_8.setRowCount(0)
        self.tableWidget_8.insertRow(0)
        self.cur.execute('''
            SELECT client_name, branch_name, branch_code, branch_address, join_date, edit_date FROM client_branch
        ''')
        client_branches = self.cur.fetchall()
        for row, form in enumerate(client_branches):
            for col, item in enumerate(form):
                self.tableWidget_8.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_8.rowCount()
            self.tableWidget_8.insertRow(row_position) 

    def Show_Client_Branshes(self):
        client_name = self.editbracoboxclname.currentText()
        self.tableWidget_8.setRowCount(0)
        self.tableWidget_8.insertRow(0)
        self.cur.execute('''
            SELECT client_name, branch_name, branch_code, branch_address, join_date, edit_date FROM client_branch WHERE client_name=%s
        ''', (client_name,))
        client_information = self.cur.fetchall()
        for row, form in enumerate(client_information):
            for col, item in enumerate(form):
                self.tableWidget_8.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_8.rowCount()
            self.tableWidget_8.insertRow(row_position)

    def Show_Branch_Name(self):
        self.branchcomname.clear()
        dash = ''
        self.branchcomname.addItem(dash)
        client_name = self.editbracoboxclname.currentText()
        self.cur.execute('''
            SELECT branch_name FROM client_branch WHERE client_name=%s
        ''', (client_name,))
        all_branches = self.cur.fetchall()
        for branches in all_branches:
            self.branchcomname.addItem(branches[0])
            self.branchcomname.activated.connect(self.Show_Branch_Code)

    def Show_Branch_Code(self):
        branch_name = self.branchcomname.currentText()
        if len(branch_name)==0:
            self.lineEdit_59.setText('')
            self.textEdit_11.setText('')
        else:
            self.cur.execute('''
                SELECT branch_code FROM client_branch WHERE branch_name=%s
            ''', (branch_name,))
            branch_code = self.cur.fetchone()
            self.lineEdit_58.setText(str(branch_code[0]))

    def Search_Branch(self):
        branch_code = self.lineEdit_58.text()
        self.cur.execute('''
            SELECT branch_name, branch_address FROM client_branch WHERE branch_code=%s
        ''', (branch_code,))
        info_branchs = self.cur.fetchall()
        for info in info_branchs:
            self.lineEdit_59.setText(info[0])
            self.textEdit_11.setText(info[1])
        self.tableWidget_8.setRowCount(0)
        self.tableWidget_8.insertRow(0)
        self.cur.execute('''
            SELECT client_name, branch_name, branch_code, branch_address, join_date, edit_date FROM client_branch WHERE branch_code=%s
        ''', (branch_code,))
        client_information = self.cur.fetchall()
        for row, form in enumerate(client_information):
            for col, item in enumerate(form):
                self.tableWidget_8.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_8.rowCount()
            self.tableWidget_8.insertRow(row_position)
    
    def Clear_Search_Branch(self):
        self.editbracoboxclname.clear()
        self.branchcomname.clear()
        self.lineEdit_58.clear()
        self.lineEdit_59.clear()
        self.textEdit_11.clear()
        self.label_32.clear()
        self.Show_Branch()
        self.Activate_Combo()
        self.Show_Branch_Name()

    def Edit_Client_Branch(self):
        branch_name = self.lineEdit_59.text()
        branch_address = self.textEdit_11.toPlainText()
        branch_code = self.lineEdit_58.text()
        edit_date = self.dateEdit_20.text()
        if len(branch_name)==0 and len(branch_address)==0:
            self.label_32.setText('Please Fillin all necessary field/fields')
        else:
            self.cur.execute('''
                UPDATE client_branch SET branch_name=%s, branch_address=%s, edit_date=%s WHERE branch_code=%s
            ''', (branch_name, branch_address, edit_date, branch_code))
            self.db.commit()
            QMessageBox.information(self, 'success', 'Client Branch Has Been Updated!')
            self.Search_Branch()
            self.lineEdit_59.clear()
            self.textEdit_11.clear()
    
    def Clear_Edit_Branch(self):
        self.editbracoboxclname.setCurrentIndex(0)
        self.branchcomname.setCurrentIndex(0)
        self.lineEdit_59.clear()
        self.textEdit_11.clear()
        self.lineEdit_58.clear()
        self.label_32.clear()

    def Delete_Branch(self):
        branch_code = self.lineEdit_58.text()
        if len(branch_code)==0:
            self.label_32.setText('Please Select Branch to Delete!')
        else:
            self.branchcomname.clear()
            client_name = self.editbracoboxclname.currentText()
            self.cur.execute('''
                SELECT branch_name FROM client_branch WHERE client_name=%s
            ''', (client_name,))
            all_branches = self.cur.fetchall()
            for item in all_branches:
                for conname in item:
                    self.branchcomname.addItem(conname)
            count = self.branchcomname.count()
            if count==1:
                self.label_32.setText('Can Not Delete Default Branch')
            else:
                branch_code = self.lineEdit_58.text()
                question_mark = QtWidgets.QMessageBox
                choice_del = question_mark.question(self, 'warning', 'Are You Sure Delete This Client Branch', question_mark.Yes | question_mark.No)
                if choice_del == question_mark.Yes:
                    self.cur.execute('''
                        DELETE FROM client_branch WHERE branch_code=%s
                    ''', (branch_code,))
                    self.db.commit()
                    self.Clear_Edit_Branch()
                    self.Show_Branch()
                    self.Activate_Combo()
                    QMessageBox.information(self, 'success', 'Client Branch Has Been Deleted')

    def Show_Client_CoMaEd(self):
        client_name = self.machincoboxclname_2.currentText()
        if len(client_name)==0:
            self.lineEdit_clcode_2.setText('')
        else:
            self.cur.execute('''
                SELECT client_code FROM new_client WHERE client_name=%s
            ''', (client_name,))
            all_code = self.cur.fetchone()
            for code in all_code:
                self.lineEdit_clcode_2.setText(str(code))

    def Show_Branch_COMaEd(self):
        branch_name = self.machinecombrname_2.currentText()
        if len(branch_name)==0:
            self.lineEdit_60.setText('')
        else:
            self.cur.execute('''
                SELECT branch_code FROM client_branch WHERE branch_name=%s
            ''', (branch_name,))
            all_code = self.cur.fetchone()
            for code in all_code:
                self.lineEdit_60.setText(str(code))
    
    def Show_Machine_InfoMaEd(self):
        branch_name = self.machinecombrname_2.currentText()
        machine_type = self.machinetype_2.currentText()
        if machine_type=='Coffe Machine' or machine_type=='Graneta Machine':
            self.cur.execute('''
                SELECT machine_model, machine_serial, machine_group FROM client_machine WHERE branch_name=%s AND machine_type=%s 
            ''', (branch_name, machine_type))
            all_info2 = self.cur.fetchall()
            for info2 in all_info2:
                self.lineEdit_95.setText(info2[0])
                self.lineEdit_94.setText(info2[1])
                if info2[2]=='1 Group':
                    self.groupsnum_2.setCurrentIndex(1)
                elif info2[2]=='2 Groups':
                    self.groupsnum_2.setCurrentIndex(2)
                elif info2[2]=='3 Groups':
                    self.groupsnum_2.setCurrentIndex(3)
        else:
            self.cur.execute('''
                SELECT machine_model, machine_serial FROM client_machine WHERE branch_name=%s AND machine_type=%s
            ''', (branch_name, machine_type))
            all_info = self.cur.fetchall()
            for info in all_info:
                self.lineEdit_95.setText(info[0])
                self.lineEdit_94.setText(info[1])

    def Table9_Column_Width(self):
        self.tableWidget_9.setColumnWidth(0, 150)
        self.tableWidget_9.setColumnWidth(1, 200)
        self.tableWidget_9.setColumnWidth(2, 150)
        self.tableWidget_9.setColumnWidth(3, 200)
        self.tableWidget_9.setColumnWidth(4, 150)
        self.tableWidget_9.setColumnWidth(5, 155)
        self.tableWidget_9.setColumnWidth(6, 100)
        self.tableWidget_9.setColumnWidth(7, 100)

    def Show_Machine(self):
        self.Table9_Column_Width()
        self.tableWidget_9.setRowCount(0)
        self.tableWidget_9.insertRow(0)
        self.cur.execute('''
            SELECT client_name, branch_name, machine_type, machine_model, machine_serial, machine_group, join_date, edit_date FROM client_machine
        ''')
        machines = self.cur.fetchall()
        for row1, form1 in enumerate(machines):
            for col1, item1 in enumerate(form1):
                self.tableWidget_9.setItem(row1, col1, QTableWidgetItem(str(item1)))
                col1 += 1
            row_position = self.tableWidget_9.rowCount()
            self.tableWidget_9.insertRow(row_position)

    def Search_Bytype_Machine(self):
        client_name = self.machincoboxclname_2.currentText()
        branch_name = self.machinecombrname_2.currentText()
        machine_type = self.machinetype_2.currentText()
        if len(branch_name)==0:
            self.tableWidget_9.setRowCount(0)
            self.tableWidget_9.insertRow(0)
            self.cur.execute('''
                SELECT client_name, branch_name, machine_type, machine_model, machine_serial, machine_group, join_date, edit_date FROM client_machine WHERE client_name=%s
            ''', (client_name,))
            all_data = self.cur.fetchall()
            for row2, form2 in enumerate(all_data):
                for col2, item2 in enumerate(form2):
                    self.tableWidget_9.setItem(row2, col2, QTableWidgetItem(str(item2)))
                    col2 += 1
                row_position = self.tableWidget_9.rowCount()
                self.tableWidget_9.insertRow(row_position)
        elif len(machine_type)==0:
            self.tableWidget_9.setRowCount(0)
            self.tableWidget_9.insertRow(0)
            self.cur.execute('''
                SELECT client_name, branch_name, machine_type, machine_model, machine_serial, machine_group, join_date, edit_date FROM client_machine WHERE branch_name=%s
            ''', (branch_name,))
            branches = self.cur.fetchall()
            for row, form in enumerate(branches):
                for col, item in enumerate(form):
                    self.tableWidget_9.setItem(row, col, QTableWidgetItem(str(item)))
                    col += 1
                row_position = self.tableWidget_9.rowCount()
                self.tableWidget_9.insertRow(row_position)
        else:
            self.tableWidget_9.setRowCount(0)
            self.tableWidget_9.insertRow(0)
            self.cur.execute('''
                SELECT client_name, branch_name, machine_type, machine_model, machine_serial, machine_group, join_date, edit_date FROM client_machine WHERE client_name=%s AND branch_name=%s AND machine_type=%s
            ''', (client_name, branch_name, machine_type,))
            machines = self.cur.fetchall()
            for row, form in enumerate(machines):
                for col, item in enumerate(form):
                    self.tableWidget_9.setItem(row, col, QTableWidgetItem(str(item)))
                    col += 1
                row_position = self.tableWidget_9.rowCount()
                self.tableWidget_9.insertRow(row_position)
            self.Show_Machine_InfoMaEd()

    def Clear_Search_Machine(self):
        self.machincoboxclname_2.setCurrentIndex(0)
        self.machinecombrname_2.setCurrentIndex(0)
        self.label_35.clear()
        self.lineEdit_95.clear()
        self.lineEdit_94.clear()
        self.lineEdit_clcode_2.clear()
        self.lineEdit_60.clear()
        self.machinetype_2.setCurrentIndex(0)
        self.groupsnum_2.setCurrentIndex(0)
        self.Activate_Combo()
        self.Show_Machine()

    def Edit_Client_Machine(self):
        branch_name = self.machinecombrname_2.currentText()
        machine_type = self.machinetype_2.currentText()
        machine_model = self.lineEdit_95.text()
        machine_serial = self.lineEdit_94.text()
        group_nums = self.groupsnum_2.currentText()
        edit_date = self.dateEdit_21.text()

        if len(branch_name)==0 or len(machine_type)==0:
            self.label_35.setText('Select Branch Name And Machine Type!')
        elif len(machine_model)==0:
            self.label_35.setText('Insert Machine Model!')
        elif machine_type=='Coffe Machine' and machine_type=='Graneta Machine':
            self.label_35.setText('Please Select Groups Numbers!')
        else:
            self.cur.execute('''
                UPDATE client_machine SET machine_model=%s, machine_serial=%s, machine_group=%s, edit_date=%s WHERE branch_name=%s AND machine_type=%s
            ''', (machine_model, machine_serial, group_nums, edit_date, branch_name, machine_type))
            self.db.commit()
            QMessageBox.information(self, 'success', 'Client Machine Has Been Updated')
            self.Clear_Search_Machine()

    def Delete_Machine(self):
        branch_name = self.machinecombrname_2.currentText()
        machine_type = self.machinetype_2.currentText()
        question_mark = QtWidgets.QMessageBox
        choice_del = question_mark.question(self, 'warning', 'Are You Sure Delete This Client Machine', question_mark.Yes | question_mark.No)
        if choice_del == question_mark.Yes:
            if len(machine_type)==0:
                self.label_35.setText('Please Select Machine Type!')
            else:
                self.cur.execute('''
                    DELETE FROM client_machine WHERE branch_name=%s AND machine_type=%s
                ''', (branch_name, machine_type))
                self.db.commit()
                self.Show_Machine()
                QMessageBox.information(self, 'success', 'Client Machine Has Been Deleted')    

    def Show_Technician_History(self):
        self.tableWidget_3.setRowCount(0)
        self.tableWidget_3.insertRow(0)
        self.cur.execute('''
            SELECT call_number, user_name,  technician_name, client_name, branch_name, call_type, machine_type, recieve_date FROM technician_calls
        ''')
        allcalls = self.cur.fetchall()
        for row, form in enumerate(allcalls):
            for col, item in enumerate(form):
                self.tableWidget_3.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_3.rowCount()
            self.tableWidget_3.insertRow(row_position)

    def Search_Technician_Jobs(self):
        tech_hist = self.techcomhist.currentText()
        self.tableWidget_3.setRowCount(0)
        self.tableWidget_3.insertRow(0)
        self.cur.execute('''
            SELECT call_number, user_name,  technician_name, client_name, branch_name, call_type, machine_type, recieve_date FROM technician_calls WHERE technician_name=%s
        ''', (tech_hist,))
        allcalls = self.cur.fetchall()
        for row, form in enumerate(allcalls):
            for col, item in enumerate(form):
                self.tableWidget_3.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_3.rowCount()
            self.tableWidget_3.insertRow(row_position)

    def Clear_Technician_Jobs(self):
        self.Show_Technician_History()

    def Export_Tech_Calls(self):
        technician_name = self.techcomhist.currentText()
        if technician_name!=0:
            self.Export_TechName_Calls()
        else:
            self.cur.execute('''
                SELECT call_number, user_name,  technician_name, client_name, branch_name, call_type, machine_type, recieve_date FROM technician_calls 
            ''')
            all_techs = self.cur.fetchall()
            filename = QFileDialog.getSaveFileName(self, 'Save File', 'C:', ".xls(*.xls)")
            my_filet = filename[0]
            excel_file = xlwt.Workbook()
            self.sheet1 = excel_file.add_sheet('Tehnician Calls', cell_overwrite_ok=True)
            self.sheet1.write(0, 0, 'Call Number')
            self.sheet1.write(0, 1, 'User Name')
            self.sheet1.write(0, 2, 'Technician Name')
            self.sheet1.write(0, 3, 'Client Name')
            self.sheet1.write(0, 4, 'Branch Name')
            self.sheet1.write(0, 5, 'Call Type')
            self.sheet1.write(0, 6, 'Machine Type')
            self.sheet1.write(0, 7, 'Date')
            row_number = 1
            for row in all_techs:
                col = 0
                for item in row:
                    self.sheet1.write(row_number, col, str(item))
                    col += 1
                row_number += 1
            if my_filet=='':
                self.Technician_Tab()
            else:
                excel_file.save(my_filet)

    def Export_TechName_Calls(self):
        technician_name = self.techcomhist.currentText()
        self.cur.execute('''
            SELECT call_number, user_name,  technician_name, client_name, branch_name, call_type, machine_type, recieve_date FROM technician_calls WHERE technician_name=%s
        ''', (technician_name,))
        all_techs = self.cur.fetchall()
        filename = QFileDialog.getSaveFileName(self, 'Save File', 'C:', ".xls(*.xls)")
        my_filet = filename[0]
        excel_file = xlwt.Workbook()
        self.sheet1 = excel_file.add_sheet('Tehnician Calls', cell_overwrite_ok=True)
        self.sheet1.write(0, 0, 'Call Number')
        self.sheet1.write(0, 1, 'User Name')
        self.sheet1.write(0, 2, 'Technician Name')
        self.sheet1.write(0, 3, 'Client Name')
        self.sheet1.write(0, 4, 'Branch Name')
        self.sheet1.write(0, 5, 'Call Type')
        self.sheet1.write(0, 6, 'Machine Type')
        self.sheet1.write(0, 7, 'Date')
        row_number = 1
        for row in all_techs:
            col = 0
            for item in row:
                self.sheet1.write(row_number, col, str(item))
                col += 1
            row_number += 1
        if my_filet=='':
            self.Technician_Tab()
        else:
            excel_file.save(my_filet)

    def Email_To_Send(self):
        recipient_email2 = self.lineEdit.text()
        subject2 = self.lineEdit_2.text()
        content2 = self.lineEdit_3.text()
        
        filename = QFileDialog.getOpenFileName(self, 'Open File', 'C:', ".xls(*.xls)")
        my_file = filename[0]

        if my_file=='':
            self.Send_Email_Page()
        else:
            send_mail_with_excel(recipient_email2, subject2, content2, my_file)

    def Whats_Message(self):
        try:
            import pywhatkit
        except:
            self.label_37.setText("No internet connection available.")
        else:
            country_key = "+2"
            mobile_num = self.techmobile.text()
            mobile_complete = country_key + mobile_num
            date = datetime.datetime.now()
            client_name = self.clientcoboxname.currentText()
            branch_name = self.branchcoboxname.currentText()
            machine_type = self.machinetypecalls.currentText()
            client_complain = self.clientcomplain.toPlainText()
            client_address = self.branchaddress.toPlainText()
            engineer = self.callsusername.text()
            message_body = f"Date: {date}\nClient Name: {client_name}\nBranch Name: {branch_name}\nMachine_Type: {machine_type}\nClient Complain: {client_complain}\nClient Address: {client_address}\nEngineer: {engineer}"
            now = datetime.datetime.now()
            get_hour = now.hour
            get_minute = now.minute+2
            if len(mobile_num)==0 or len(client_name)==0 or len(branch_name)==0 or len(client_complain)==0 or len(client_address)==0 or len(engineer)==0:
                self.label_37.setText("Please Fillout All Fields")
            else:
                get_secondssex = 105-now.second
                self.label_181.setText(f"In {get_secondssex} seconds Whats will open and after 15 Seconds Message will be Delivered!")
                pywhatkit.sendwhatmsg(mobile_complete, message_body, get_hour, get_minute)

    def Show_Brname_Month(self):
        self.visittype.setCurrentIndex(0)
        self.technamemonth.setCurrentIndex(0)
        self.brnamemonth.clear()
        cli_name = self.clnamemonth.currentText()
        self.cur.execute('''
            SELECT branch_name FROM client_branch WHERE client_name=%s ORDER BY branch_name
        ''', (cli_name,))
        many_bran = self.cur.fetchall()
        self.brnamemonth.clear()
        dash = ''
        self.brnamemonth.addItem(dash)
        for branch in many_bran:
            self.brnamemonth.addItems(branch)

    def Add_Monthly_Followup(self):
        date_execute = self.dateEditmonth.text()
        client_name = self.clnamemonth.currentText()
        visit_type = self.visittype.currentText()
        tech_name = self.technamemonth.currentText()
        branch_name = self.brnamemonth.currentText()
        if len(client_name)==0 or len(visit_type)==0 or len(tech_name)==0 or len(branch_name)==0:
            self.label_204.setText('All Fields Are Required!')
        else:
            self.cur.execute('''
                INSERT INTO monthly_followup(date, client_name, tech_name, branch_name, visit_type)
                VALUES(%s, %s, %s, %s, %s)
            ''', (date_execute, client_name, tech_name, branch_name, visit_type))
            self.db.commit()
            self.Show_Monthly_Followup()
            QMessageBox.information(self, 'success', 'New Monthly Followup Has Been Added!')
            self.Clear_Selection_Monthly()

    def Show_Monthly_Followup(self):
        self.tableWidget_5.setRowCount(0)
        self.tableWidget_5.insertRow(0)
        self.cur.execute('''
            SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup
        ''')
        allcalls = self.cur.fetchall()
        for row, form in enumerate(allcalls):
            for col, item in enumerate(form):
                self.tableWidget_5.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_5.rowCount()
            self.tableWidget_5.insertRow(row_position)

    def Clear_Selection_Monthly(self):
        self.label_204.clear()
        self.visittype.setCurrentIndex(0)
        self.technamemonth.setCurrentIndex(0)
        self.clnamemonth.setCurrentIndex(0)
        self.brnamemonth.setCurrentIndex(0)

    def Show_Branch_TecMon(self):
        clie_name = self.clnamemonth_2.currentText()
        self.cur.execute('''
            SELECT branch_name FROM client_branch WHERE client_name=%s ORDER BY branch_name
        ''', (clie_name,))
        many_bran = self.cur.fetchall()
        self.brnamemonth_2.clear()
        dash = ''
        self.brnamemonth_2.addItem(dash)
        for branch in many_bran:
            self.brnamemonth_2.addItems(branch)

    def Search_Tech_Monthly(self):
        curr_date = self.dateEditmonth_2.text()
        tech_name = self.technamemonth_2.currentText()
        cli_name = self.clnamemonth_2.currentText()
        bran_name = self.brnamemonth_2.currentText()
        visit_type = self.visittype_2.currentText()

        if len(tech_name)==0:
            self.label_211.setText("Select Technician Name!")
        else:
            self.cur.execute('''
                SELECT date FROM monthly_followup WHERE tech_name=%s
            ''', (tech_name,))
            all_date = self.cur.fetchall()
            date_list = []
            for one_date in all_date:
                for single_date in one_date:
                    date_list.append(single_date)
            if curr_date in date_list:
                self.Search_ByDate()
            else:
                if len(tech_name)!=0 and len(cli_name)==0 and len(bran_name)==0 and len(visit_type)==0:
                    self.Search_ByTech_Name()
                elif len(tech_name)!=0 and len(cli_name)!=0 and len(bran_name)==0 and len(visit_type)==0:
                    self.Search_ByTech_ClName()
                elif len(tech_name)!= 0 and len(cli_name)!=0 and len(bran_name)!=0 and len(visit_type)==0:
                    self.Search_ByTech_Cl_BrName()
                elif len(tech_name)!= 0 and len(cli_name)!=0 and len(bran_name)!=0 and len(visit_type)!=0:
                    self.Search_ByTech_Cl_Br_Visit()
                elif len(tech_name)!= 0 and len(cli_name)==0 and len(bran_name)==0 and len(visit_type)!=0:
                    self.Search_ByTech_Visit()
                elif len(tech_name)!= 0 and len(cli_name)!=0 and len(bran_name)==0 and len(visit_type)!=0:
                    self.Search_ByTech_Cl_Visit()
    
    def Search_ByDate(self):
        tech_name = self.technamemonth_2.currentText()
        curr_date = self.dateEditmonth_2.text()
        curr_date2 = self.dateEditmonth_3.text()
        if len(tech_name)==0:
                self.label_211.setText("Select Technician Name!")
        else:
            self.label_211.clear()
            self.tableWidget_6.setRowCount(0)
            self.tableWidget_6.insertRow(0)
            self.cur.execute('''
                SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s AND date BETWEEN %s And %s
            ''', (tech_name, curr_date, curr_date2))
            allcalls = self.cur.fetchall()
            for row, form in enumerate(allcalls):
                for col, item in enumerate(form):
                    self.tableWidget_6.setItem(row, col, QTableWidgetItem(str(item)))
                    col += 1
                row_position = self.tableWidget_6.rowCount()
                self.tableWidget_6.insertRow(row_position)

    def Search_ByTech_Name(self):
        tech_name = self.technamemonth_2.currentText()
        self.tableWidget_6.setRowCount(0)
        self.tableWidget_6.insertRow(0)
        self.cur.execute('''
            SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s
        ''', (tech_name,))
        allcalls = self.cur.fetchall()
        for row, form in enumerate(allcalls):
            for col, item in enumerate(form):
                self.tableWidget_6.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_6.rowCount()
            self.tableWidget_6.insertRow(row_position)

    def Search_ByTech_ClName(self):
        tech_name = self.technamemonth_2.currentText()
        cli_name = self.clnamemonth_2.currentText()
        self.tableWidget_6.setRowCount(0)
        self.tableWidget_6.insertRow(0)
        self.cur.execute('''
            SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s AND client_name=%s
        ''', (tech_name, cli_name))
        allcalls = self.cur.fetchall()
        for row, form in enumerate(allcalls):
            for col, item in enumerate(form):
                self.tableWidget_6.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_6.rowCount()
            self.tableWidget_6.insertRow(row_position)

    def Search_ByTech_Cl_BrName(self):
        tech_name = self.technamemonth_2.currentText()
        cli_name = self.clnamemonth_2.currentText()
        bran_name = self.brnamemonth_2.currentText()
        self.tableWidget_6.setRowCount(0)
        self.tableWidget_6.insertRow(0)
        self.cur.execute('''
            SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s AND client_name=%s AND branch_name=%s
        ''', (tech_name, cli_name, bran_name))
        allcalls = self.cur.fetchall()
        for row, form in enumerate(allcalls):
            for col, item in enumerate(form):
                self.tableWidget_6.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_6.rowCount()
            self.tableWidget_6.insertRow(row_position)

    def Search_ByTech_Cl_Br_Visit(self):
        tech_name = self.technamemonth_2.currentText()
        cli_name = self.clnamemonth_2.currentText()
        bran_name = self.brnamemonth_2.currentText()
        visit_type = self.visittype_2.currentText()
        self.tableWidget_6.setRowCount(0)
        self.tableWidget_6.insertRow(0)
        self.cur.execute('''
            SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s AND client_name=%s AND branch_name=%s AND visit_type=%s
        ''', (tech_name, cli_name, bran_name, visit_type))
        allcalls = self.cur.fetchall()
        for row, form in enumerate(allcalls):
            for col, item in enumerate(form):
                self.tableWidget_6.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_6.rowCount()
            self.tableWidget_6.insertRow(row_position)

    def Search_ByTech_Visit(self):
        tech_name = self.technamemonth_2.currentText()
        visit_type = self.visittype_2.currentText()
        self.tableWidget_6.setRowCount(0)
        self.tableWidget_6.insertRow(0)
        self.cur.execute('''
            SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s AND visit_type=%s
        ''', (tech_name, visit_type))
        allcalls = self.cur.fetchall()
        for row, form in enumerate(allcalls):
            for col, item in enumerate(form):
                self.tableWidget_6.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_6.rowCount()
            self.tableWidget_6.insertRow(row_position)
    
    def Search_ByTech_Cl_Visit(self):
        tech_name = self.technamemonth_2.currentText()
        cli_name = self.clnamemonth_2.currentText()
        visit_type = self.visittype_2.currentText()
        self.tableWidget_6.setRowCount(0)
        self.tableWidget_6.insertRow(0)
        self.cur.execute('''
            SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s AND client_name=%s AND visit_type=%s
        ''', (tech_name, cli_name, visit_type))
        allcalls = self.cur.fetchall()
        for row, form in enumerate(allcalls):
            for col, item in enumerate(form):
                self.tableWidget_6.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_6.rowCount()
            self.tableWidget_6.insertRow(row_position)

    def Clear_Search_Tech_Monthly(self):
        self.label_211.clear()
        self.technamemonth_2.setCurrentIndex(0)
        self.brnamemonth_2.clear()
        self.clnamemonth_2.setCurrentIndex(0)
        self.visittype_2.setCurrentIndex(0)

        d2 = '01-01-2000'
        start_date = datetime.datetime.strptime(d2, '%d-%m-%Y')
        self.dateEditmonth_2.setDate(start_date)
        self.dateEditmonth_3.setDate(start_date)

        self.tableWidget_6.setRowCount(0)
        self.tableWidget_6.insertRow(0)
        self.cur.execute('''
            SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup
        ''')
        allcalls = self.cur.fetchall()
        for row, form in enumerate(allcalls):
            for col, item in enumerate(form):
                self.tableWidget_6.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_6.rowCount()
            self.tableWidget_6.insertRow(row_position)

    def Export_Monthly_Record(self):
        tech_name = self.technamemonth_2.currentText()
        if len(tech_name)==0:
            self.Export_All_Record()
        else:
            self.Export_By_TechName()

    def Export_By_TechName(self):
        tech_name = self.technamemonth_2.currentText()
        self.cur.execute('''
            SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s
        ''', (tech_name,))
        all_records = self.cur.fetchall()
        filename = QFileDialog.getSaveFileName(self, 'Save File', 'C:', ".xls(*.xls)")
        my_filet = filename[0]
        excel_file = xlwt.Workbook()
        self.sheet1 = excel_file.add_sheet('Tehnician Monthly Record', cell_overwrite_ok=True)
        self.sheet1.write(0, 0, 'Executed Date')
        self.sheet1.write(0, 1, 'Client Name')
        self.sheet1.write(0, 2, 'Technician Name')
        self.sheet1.write(0, 3, 'Branch Name')
        self.sheet1.write(0, 4, 'Visit Type')
        row_number = 1
        for row in all_records:
            col = 0
            for item in row:
                self.sheet1.write(row_number, col, str(item))
                col += 1
            row_number += 1
        if my_filet=='':
            self.Search_Delete_Monthly_Tab()
        else:
            excel_file.save(my_filet)
        QMessageBox.information(self, 'success', f'Monthly Report For Technician: {tech_name} Has Been Created!')
    
    def Export_All_Record(self):
        self.cur.execute('''
            SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup
        ''')
        all_records = self.cur.fetchall() 
        filename = QFileDialog.getSaveFileName(self, 'Save File', 'C:', ".xls(*.xls)")
        my_filet = filename[0]
        excel_file = xlwt.Workbook()
        self.sheet1 = excel_file.add_sheet('Tehnician Monthly Record', cell_overwrite_ok=True)
        self.sheet1.write(0, 0, 'Executed Date')
        self.sheet1.write(0, 1, 'Client Name')
        self.sheet1.write(0, 2, 'Technician Name')
        self.sheet1.write(0, 3, 'Branch Name')
        self.sheet1.write(0, 4, 'Visit Type')
        row_number = 1
        for row in all_records:
            col = 0
            for item in row:
                self.sheet1.write(row_number, col, str(item))
                col += 1
            row_number += 1
        if my_filet=='':
            self.Search_Delete_Monthly_Tab()
        else:
            excel_file.save(my_filet)
        QMessageBox.information(self, 'success', 'Monthly Report For All Technicians Has Been Created!')

    def Show_Branch_DelMon(self):
        clie_name = self.clnamemonth_3.currentText()
        self.cur.execute('''
            SELECT branch_name FROM client_branch WHERE client_name=%s ORDER BY branch_name
        ''', (clie_name,))
        many_bran = self.cur.fetchall()
        self.brnamemonth_3.clear()
        dash = ''
        self.brnamemonth_3.addItem(dash)
        for branch in many_bran:
            self.brnamemonth_3.addItems(branch)
    
    def Show_Monthly_Record_Todelete(self):
        tech_name = self.technamemonth_3.currentText()
        date_from = self.dateEditmonth_5.text()
        date_to = self.dateEditmonth_4.text()
        clie_name = self.clnamemonth_3.currentText()
        bran_name = self.brnamemonth_3.currentText()
        visit_type = self.visittype_3.currentText()

        self.cur.execute('''
            SELECT date FROM monthly_followup WHERE tech_name=%s
        ''', (tech_name,))
        all_date = self.cur.fetchall()
        date_list = []
        for one_date in all_date:
            for single_date in one_date:
                date_list.append(single_date)
        if date_from in date_list:
            if len(tech_name)==0:
                self.label_215.setText("Select Technician Name!")
            else:
                self.label_215.clear()
                self.tableWidget_11.setRowCount(0)
                self.tableWidget_11.insertRow(0)
                self.cur.execute('''
                    SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s AND date BETWEEN %s And %s
                ''', (tech_name, date_from, date_to))
                allcalls = self.cur.fetchall()
                for row, form in enumerate(allcalls):
                    for col, item in enumerate(form):
                        self.tableWidget_11.setItem(row, col, QTableWidgetItem(str(item)))
                        col += 1
                    row_position = self.tableWidget_11.rowCount()
                    self.tableWidget_11.insertRow(row_position)
        else:
            if len(tech_name)!=0 and len(clie_name)==0 and len(bran_name)==0 and len(visit_type)==0:
                self.label_215.setText("Selected Date 'From' Not In Database!")
                self.tableWidget_11.setRowCount(0)
                self.tableWidget_11.insertRow(0)
                self.cur.execute('''
                    SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s
                ''', (tech_name,))
                allcalls = self.cur.fetchall()
                for row, form in enumerate(allcalls):
                    for col, item in enumerate(form):
                        self.tableWidget_11.setItem(row, col, QTableWidgetItem(str(item)))
                        col += 1
                    row_position = self.tableWidget_11.rowCount()
                    self.tableWidget_11.insertRow(row_position)
            elif len(tech_name)!= 0 and len(clie_name)!=0 and len(bran_name)==0 and len(visit_type)==0:
                self.label_215.setText("Selected Date 'From' Not In Database!")
                self.tableWidget_11.setRowCount(0)
                self.tableWidget_11.insertRow(0)
                self.cur.execute('''
                    SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s AND client_name=%s
                ''', (tech_name, clie_name))
                allcalls = self.cur.fetchall()
                for row, form in enumerate(allcalls):
                    for col, item in enumerate(form):
                        self.tableWidget_11.setItem(row, col, QTableWidgetItem(str(item)))
                        col += 1
                    row_position = self.tableWidget_11.rowCount()
                    self.tableWidget_11.insertRow(row_position)
            elif len(tech_name)!= 0 and len(clie_name)!=0 and len(bran_name)!=0 and len(visit_type)==0:
                self.label_215.setText("Selected Date 'From' Not In Database!")
                self.tableWidget_11.setRowCount(0)
                self.tableWidget_11.insertRow(0)
                self.cur.execute('''
                    SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s AND client_name=%s AND branch_name=%s
                ''', (tech_name, clie_name, bran_name))
                allcalls = self.cur.fetchall()
                for row, form in enumerate(allcalls):
                    for col, item in enumerate(form):
                        self.tableWidget_11.setItem(row, col, QTableWidgetItem(str(item)))
                        col += 1
                    row_position = self.tableWidget_11.rowCount()
                    self.tableWidget_11.insertRow(row_position)
            elif len(tech_name)!= 0 and len(clie_name)!=0 and len(bran_name)!=0 and len(visit_type)!=0:
                self.label_215.setText("Selected Date 'From' Not In Database!")
                self.tableWidget_11.setRowCount(0)
                self.tableWidget_11.insertRow(0)
                self.cur.execute('''
                    SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s AND client_name=%s AND branch_name=%s AND visit_type=%s
                ''', (tech_name, clie_name, bran_name, visit_type))
                allcalls = self.cur.fetchall()
                for row, form in enumerate(allcalls):
                    for col, item in enumerate(form):
                        self.tableWidget_11.setItem(row, col, QTableWidgetItem(str(item)))
                        col += 1
                    row_position = self.tableWidget_11.rowCount()
                    self.tableWidget_11.insertRow(row_position)
            elif len(tech_name)!= 0 and len(clie_name)==0 and len(bran_name)==0 and len(visit_type)!=0:
                self.label_215.setText("Selected Date 'From' Not In Database!")
                self.tableWidget_11.setRowCount(0)
                self.tableWidget_11.insertRow(0)
                self.cur.execute('''
                    SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s AND visit_type=%s
                ''', (tech_name, visit_type))
                allcalls = self.cur.fetchall()
                for row, form in enumerate(allcalls):
                    for col, item in enumerate(form):
                        self.tableWidget_11.setItem(row, col, QTableWidgetItem(str(item)))
                        col += 1
                    row_position = self.tableWidget_11.rowCount()
                    self.tableWidget_11.insertRow(row_position)
            elif len(tech_name)!= 0 and len(clie_name)!=0 and len(bran_name)==0 and len(visit_type)!=0:
                self.label_215.setText("Selected Date 'From' Not In Database!")
                self.tableWidget_11.setRowCount(0)
                self.tableWidget_11.insertRow(0)
                self.cur.execute('''
                    SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup WHERE tech_name=%s AND client_name=%s AND visit_type=%s
                ''', (tech_name, clie_name, visit_type))
                allcalls = self.cur.fetchall()
                for row, form in enumerate(allcalls):
                    for col, item in enumerate(form):
                        self.tableWidget_11.setItem(row, col, QTableWidgetItem(str(item)))
                        col += 1
                    row_position = self.tableWidget_11.rowCount()
                    self.tableWidget_11.insertRow(row_position)

    def Clear_Search_Monthly_Todelete(self):
        self.label_215.clear()
        self.clnamemonth_3.setCurrentIndex(0)
        self.technamemonth_3.setCurrentIndex(0)
        self.visittype_3.setCurrentIndex(0)
        self.brnamemonth_3.setCurrentIndex(0)
        d2 = '01-01-2000'
        start_date = datetime.datetime.strptime(d2, '%d-%m-%Y')
        self.dateEditmonth_5.setDate(start_date)
        self.dateEditmonth_4.setDate(start_date)
        
        self.tableWidget_11.setRowCount(0)
        self.tableWidget_11.insertRow(0)
        self.cur.execute('''
            SELECT date, client_name, tech_name, branch_name, visit_type FROM monthly_followup
        ''')
        allcalls = self.cur.fetchall()
        for row, form in enumerate(allcalls):
            for col, item in enumerate(form):
                self.tableWidget_11.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            row_position = self.tableWidget_11.rowCount()
            self.tableWidget_11.insertRow(row_position)

    def Delete_Monthly_Follow(self):
        tech_name = self.technamemonth_3.currentText()
        clie_name = self.clnamemonth_3.currentText()
        bran_name = self.brnamemonth_3.currentText()
        if len(tech_name)==0 or len(clie_name)==0 or len(bran_name)==0:
            self.label_215.setText("Please Select Technician Name, Client Name and Branch Name!")
        else:
            question_mark = QtWidgets.QMessageBox
            choice_del = question_mark.question(self, 'warning', 'Are You Sure Delete This Record?', question_mark.Yes | question_mark.No)
            if choice_del == question_mark.Yes:
                if len(tech_name)==0 or len(clie_name)==0 or len(bran_name)==0:
                    self.label_215.setText("Please Select Technician Name, Client Name and branch Name!")
                else:
                    self.cur.execute('''
                        DELETE FROM monthly_followup WHERE tech_name=%s AND client_name=%s AND branch_name=%s
                    ''', (tech_name, clie_name, bran_name))
                    self.db.commit()
                    QMessageBox.information(self, 'success', 'Record Has Been Deleted!')
                    self.Clear_Search_Monthly_Todelete()

    def Show_Clname_Spare(self):
        self.clcombonamespare.clear()
        dash = ''
        self.clcombonamespare.addItem(dash)
        self.cur.execute('''
            SELECT client_name FROM new_client
        ''')
        client_name = self.cur.fetchall()
        for item in client_name:
            self.clcombonamespare.addItem(item[0])
            self.clcombonamespare.activated.connect(self.Show_Brname_Spare)
    
    def Show_Brname_Spare(self):
        cli_name = self.clcombonamespare.currentText()
        self.cur.execute('''
            SELECT branch_name FROM client_branch WHERE client_name=%s
        ''', (cli_name,))
        branches = self.cur.fetchall()
        self.branchcommaint.clear()
        dash = ''
        self.branchcommaint.addItem(dash)
        for branch in branches:
            self.branchcommaint.addItems(branch)
            self.branchcommaint.activated.connect(self.Branch_Joindate)

    def Branch_Joindate(self):
        branch_name = self.branchcommaint.currentText()
        if len(branch_name) == 0:
            self.label_179.setText('Please Select Branch!')
        else:
            self.label_179.clear()
            self.cur.execute('''
                SELECT join_date FROM client_branch WHERE branch_name=%s
            ''', (branch_name,))
            join_date = self.cur.fetchone()
            d2 = join_date[0]
            d1 = self.dateEdit_10.text()
            start_date = datetime.datetime.strptime(d2, '%d-%m-%Y')
            end_date = datetime.datetime.strptime(d1, '%d-%m-%Y')
            self.dateEdit_9.setDate(start_date)
            delta = relativedelta.relativedelta(end_date, start_date)
            self.lineEdit.setText(f"Open since {delta.years} years, {delta.months} months and {delta.days} days!")
            delta_month = delta.months+(delta.years*12)
            re_type = delta_month - delta_month%6
            if re_type < 6:
                self.lineEdit_2.setText(f"New Branch! open less than 6 months!")
                self.dateEdit_14.setDate(datetime.datetime.strptime(start_date, '%d-%m-%Y'))
            elif re_type == (delta.years*12):
                month_less = re_type - 6
                dt_month = self.dateEdit_9.date().addMonths(month_less)
                self.dateEdit_14.setDate(dt_month.toPyDate())
                last_m = self.dateEdit_14.text()
                self.lineEdit_2.setText(f"Last 6 months maintenance {month_less} months on: ({last_m})")
            else:
                dt_month = self.dateEdit_9.date().addMonths(re_type)
                self.dateEdit_14.setDate(dt_month.toPyDate())
                last_m = self.dateEdit_14.text()
                self.lineEdit_2.setText(f"Last 6 months maintenance {re_type} months on: ({last_m})")
            if delta.years < 1:
                self.lineEdit_3.setText(f"New Branch! open less than 12 months!")
                dt_join = self.dateEdit_9.text()
                self.dateEdit_8.setDate(datetime.datetime.strptime(dt_join, '%d-%m-%Y'))
            else:
                delta_year = delta.years*12
                dt_year = self.dateEdit_9.date().addMonths(delta_year)
                self.dateEdit_8.setDate(dt_year.toPyDate())
                last_y = self.dateEdit_8.text()
                self.lineEdit_3.setText(f"Last 12 months maintenance {delta_year} months on: ({last_y})")
            next_6months = 12
            nt_month = self.dateEdit_14.date().addMonths(next_6months)
            self.dateEdit_12.setDate(nt_month.toPyDate())
            next_m = self.dateEdit_12.text()
            if re_type == (delta.years*12):
                self.lineEdit_4.setText(f"Next 6 months maintenance {re_type+6} months on: ({next_m})")
            else:
                self.lineEdit_4.setText(f"Next 6 months maintenance {re_type+12} months on: ({next_m})")
            next_year = 12
            delta_year = delta.years*12
            num_year = delta_year + 12
            nt_year = self.dateEdit_8.date().addMonths(next_year)
            self.dateEdit_13.setDate(nt_year.toPyDate())
            nex_y = self.dateEdit_13.text()
            self.lineEdit_5.setText(f"Next 12 months maintenance {num_year} months on: ({nex_y})")

    def Show_Clname_Cleaners(self):
        self.clcombonamecleaner.clear()
        dash = ''
        self.clcombonamecleaner.addItem(dash)
        self.cur.execute('''
            SELECT client_name from new_client
        ''')
        client_name = self.cur.fetchall()
        for item in client_name:
            self.clcombonamecleaner.addItem(item[0])
            self.clcombonamecleaner.activated.connect(self.Show_Brname_Cleaners)

    def Show_Brname_Cleaners(self):
        cli_name = self.clcombonamecleaner.currentText()
        self.cur.execute('''
            SELECT branch_name FROM client_branch WHERE client_name=%s
        ''', (cli_name,))
        branches = self.cur.fetchall()
        self.branchclean.clear()
        dash = ''
        self.branchclean.addItem(dash)
        for branch in branches:
            self.branchclean.addItems(branch)
            self.branchclean.activated.connect(self.Branch_Cleaners)

    def Branch_Cleaners(self):
        branch_name = self.branchclean.currentText()
        if len(branch_name) == 0:
            self.label_180.setText('Please Select Branch!')
        else:
            self.label_180.clear()
            self.cur.execute('''
                SELECT join_date FROM client_branch WHERE branch_name=%s
            ''', (branch_name,))
            join_date = self.cur.fetchone()
            d2 = join_date[0]
            d1 = self.dateEdit_18.text()
            start_date = datetime.datetime.strptime(d2, '%d-%m-%Y')
            end_date = datetime.datetime.strptime(d1, '%d-%m-%Y')
            self.dateEdit_16.setDate(start_date)
            delta = relativedelta.relativedelta(end_date, start_date)
            self.lineEdit_11.setText(f"Open since {delta.years} years, {delta.months} months and {delta.days} days!")
            delta_month = delta.months + (delta.years*12)
            deliv_date = delta_month - delta_month%3
            if deliv_date < 3:
                self.dateEdit_17.setDate(datetime.datetime.strptime(d2, '%d-%m-%Y'))
                last_de = self.dateEdit_17.text()
                self.lineEdit_14.setText(f"New Branch, less than 3 months on: ({last_de})")
                add_deliv = 3
                new_month = self.dateEdit_16.date().addMonths(add_deliv)
                self.dateEdit_15.setDate(new_month.toPyDate())
                new_deliv = self.dateEdit_15.text()
                self.lineEdit_13.setText(f"Next deliver date on: ({new_deliv})")
            elif deliv_date > 3 :
                last_deliv = self.dateEdit_16.date().addMonths(deliv_date)
                self.dateEdit_17.setDate(last_deliv.toPyDate())
                con_date = self.dateEdit_17.text()
                self.lineEdit_14.setText(f"Last deliver on: ({con_date})")
                add_three = 3
                next_deliver = self.dateEdit_17.date().addMonths(add_three)
                self.dateEdit_15.setDate(next_deliver.toPyDate())
                next_date = self.dateEdit_15.text()
                self.lineEdit_13.setText(f"Next deliver on: ({next_date})")

    def Open_Welcome_Page(self):
        first_page = FirstScreen()
        widget.setFixedHeight(728)
        widget.setFixedWidth(1210)
        widget.addWidget(first_page)
        widget.setCurrentIndex(widget.currentIndex()+1)

    def Exit_Program_Func(self):
        self.db.commit()
        self.db.close()
        sys.exit()

app = QApplication(sys.argv)
welcome = FirstScreen()
widget = QtWidgets.QStackedWidget()
title = "Operation Department"
widget.setWindowTitle(title)
widget.addWidget(welcome)
widget.setFixedHeight(728)
widget.setFixedWidth(1210)
widget.show()

sys.exit(app.exec_())
