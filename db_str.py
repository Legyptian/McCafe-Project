from peewee import *
import datetime

db = MySQLDatabase('db-name', user='root', password='password', host='localhost', port=3306)

CALL_TYPE_CHOICE = (
    (1, 'Maintenance'),
    (2, 'New Installation'),
    (3, 'Uninstall_Equipment'),
    (4, 'Inspection & Equipment Investigation'))

MACHINE_TYPE = (
    (1, 'Coffe Machine'),
    (2, 'Grinder'),
    (3, 'Graneta Machine'),
    (4, 'Wash Machine'),
    (5, 'Water Filter')
)

TECHNICIAN_NAME = (
    (1, 'Technician One'),
    (2, 'Technician Two'),)

CLIENT_TYPE = (
    (1, 'New Client'),
    (2, 'Exist Client'),
)

class ClientType(Model):
    client_type = CharField(choices=CLIENT_TYPE)
    
    class Meta:
        database = db

class CallsInformation(Model):
    call_number = IntegerField(null=False, unique=True)
    call_type = CharField(choices=CALL_TYPE_CHOICE)
    call_by = CharField(null=False)
    mobile = IntegerField(null=False)
    client_name = CharField(null=False, unique=True)
    branch_name = CharField(null=False, unique=True)
    machine_type = CharField(choices=MACHINE_TYPE)
    client_code = IntegerField(null=False, unique=True)
    branch_code = IntegerField(null=False, unique=True)
    branch_address = CharField(null=True)
    user_name = CharField(null=False)
    client_complain = CharField(null=False)
    technician_name = CharField(choices=TECHNICIAN_NAME)
    recieve_date = DateTimeField(default=datetime.datetime.now)
    
    class Meta:
        database = db

class CallsHistory(Model):
    user_name = CharField(null=False)
    client_name = CharField(null=False, unique=True)
    branch_name = CharField(null=False, unique=True)
    machine_type = CharField(choices=MACHINE_TYPE)
    client_complain = CharField(null=False)
    call_by = CharField(null=False)
    mobile = IntegerField(null=False)
    technician_name = CharField(choices=TECHNICIAN_NAME)
    call_number = IntegerField(null=False, unique=True)
    recieve_date = DateTimeField(default=datetime.datetime.now)

    class Meta:
        database = db

USER_PROFESSIONAL=(
    (1, 'Engineer'),
    (3, 'Technician'),
    (4, 'Calls Manager'),
    (2, 'Employee'),)

class Users(Model):
    user_name = CharField(null=False, unique=True)
    user_code = IntegerField(null=False, unique=True)
    password = CharField(null=False, unique=True)
    user_professional = CharField(choices=USER_PROFESSIONAL)

    class Meta:
        database = db

GINDER=(
    (1, 'Male'),
    (2, 'Female'))

class UsersProfile(Model):
    user_name = CharField(null=False, unique=True)
    user_code = IntegerField(null=False, unique=True)
    user_professional = CharField(choices=USER_PROFESSIONAL)
    ginder = CharField(choices=GINDER)
    name = CharField(null=False)
    address = TextField(null=False)
    mobile = CharField(null=False)
    smobile = CharField(null=False)
    image = BlobField(null=False)
    
    class Meta:
        database = db

class Technician_Calls(Model):
    call_number = IntegerField(null=False, unique=True)
    user_name = CharField(null=False, unique=True)
    technician_name = CharField(choices=TECHNICIAN_NAME)
    client_name = CharField(null=False, unique=True)
    branch_name = CharField(null=False, unique=True)
    call_type = CharField(choices=CALL_TYPE_CHOICE)
    machine_type = CharField(choices=MACHINE_TYPE)
    recieve_date = DateTimeField(default=datetime.datetime.now)

    class Meta:
        database = db

class New_Client(Model):
    client_code = IntegerField(null=False, unique=True)
    client_name = CharField(null=False)
    main_address = CharField(null=True)
    join_date = DateTimeField(default=datetime.datetime.now)
    edit_date = DateTimeField(default=datetime.datetime.now)

    class Meta:
        database = db

class Client_Contact(Model):
    client_name =  CharField(null=False)
    contact_name = CharField(null=False)
    contact_code = IntegerField(null=False, unique=True)
    mobile = CharField(null=False)
    second_mobile = CharField(null=False)
    join_date = DateTimeField(default=datetime.datetime.now)
    edit_date = DateTimeField(default=datetime.datetime.now)

    class Meta:
        database = db

class Client_Branch(Model):
    client_name =  CharField(null=False)
    branch_name = CharField(null=False)
    branch_code = IntegerField(null=False, unique=True)
    branch_address = CharField(null=True)
    join_date = DateTimeField(default=datetime.datetime.now)
    edit_date = DateTimeField(default=datetime.datetime.now)

    class Meta:
        database = db

class Client_Machine(Model):
    client_name = CharField(null=False)
    branch_name =  CharField(null=False)
    machine_type = CharField(choices=MACHINE_TYPE)
    machine_model = CharField(null=True)
    machine_serial = CharField(null=True)
    machine_group = CharField(null=True)
    join_date = DateTimeField(default=datetime.datetime.now)
    edit_date = DateTimeField(default=datetime.datetime.now)

    class Meta:
        database = db

class McCafe_Branches(Model):
    branch_name = CharField(null=False)
    join_date = CharField(null=False)
    branch_code = IntegerField(null=False, unique=True)
    branch_address = CharField(null=False)
    machine_model = CharField(null=False)
    machine_serial = IntegerField(null=False, unique=True)
    number_of_groups = CharField(null=False)
    grinder_model = CharField(null=False)
    grinder_serial = CharField(null=False, unique=True)
    update_date = CharField(null=False)

    class Meta:
        database = db

class McCafe_Maintenance(Model):
    branch_name = CharField(null=False)
    join_date = CharField(null=False)
    last_maintenance = CharField(null=False)
    last_date = CharField(null=False)
    next_maintenance = CharField(null=False)
    next_6months = CharField(null=False)
    next_12months = CharField(null=False)

    class Meta:
        database = db

db.connect()
db.create_tables([ClientType, CallsInformation, CallsHistory, Users, UsersProfile, Technician_Calls, New_Client, Client_Contact, Client_Branch, Client_Machine, McCafe_Branches, McCafe_Maintenance])