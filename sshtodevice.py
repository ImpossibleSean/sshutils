import openpyxl
import sys
import pprint
from netmiko import SSHDetect
from netmiko import ConnectHandler
from netmiko.ssh_exception import NetMikoTimeoutException
from netmiko.ssh_exception import NetMikoAuthenticationException


def excelimportcol(filename, sheetname='Sheet1'):
    # Creates a key:value pair dictionary based on the key being the first item in the column
    # and adds all dictionaries to a list.
    workbook = openpyxl.load_workbook(filename)
    worksheet = workbook[sheetname]
    excellist = []
    for row in range(2, worksheet.max_row + 1):
        config = {}
        for col in range(1, worksheet.max_column + 1):
            key = worksheet.cell(row=1, column=col).value
            value = worksheet.cell(row=row, column=col).value
            config[str(key)] = str(value)
        excellist.append(config)
    return excellist

def autodetect(connect_kwargs):
    try:
        net_connect = SSHDetect(**connect_kwargs)
        print(net_connect.autodetect())

    except NetMikoTimeoutException:
        print('The IP address: {} did not respond, it might be unreachable'.format())
    except NetMikoAuthenticationException:
        print('Authentication failed for IP address: {}'.format(ipAddress))

def get_connection(connect_kwargs):
    try:
        net_connect = ConnectHandler(**connect_kwargs)
        print(net_connect.autodetect())

    except NetMikoTimeoutException:
        print('The IP address: {} did not respond, it might be unreachable'.format(ipAddress))
    except NetMikoAuthenticationException:
        print('Authentication failed for IP address: {}'.format(ipAddress))


def gathershowoutput(device, name):
    print('Connecting to ' + name)
    try:
        connect_kwargs = { 'device_type': device['device_type'],
                                'host': device['IP'],
                                'username': device['username'],
                                'password':device['password']}
        print(net_connect.autodetect())
        #net_connect = ConnectHandler(**device)
        output = net_connect.connection.send_command('vcgencmd measure_temp')
        print(output)
        # net_connect.find_prompt()

        # if "#" in net_connect.find_prompt():
        #     print('Connected to ' + net_connect.find_prompt().rstrip('#'))
        #     showIpIntBr = net_connect.send_command("show ip int br")
        #     print(showIpIntBr)
        #     outputs.setdefault(deviceName, showIpIntBr)
            # net_connect.disconnect()

    except NetMikoTimeoutException:
        print('The IP address: {} did not respond, it might be unreachable'.format(ipAddress))
    except NetMikoAuthenticationException:
        print('Authentication failed for IP address: {}'.format(ipAddress))


def removefalsecommits(devicelist):
    for devicedict in devicelist:
        if devicedict['Commit'] == 'True':
            next
        else:
            pprint.pprint('Removing element ')
            pprint.pprint(devicedict)
            devicelist.remove(devicedict)


def listdictappend(listofdicts, dict2):
    for dict1 in listofdicts:
        dict1.update( dict2 )

def readconfig(filename):
    f = open(filename, "r")
    for line in f.readlines():
        print(line)

# username = sys.argv[1]
# password = sys.argv[2]
username = 'nolooking'
password = 'definitelynotapassword'




devices = excelimportcol('workbook.xlsx')
print(devices)
removefalsecommits(devices)

deviceappend = {'device_type': 'autodetect', 'username': username, 'password': password}
listdictappend(devices, deviceappend)

# readconfig('config.txt')

print(devices)
#pprint.pprint(devices)
for devicedict in devices:
    gathershowoutput(devicedict, 'seanyboy')