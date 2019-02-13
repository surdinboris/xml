#old xml library
#import xml.etree.ElementTree as ET
from lxml import etree as ET
import gzip
import os
import os.path
import glob
#import sys, getopt
import subprocess
import shutil
import re
import time
import tkinter as tk
import tkinter.scrolledtext as tkst
from tkinter import *
from PIL import ImageTk, Image
from IPy import IP
# import smtplib
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# import datetime
import xlsxwriter
import telnetlib
import nmap
#generator for AB style excell cells
def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string
running = True
hardware_golden = 'HardwareInventory.golden'
configuration_golden = 'ConfigurationInventory.golden'
#pdus/sensors lists
sensors = {'sensor1': 'Front-Down', 'sensor2': 'Front-Up', 'sensor3': 'Rear-Down', 'sensor4': 'Rear-Up'}
pdus=['10.48.228.51', '10.48.228.52', '10.48.228.53', '10.48.228.54']

#servers count for power test
servers_count=26
#additional attributes to collect for dynamic configuration data (FQDD, <!-- <Attribute Name=" ....)
additional_conf_collect = {}
additional_conf_collect.update({"Disk.Virtual.0:RAID.Integrated.1-1": ['Name', 'Size', 'StripeSize', 'SpanDepth', 'SpanLength', 'RAIDTypes', 'IncludedPhysicalDiskID']})
#additional_conf_collect.update({"iDRAC.Embedded.1": ["IPv4Static.1#Address"]})
# summary object init
summary = {}
errors={}
failoverresult= {'PDU-{}'.format(num): {} for num, pdu in enumerate(pdus, 1)}
#harware collection constructor
hw_collect=[]
hw_collect.append({'displayname': 'ServiceTag', 'classname': 'DCIM_SystemView', 'name': 'ServiceTag', 'excluded_for_validation': 2})
hw_collect.append({'displayname': 'HostName', 'classname': 'DCIM_SystemView', 'name': 'HostName', 'excluded_for_validation': 2})
hw_collect.append({'displayname': 'Inventory date', 'classname': 'DCIM_SystemView', 'name': 'LastSystemInventoryTime', 'excluded_for_validation': 2})
hw_collect.append({'displayname': 'CPU model', 'classname': 'DCIM_CPUView', 'name': 'Model', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'PCI device', 'classname': 'DCIM_PCIDeviceView', 'name': 'Description', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'Memory tot.size', 'classname': 'DCIM_SystemView', 'name': 'SysMemTotalSize', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'Memory serial', 'classname': 'DCIM_MemoryView', 'name': 'SerialNumber', 'excluded_for_validation': 2})
hw_collect.append({'displayname': 'Memory P/Ns', 'classname': 'DCIM_MemoryView', 'name': 'PartNumber', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'Memory slot', 'classname': 'DCIM_MemoryView', 'name': 'FQDD', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'HDD serial', 'classname': 'DCIM_PhysicalDiskView', 'name': 'SerialNumber', 'excluded_for_validation': 2})
hw_collect.append({'displayname': 'HDD model', 'classname': 'DCIM_PhysicalDiskView', 'name': 'Model', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'HDD fw', 'classname': 'DCIM_PhysicalDiskView', 'name': 'Revision', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'HDD slot pop.', 'classname': 'DCIM_PhysicalDiskView', 'name': 'Slot', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'PSU P/Ns', 'classname': 'DCIM_PowerSupplyView', 'name': 'PartNumber', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'PSU serial', 'classname': 'DCIM_PowerSupplyView', 'name': 'SerialNumber', 'excluded_for_validation': 2})
hw_collect.append({'displayname': 'PSU fw', 'classname': 'DCIM_PowerSupplyView', 'name': 'FirmwareVersion', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'PSU status', 'classname': 'DCIM_PowerSupplyView', 'name': 'PrimaryStatus', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'NIC status', 'classname': 'DCIM_NICView', 'name': 'LinkSpeed', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'HealthStatus', 'classname': 'DCIM_SystemView', 'name': 'PrimaryStatus', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'PowerState', 'classname': 'DCIM_SystemView', 'name': 'PowerState', 'excluded_for_validation': 0})
#dynamic_collect.update({"Disk.Bay.6:Enclosure.Internal.0-1:RAID.Integrated.1-1": ["RAIDHotSpareStatus"]})

#getroot helper for use directly from report (for configuration pasing)
#  and in getdata requests (for hw inventory parsing)
def getroot(xml):
    with open(xml, 'r') as x:
        data = x.read()
    #old library
    #root = ET.fromstring(data)
    root = ET.fromstring(data)
    return root


# adding RAID data (from dynamic -commented- part
def add_dynamic_attrs(FQDD, collect, xml):
    result={}
    tree = getroot(xml)
    for compon in tree.iter():
        if compon.get('FQDD') == FQDD:
            for ref in compon.getchildren():
                # print('par name', ref.items(), ref.get('Name'), ref.getparent().get('FQDD'))
                if ref.get('Name') == None:
                    # print('-' * 40)
                    ref = str(ref)
                    strref = ref.strip().replace('<!--', '').replace('-->', '').replace('ReadOnly', '')
                    prop = ET.fromstring(strref)
                    val = prop.text
                    key = prop.attrib['Name']
                    if key in collect:
                        result.update({key: val})
    return result

def getdata(xml,classname='', name=''):
    root = getroot(xml)
    #hwinventory collect helper
    def collect(inst, classnameattr):
        hwinventory = {}
        for i in inst:
            # gathering results. example: Component Classname="DCIM_ControllerView
            if i.attrib[classnameattr] == classname:
                props = i.findall('PROPERTY')
                fqdd= ''
                val=''
                for prop in props:
                    if prop.attrib['NAME'] == name:
                        val = prop.find('VALUE').text
                    elif prop.attrib['NAME'] == "FQDD":
                        fqdd = prop.find('VALUE').text

                hwinventory[fqdd] = val

        hwinventory=dict(sorted(hwinventory.items()))
        if len(hwinventory) == 0:
            return ['n/a']
        return list(hwinventory.values())
    #router to use both two types of hwinventory retrieved via web interface or
    #racadmin and additional support for segregate requests configuration parsing (possibly not needed)

    #collecting hwinventory items in case of hwinventory detected
    if root.tag =='Inventory':
        inst = root.findall('Component')
        classnameattr = 'Classname'
        return collect(inst, classnameattr)

    elif root.tag == 'CIM':
        inst = root.findall('MESSAGE/SIMPLEREQ/VALUE.NAMEDINSTANCE/INSTANCE')
        classnameattr = 'CLASSNAME'
        return collect(inst, classnameattr)

    #collecting hwinventory items in case of configuration parsing detected
    #and building custom structure attribute-value pairs
    elif root.tag =='SystemConfiguration':
        confinventory=[]
        #Additionally parsing for commented raidconf attrs
        raidconf= 'RAID conf n/a'

        inst = root.findall('Component')
        for i in inst:
            FQDD = i.attrib['FQDD']
            # gathering results examle: FQDD="LifecycleController.Embedded.1
            props = i.findall('Attribute')
            for prop in props:
                val = prop.text
                key = FQDD+' '+prop.attrib['Name']
                confinventory.append({key: val})
        for FQDD in additional_conf_collect:
            confinventory.append(add_dynamic_attrs(FQDD, additional_conf_collect[FQDD],xml))
        return confinventory


def main(argv):
    #helpers
    #gui operations
    def print_to_gui(txtstr):
        _texbox.config(state='normal')
        _texbox.insert('end', '%s\n' %txtstr)
        _texbox.config(state="disabled")
        _texbox.see('end')
        _root.update()
    def texboxclear():
        _texbox.config(state='normal')
        _texbox.delete('1.0', END)
        _texbox.config(state="disabled")
        _root.update()
    #network scan
    def nmapscan():
        nm = nmap.PortScanner()
        nm.scan('10.48.228.1-40', '22')
        """Sort an IP address list."""
        ipl = [(IP(ip).int(), ip) for ip in nm.all_hosts()]
        ipl.sort()
        ipl= list(map(lambda i: i[1], ipl))

        print("Found hosts:")
        for host in ipl:
            print('.' * 40)
            print('Host : {}'.format(host))
        print('.' * 40)
        return ipl
    #pdu operations
    def sendcom(tel, cmd=b""):
        tel.write(cmd)
        tel.write(bytes("\r", encoding='ascii'))

    def login(tel):
        user = 'root'
        password = 'wildcat1'
        tel.read_until(b"Username:")
        sendcom(tel, user.encode())
        tel.read_until(b"Password:")
        sendcom(tel, password.encode())
        tel.read_until(b"#")

    def pdu_command(host, cmd):
        tel = telnetlib.Telnet(host)
        login(tel)
        #print(cmd.encode())
        sendcom(tel,cmd.encode())
        tel.read_until(b"[y/n]")
        #print('y/n')
        sendcom(tel,'y'.encode())
        tel.read_until(b"#")
        sendcom(tel, b'exit')

    def sensorscheck(pdu):
        with telnetlib.Telnet(pdu) as t:
            t.read_until(b'Username:')
            t.write(b'root\n')
            t.read_until(b'Password:')
            t.write(b'wildcat1\n')
            t.read_until(b"#")
            t.write(b'show sensor externalsensor 2\n')
            temp_output = t.read_until(b"#").decode('utf-8')
            t.write(b'show inlets details\n')
            pow_output = t.read_until(b"#").decode('utf-8')
            temp = re.search('Reading:\s+(\d+\.\d)', temp_output).group(1)
            wattage = re.search('Active Power:\s+(\d+)', pow_output).group(1)
            return {'wattage': wattage, 'temp': temp}


    def failover_check():
        print_to_gui('Starting pdu failover check')
        #building data structure
        #{'PDU-1':{'result':'pass','sensor1':{'wattage':'100','temp':'27'},'sensor2':{'wattage':'100','temp':'27'}}, 'PDU-2':...}
        #executing
        for num, pdu in enumerate(pdus, 1):
            print('System powered without PDU-{}'.format(num))
            print_to_gui('turning off PDU-{}'.format(num))
            pdu_command(pdu, 'power outlets all off')
            time.sleep(5)
            if len(nmapscan()) == servers_count:
                print_to_gui('All {} system servers are online'.format(len(nmapscan())))
                pdu_command(pdu, 'power outlets all on')
                time.sleep(12)
                for wnum, wpdu in enumerate(pdus, 1):
                    #print_to_gui('Checking wattage and temperature for sensor{}'.format(wnum))
                    #creating sensor record
                    sensor = failoverresult['PDU-{}'.format(num)]['sensor{}'.format(wnum)] = {}
                    sensorsdata = sensorscheck(wpdu)
                    sensor['wattage'] = sensorsdata['wattage']
                    sensor['temp'] = sensorsdata['temp']
                failoverresult['PDU-{}'.format(num)]['result'] = 'pass'
            else:
                print_to_gui('Error: found {} active servers, while should be {}'.format(len(nmapscan()), servers_count))
                pdu_command(pdu, 'power outlets all on')
                failoverresult['PDU-{}'.format(num)]['result']='fail'
                pdu_command(pdu, 'power outlets all on')
                time.sleep(12)

    def disbutt(opt):
        for bu in buttons:
            bu['state'] = opt

    _root=Tk()
    retrieveinitial = IntVar(value=0)
    applygolden = IntVar(value=0)
    collectfinal = IntVar(value=1)
    checkfailover = IntVar(value=1)
    spec_ip = StringVar()
    spec_pwd = StringVar()
    #telad_logo = ImageTk.PhotoImage(Image.open(os.path.join(os.getcwd(), "logo.gif")))
    liveperson_logo = ImageTk.PhotoImage(Image.open(os.path.join(os.getcwd(), "liveperson.gif")))
    _root.title('Liveperson dell inventory tool')
    _root.resizable(width=False, height=False)
    _mainframe = tk.Frame(_root)
    _mainframe.grid(row=0, column=0, sticky=(E, W, N, S))
    #_telad_logo = Label(_mainframe, image=telad_logo)
    #_telad_logo.grid(row=0, padx=5, pady=5, column=0, sticky=(W, N))
    _liveperson_logo = Label(_mainframe, image=liveperson_logo)
    _liveperson_logo.grid(row=0, padx=5, pady=5, column=0, sticky=(W, N))
    # output part
    _textboxframe = tk.LabelFrame(_mainframe, text='Work log')
    _textboxframe.grid(row=0, padx=5, pady=5, column=1, rowspan=3, sticky=(W, N))
    _textboxframe.columnconfigure(0, weight=1)
    _textboxframe.rowconfigure(0, weight=1)
    _texbox = tkst.ScrolledText(_textboxframe, wrap='word', width=35, height=20, state='disabled')
    _texbox.grid(row=0, column=1, sticky=(W, N))
    #options

    _optionsframe = tk.LabelFrame(_mainframe, text='Network run options')
    _optionsframe.grid(row=1, padx=3, pady=3, column=0, sticky=(W, N))

    _ipText = Label(_optionsframe, text='IP:')
    _ipText.grid(row=0, padx=3, pady=3, column=0, sticky=(E, N))

    _ipaddress = Entry(_optionsframe, textvariable= spec_ip)
    _ipaddress.grid(row=0, padx=3, pady=3, column=1, sticky=(E, N))

    _pwdText = Label(_optionsframe, text='Pass:')
    _pwdText.grid(row=1, padx=3, pady=3, column=0, sticky=(E, N))

    _password = Entry(_optionsframe, textvariable=spec_pwd)
    _password.grid(row=1, padx=3, pady=3, column=1, sticky=(E, N))

    _retrieveinitial = Checkbutton(_optionsframe, text="Initial inventory", variable=retrieveinitial)
    _retrieveinitial.grid(row=2, padx=3, pady=3, column=0, columnspan=2, sticky=(W, N))

    _applygolden = Checkbutton(_optionsframe, text="Apply golden settings", variable=applygolden)
    _applygolden.grid(row=3, padx=3, pady=3, column=0, columnspan=2, sticky=(W, N))

    _collectfinal= Checkbutton(_optionsframe, text="Collect final report", variable=collectfinal)
    _collectfinal.grid(row=4, padx=3, pady=3, column=0, columnspan=2, sticky=(W, N))

    _collectfinal= Checkbutton(_optionsframe, text="Check PDU failover", variable=checkfailover)
    _collectfinal.grid(row=5, padx=3, pady=3, column=0, columnspan=2, sticky=(W, N))
    # testing part
    _testingframe = tk.LabelFrame(_mainframe, text='Testing')
    _testingframe.grid(row=2, padx=3, pady=3, column=0,  sticky=(W,  N))

    # _testingframe.columnconfigure(1, weight=10)
    # _testingframe.rowconfigure(1, weight=10)
    #control
    # test buttons - start stop test
    _startnetbutton = tk.Button(_testingframe, text='Start (network)',width=20, height=2,
                                command = lambda: start('network'))
    _startnetbutton.grid(row=0, padx=3, pady=3, column=0, sticky=(W, N))

    _startofflinebutton = tk.Button(_testingframe, text='Start (offline)', width=20, height=2,
                                    command = lambda: start('offline'))
    _startofflinebutton.grid(row=1, padx=3, pady=3, column=0, sticky=(W, N))

    #to be implemented
    # _stopbutton = tk.Button(_testingframe, text='Stop execution', width=20, height=2,
    #                                 command=lambda: start('stop'))
    # _stopbutton.grid(row=2, padx=3, pady=3, column=0, sticky=(W, N))

    buttons = [_startnetbutton,_startofflinebutton]
    def start(mode):
        global errors
        global summary
        global failoverresult
        disbutt('disabled')
        print_to_gui('')
        print_to_gui('Test started in {} mode '.format(mode))
        # fallbacks - to current workdir
        temp = os.path.join(os.getcwd(), 'temp')
        arrived = os.path.join(os.getcwd(), 'arrived')
        xmldatadir=os.path.join(os.getcwd(), 'passed','testresults_'+time.strftime("%d%m%Y_%H-%M-%S", time.gmtime(time.time())))
        os.mkdir(xmldatadir)
        #xmldatadir  = os.path.join(repdirname,'summary_'+time.strftime("%d%m%Y_%H-%M-%S", time.gmtime(time.time())))
        def cleantemp(temp):
            for inputfile in os.listdir(temp):
                print('clearing', os.path.join(temp, inputfile))
                os.remove(os.path.join(temp, inputfile))
            if len(os.listdir(temp)) != 0:
                raise FileExistsError('Clearing of temporary dir failed, please check!')
        repname=os.path.join(os.getcwd(),'summary_'+time.strftime("%d%m%Y_%H-%M-%S", time.gmtime(time.time()))+'.xlsx')
        workbook = xlsxwriter.Workbook(repname)
        #prepairing summary report first
        summary_report=workbook.add_worksheet('summary_report')
        if mode == 'network':
            #########Network run
            # retrieving hosts information
            if len(spec_ip.get()) > 0:
                print_to_gui('Predefined ip {}'.format(spec_ip.get()))
                active_hosts = [spec_ip.get()]
            else:
                active_hosts = nmapscan()

            print_to_gui('Found {} active hosts'.format(len(active_hosts)))
            #cli part
            # answer = input("Found {} hosts. Do you want to proceed?[y/n]".format(len(active_hosts)))
            # if not answer or answer[0].lower() != 'y':
            #     print('Interrupting')
            #     exit(1)

            for host in active_hosts:
                print('\n' * 2)
                print('-_' * 30)
                print_to_gui("Connecting to host {}".format(host))
                print("Connecting to host {}".format(host))
                cleantemp(temp)
                ####first part - disabled performed via operator's script
                password = "wildcat1"
                if len(spec_pwd.get()) > 0:
                    password = spec_pwd.get()

                if retrieveinitial.get() == 1:
                    #changing default password tio calvin to collect arrived inv
                    password= "calvin"
                    if len(spec_pwd.get()) > 0:
                        password = spec_pwd.get()

                    print_to_gui('- Collect arrived inventory')
                    #get orig data via racadm - disabled implemented at the earlier stage:
                    #os.system("racadm -r {host} -u root -p calvin hwinventory export -f {fn}".format(host,os.path.join(temp,"hw_orig_tmp.xml")))

                    subprocess.run(["racadm", "-r", host, "-u", "root", "-p", password, "hwinventory", "export", "-f",
                                    "{}".format(os.path.join(temp,"hw_orig_tmp.xml"))])
                    print_to_gui('- Collect arrived configuration')
                    subprocess.run(["racadm", "-r", host, "-u", "root", "-p", password, "--nocertwarn", "get", "-t", "xml", "-f",
                                    "{}".format(os.path.join(temp,"conf_orig.tmp.xml"))])
                    files_processing(temp, arrived,workbook, step='arrived')
                    cleantemp(temp)

                if applygolden.get() == 1:
                    print_to_gui('- Applying Golden configuration')
                    #applying golden template
                    print("Applying Golden configuration, please wait....")
                    subprocess.run(["racadm", "-r", host, "-u", "root", "-p", password, "--nocertwarn", "set", "-f",
                                    "{}".format(os.path.join(os.getcwd(), "ConfigurationInventory.golden")), "-t", "xml", "-b",
                                    "graceful", "-w", "600", "-s", "on"])
                if collectfinal.get() == 1:

                    print_to_gui('- Collect final inventory')
                    # getting data after golden termplate enrollment:
                    subprocess.run(["racadm", "-r", host, "-u", "root", "-p", password, "hwinventory", "export", "-f",
                                    "{}".format(os.path.join(temp, "hw_passed.xml"))])
                    print_to_gui('- Collect final configuration')
                    subprocess.run(
                        ["racadm", "-r", host, "-u", "root", "-p", password, "--nocertwarn", "get", "-t", "xml", "-f",
                         "{}".format(os.path.join(temp, "conf_passed.xml"))])

                # ##{'Health': 'OK', 'PowerState': 'Off'} or {'Health': None,'PowerState': None}
                # hwinfo = subprocess.run(
                #     ["python3.6", "GetSystemHWInventoryREDFISH.py", "-ip", host, "-u", "root", "-p", password, "-s", "y"],
                #     stdout=subprocess.PIPE)
                # hwinfo = hwinfo.stdout.decode().split("\n")
                # server_status = {'Health': None, 'PowerState': None}
                # for h in hwinfo:
                #     health = re.search("Status: {'Health': '(\w+)'.*}", h)
                #     if health:
                #         server_status.update({'Health': health[1]})
                #     power_on = re.search("PowerState: (\w+)", h)
                #     if power_on:
                #         server_status.update({'PowerState': power_on[1]})

                # verifying against golden template
                files_processing(temp, xmldatadir, workbook, step='golden', ip=host)
                cleantemp(temp)
            if checkfailover.get() == 1:
                failover_check()
            writesummary(workbook, summary_report)
            #reinit of data placeholders
            summary = {}
            failoverresult = {'PDU-{}'.format(num): {} for num, pdu in enumerate(pdus, 1)}
            print_to_gui(' - Process finished. Please inspect {}'.format(repname))


        elif mode == 'offline':
            # offline run
            print_to_gui('Processing files in {}...'.format(os.path.abspath(os.getcwd())))
            #server_status={'Health': 'N/A', 'PowerState': 'N/A'}
            repsdir=os.path.join(os.getcwd(), "offline")
            files_processing(repsdir, repsdir, workbook, ip= '0.0.0.0')
            if checkfailover.get() == 1:
                failover_check()
            writesummary(workbook, summary_report)
            #reinit of data placeholders
            summary = {}
            failoverresult = {'PDU-{}'.format(num): {} for num, pdu in enumerate(pdus, 1)}
            print_to_gui(' - Process finished. Please inspect {}'.format(repname))
        workbook.close()
        if len(errors) > 0:
            print_to_gui('Following errors were detected:')
            for er in errors:
                print_to_gui("{}: {}".format(er, errors[er]))
        errors={}
        disbutt('normal')
    _root.mainloop()

def files_processing(inputdir, outputdir, workbook, step=None, ip=None):

    counter = 0
    for inputfile in os.listdir(inputdir):
        fn, ext = (os.path.splitext(inputfile))
        if ext == '.xml':
            #in case of arrived server checking - parsing xml and returning xml data
            if step == 'arrived':
                print('Found  xml file for arrived server: {} Processing...'.format(fn + ext))
                # report generation for (naming purposes only)
                cur_report = report(os.path.join(inputdir, inputfile))
                service_tag = cur_report['service_tag']
                rep_type = cur_report['rep_type']
                filename = os.path.join(outputdir, "{}_{}_{}".format(service_tag, rep_type, fn+ext))
                shutil.copyfile(os.path.join(inputdir,inputfile ), os.path.join(outputdir,filename))
                print('Arrived report for ST{} stored in {}'.format(service_tag, filename))
                counter += 1

            elif step == 'golden':
                #report_file_name = os.path.join(outputdir, os.path.join(inputdir, inputfile)) + '_report.xlsx'
                print('Found xml file for golden comparison: {} Processing...'.format(fn + ext))
                # report generation
                cur_report = report(os.path.join(inputdir, inputfile))
                service_tag = cur_report['service_tag']
                rep_type = cur_report['rep_type']
                filename = os.path.join(outputdir, "{}_{}_{}".format(service_tag, rep_type, fn + ext))
                shutil.copyfile(os.path.join(inputdir, inputfile), os.path.join(outputdir, filename))
                # report analysing
                cur_report = report_analyze(cur_report)
                # summary entry appending
                try:
                    summary[service_tag]
                except KeyError:
                    summary[service_tag] = []
                summary[service_tag].append(cur_report)
                summary[service_tag].append({'ip': ip})
                # writetoxlsx(os.path.join(outputdir, "{}_{}_{}".format(service_tag, rep_type, fn+'_report.xlsx',workbook)), cur_report)
                writetoxlsx("{}_{}".format(service_tag, rep_type), cur_report, workbook)
                counter += 1
                print('Passed report for {} stored in {}'.format(service_tag, filename))

            #default behavior - for testing only
            else:
                #report_file_name = os.path.join(outputdir, os.path.join(inputdir,inputfile)) + '_report.xlsx'
                print('Found xml file: {} Processing...'.format(fn+ext))
                #report generation
                cur_report = report(os.path.join(inputdir, inputfile))
                service_tag = cur_report['service_tag']
                #report analysing
                cur_report = report_analyze(cur_report)
                rep_type = cur_report['rep_type']
                try:
                    summary[service_tag]
                except KeyError:
                    summary[service_tag] = []
                summary[service_tag].append(cur_report)
                summary[service_tag].append({'ip': ip})
                #writetoxlsx(report_file_name, cur_report, workbook)
                writetoxlsx("{}_{}".format(service_tag, rep_type), cur_report, workbook)
                counter += 1
                #last execution block
                print('{} done. Processed {}, files'.format(service_tag, counter))

def report_analyze(currep):
    result = {}
    rep_type = currep['rep_type']
    #building up data structure as following:
    # {'Memory slot': [{'DIMM.Socket.A1': 1}, {'DIMM.Socket.A2': 1}, {'DIMM.Socket.A3': 1}, {'DIMM.Socket.A4': 1},
    #                  {'DIMM.Socket.B1': 1}, {'DIMM.Socket.B2': 1}, {'DIMM.Socket.B3': 1}, {'DIMM.Socket.B4': 1}],
    #  'PSU model': [{'PWR SPLY,750W,RDNT,DELTA      ': 1}, {'PWR SPLY,750W,RDNT,DELTA      ': 1}]}

    # routing for hwinventory  or configuration
    if rep_type =='hwinvent_report':
        master = report(os.path.join(os.getcwd(), hardware_golden))['report']
        print('Master report generated from {} \n'.format(os.path.join(os.getcwd(), hardware_golden)))
    elif rep_type == 'config_report':
        master = report(os.path.join(os.getcwd(), configuration_golden))['report']
        print('Master report generated from {} \n'.format(os.path.join(os.getcwd(), configuration_golden)))

    #extracting report
    currep = currep['report']
    for record in currep:

        #in case of record availalable in master file
        data_per = []
        if currep[record]['valid'] == 2:
            for data_item in currep[record]['data']:
                data_per.append({data_item:2,'golden': 'dynamic field'})
            result[record] = data_per
            continue
        try:
            master_record = master[record]
            #print(record,'>>>>>>>m', master_record, '>>>>>>>c', currep[record])

            data_per = []
            for i, data_item in enumerate(currep[record]['data']):
                try:
                    master_val = master[record]['data'][i]
                except IndexError:
                    master_val = 'not present in golden configuration'
                data_per.append({data_item: int(data_item == master_val), 'golden': master_val})
            result[record] = data_per
               #print('unequal', master_record['data'], current[record]['data'],'\n')
                #old result[record] = {'data': current[record]['data'], 'valid': 0}
        except KeyError:
            #if failed to find whole master values branch in master file - assign specific attribute
            data_per = []
            for data_item in currep[record]['data']:
                data_per.append({data_item: 5, 'golden': 'n/a'})
            result[record] = data_per
                #continue
            #result[record] = {'data':record['data'], 'valid': 5}
            #print(master_record)
    return {'rep_type': rep_type, 'report': result}


def unpack(latest_file):
    epath, tail =os.path.split(latest_file)
    for gzip_path in glob.glob(epath + "/*.gz"):
        if os.path.isdir(gzip_path) == False:
            inF = gzip.open(gzip_path, 'rb')
            # uncompress the gzip_path INTO THE 's' variable
            s = inF.read()
            inF.close()
            # get gzip filename (without directories)
            gzip_fname = os.path.basename(gzip_path)
            # get original filename (remove 3 characters from the end: ".gz")
            fname = gzip_fname[:-3]
            uncompressed_path = os.path.join(epath, fname)
            # store uncompressed file data from 's' variable
            open(uncompressed_path, 'wb').write(s)

        for f in os.listdir(epath):
            latest_file_spl=os.path.splitext(os.path.basename(latest_file))[0]
            if f == latest_file_spl:
                #fn, ext = (os.path.splitext(f))
                if os.path.splitext(f)[1] == '.xml':
                    return(os.path.join(epath,f))

def writesummary(workbook,worksheet):

    maxwidth = {}
    # creating xls file
    #workbook = xlsxwriter.Workbook(report_file_name)
    # header
    header_cell = workbook.add_format()
    header_cell.set_bold()
    # green cell - passed validation against master file
    green_cell = workbook.add_format()
    green_cell.set_font_color('green')
    # red cell - failed validation against master file
    red_cell = workbook.add_format()
    red_cell.set_font_color('red')
    # black_cell - dynamic data such as SN that non need to be validated
    # ( added 'excluded_for_validation': 1 to results in report constructor)
    black_cell = workbook.add_format()
    black_cell.set_font_color('gray')
    # yellow cell in case of result is not found in master file
    orange_cell = workbook.add_format()
    orange_cell.set_font_color('orange')
    # create worksheet
    #worksheet = workbook.add_worksheet('summary_report')

    # helper to calculate and update width for column
    def toStr(val, coord):
        if val == None:
            val = ''
        try:
            curr = maxwidth[coord[0]]
            if curr < len(str(val)):
                maxwidth[coord[0]] = len(str(val))
        except KeyError:
            maxwidth[coord[0]] = len(str(val))
        return str(val)

    #print(summary)

    maxheight = 2
    for result in summary:
        print('Summary detected for {}'.format(result))
        ServiceTag = result
        conf_passed = 1
        conf_error = []
        hw_error = []
        currheight = maxheight
        # entering to report data
        reps={}

        for res in summary[result]:
            try:
                rep_type= res['rep_type']
                if rep_type == 'config_report':
                    reps.update({rep_type:res['report']})
                if rep_type == 'hwinvent_report':
                    reps.update({rep_type: res['report']})
            except KeyError:
                ip = res['ip']
                # server_status = res['server_status']

    # #if rep_type == 'config_report':
        try:
            reps['config_report']
        except KeyError:
            print('Error: No configuration found for {}!!!'.format(ServiceTag))
            errors[ServiceTag] = 'Error: No configuration found for {}!!!'.format(ServiceTag)
            continue

        for ind, confsingle in enumerate(reps['config_report'], 0):
        #coords = '{}{}'.format(colnum_string(i), ind)
            #print(ind, v, )
            conf_items=reps['config_report'][confsingle]
            for confitem in conf_items:
                for key, value in confitem.items():
                    if key != 'golden':
                        if value == 1:
                            pass
                            # conf_passed.append('')
                        elif value == 0:
                            #print('conf error',confitem, key,value)
                            conf_passed=0
                            conf_error.append('Wrong value: of {},got {}  should be  {}'.format(confsingle, key, confitem['golden']))
                        elif value == 2:
                            pass


    #elif rep_type == 'hwinvent_report':
        try:
            reps['hwinvent_report']
        except KeyError:
            print('Error: No hardware inventory found for {}!!!'.format(ServiceTag))
            errors[ServiceTag] = 'Error: No hardware inventory found for {}!!!'.format(ServiceTag)
            continue

        correction=0
        for ind, hwfamily in enumerate(reps['hwinvent_report'],1):
            ind = ind+correction
            #writing head
            hw_items=reps['hwinvent_report'][hwfamily]
            hwfamily_pass = 1
            for i, hwitem in enumerate(hw_items,maxheight):
                corrflag = False
                for key, value in hwitem.items():
                    if key != 'golden':
                        if value == 1:
                            # writing head
                            if maxheight == 2:
                                coords = '{}1'.format(colnum_string(ind))
                                worksheet.write(coords, toStr(hwfamily, coords), header_cell)
                        elif value == 0:
                            hwfamily_pass=0
                            if maxheight == 2:
                                coords = '{}1'.format(colnum_string(ind))
                                worksheet.write(coords, toStr(hwfamily, coords), header_cell)

                            #print(ind,v ,hw_items)
                            hw_error.append('Wrong value: of {},got {}  should be  {} '.format(hwfamily, key, hwitem['golden']))
                        elif value == 2:
                            hwfamily_pass = 2
                            if hwfamily == "ServiceTag":
                                # writing head
                                if maxheight == 2:
                                    # writing value
                                    coords = '{}1'.format(colnum_string(ind))
                                    worksheet.write(coords, toStr(hwfamily, coords), header_cell)
                                coords = '{}{}'.format(colnum_string(ind),maxheight)
                                worksheet.write(coords, toStr(key, coords), black_cell)

                                #appending ip - making one more column
                                # worksheet.write_comment(coords, ip)
                                ind = ind+1
                                correction = correction+1
                                if maxheight == 2:
                                    # writing value
                                    coords = '{}1'.format(colnum_string(ind))
                                    worksheet.write(coords, toStr('System IP', coords), header_cell)
                                coords = '{}{}'.format(colnum_string(ind), maxheight)
                                worksheet.write(coords, toStr(ip,coords),black_cell)

                                #making space for dynamic attr insertion
                                corrflag = False
                            elif hwfamily == "HostName":
                                # writing head
                                if maxheight == 2:
                                    # writing value
                                    coords = '{}1'.format(colnum_string(ind))
                                    worksheet.write(coords, toStr(hwfamily, coords), header_cell)
                                coords = '{}{}'.format(colnum_string(ind),maxheight)
                                worksheet.write(coords, toStr(key, coords), black_cell)
                                #making space for dynamic attr insertion
                                corrflag=False
                            else:
                                corrflag=True
            if corrflag:
                correction = correction - 1
            if hwfamily_pass == 1:
                coords = '{}{}'.format(colnum_string(ind),maxheight)
                worksheet.write(coords, toStr('pass', coords), green_cell)
            if hwfamily_pass == 0:
                coords = '{}{}'.format(colnum_string(ind),maxheight)
                worksheet.write(coords, toStr('fail', coords), red_cell)
                worksheet.write_comment(coords, str(hw_error))
                hw_error=[]

        #manual index correction before configuration appending
        ind = ind + 1
        if conf_passed == 1:
            if maxheight == 2:
                coords = '{}1'.format(colnum_string(ind))
                worksheet.write(coords, toStr('Configuration', coords), header_cell)
            coords = '{}{}'.format(colnum_string(ind), maxheight)
            worksheet.write(coords, toStr('conf. pass', coords), green_cell)

        if conf_passed == 0:
            if maxheight == 2:
                coords = '{}1'.format(colnum_string(ind))
                worksheet.write(coords, toStr('Configuration', coords), header_cell)
            coords = '{}{}'.format(colnum_string(ind), maxheight)
            worksheet.write(coords, toStr('conf. fail', coords), red_cell)
            worksheet.write_comment(coords, str(conf_error))
        #writing link to Hardware report
        ind = ind + 1
        if maxheight == 2:
            coords = '{}1'.format(colnum_string(ind))
            worksheet.write(coords, toStr('Hardware report', coords), header_cell)
        coords = '{}{}'.format(colnum_string(ind), maxheight)
        worksheet.write(coords, 'internal:\'{}_hwinvent_report\'!A1'.format(ServiceTag))
        #writing link to Config report
        ind = ind + 1
        if maxheight == 2:
            coords = '{}1'.format(colnum_string(ind))
            worksheet.write(coords, toStr('Config report', coords), header_cell)
        coords = '{}{}'.format(colnum_string(ind), maxheight)
        worksheet.write(coords, 'internal:\'{}_config_report\'!A1'.format(ServiceTag))
        #appending server_status from Redfish
        # ind = ind + 1
        # if maxheight == 2:
        #     coords = '{}1'.format(colnum_string(ind))
        #     worksheet.write(coords, toStr('Health status', coords), header_cell)
        # coords = '{}{}'.format(colnum_string(ind), maxheight)
        # if server_status["Health"] == 'OK':
        #     worksheet.write(coords, toStr(server_status["Health"], coords), green_cell)
        # else:
        #     worksheet.write(coords, toStr(server_status["Health"], coords), red_cell)
        # #appending power status
        # ind = ind + 1
        # if maxheight == 2:
        #     coords = '{}1'.format(colnum_string(ind))
        #     worksheet.write(coords, toStr('Power status', coords), header_cell)
        # coords = '{}{}'.format(colnum_string(ind), maxheight)
        # if server_status["PowerState"] == 'On':
        #     worksheet.write(coords, toStr(server_status["PowerState"], coords), green_cell)
        # else:
        #     worksheet.write(coords, toStr(server_status["PowerState"], coords), red_cell)

        # # manual index correction before ip appending
        # ind = ind + 1
        # if maxheight == 2:
        #     coords = '{}1'.format(colnum_string(ind))
        #     worksheet.write(coords, toStr('IP', coords), header_cell)
        # coords = '{}{}'.format(colnum_string(ind), maxheight)
        # worksheet.write(coords, toStr(ip, coords), black_cell)


        maxheight = currheight+1
    #print('maxcoords track',maxheight, ind)
    for m in maxwidth:
        worksheet.set_column('{}:{}'.format(m,m), maxwidth[m]*1.2)
    #appending failover result
    # {'PDU-1':{'result':'pass',sensor1:{'wattage':'100','temp':'27'},sensor2:{'wattage':'100','temp':'27'}...}, 'PDU-2':...}
    if len(failoverresult) > 0:
        # writing header
        worksheet.write('A{}'.format(maxheight), 'PDU failover check:', header_cell)
        #writing table header
        maxheight = maxheight + 1
        worksheet.write('A{}'.format(maxheight), toStr('Result', 'A{}'.format(maxheight)), header_cell)
        worksheet.write('B{}'.format(maxheight), toStr('Power', 'B{}'.format(maxheight)), header_cell)
        worksheet.write('C{}'.format(maxheight), toStr('Temperature', 'C{}'.format(maxheight)), header_cell)

        for num, res in enumerate(failoverresult):
            maxheight=maxheight + 1
            coords = 'A{}'.format(maxheight)
            if failoverresult[res]['result'] == 'fail':
                worksheet.write(coords, toStr("{} {}".format(res, failoverresult[res]['result']), coords), red_cell)
            if failoverresult[res]['result'] == 'pass':
                worksheet.write(coords, toStr("{} {}".format(res, failoverresult[res]['result']), coords), green_cell)
                #iretation over sensors based on global sensors\pdu amount
                for snum, sres in enumerate(pdus, 1):
                    pdu = 'PDU-%s' % snum
                    sensor = 'sensor%s' % snum
                    #writing wattage
                    coords = 'B{}'.format(maxheight)
                    worksheet.write(coords, toStr("{}: {}W".format(pdu, failoverresult[res][sensor]['wattage']), coords))
                    #writing temp
                    coords = 'C{}'.format(maxheight)
                    worksheet.write(coords, toStr("{}: {}C".format(sensors[sensor], failoverresult[res][sensor]['temp']), coords))
                    maxheight = maxheight + 1
            else:
                worksheet.write(coords, toStr("{} {}".format(res, failoverresult[res]['result']), coords), red_cell)
    #workbook.close()

def writetoxlsx(report_file_name, cur_report, workbook):
    rep_type = cur_report['rep_type']
    #overriding report type for
    geometry='rows'
    if rep_type == 'config_report':
        geometry = 'rows'
    if rep_type == 'hwinvent_report':
        geometry = 'columns'
    #remooving attribute
    cur_report=cur_report['report']
    #for column wide calculation purpose
    maxwidth = {}
    #creating xls file
    #header
    header_cell = workbook.add_format()
    header_cell.set_bold()
    #green cell - passed validation against master file
    green_cell = workbook.add_format()
    green_cell.set_font_color('green')
    #red cell - failed validation against master file
    red_cell = workbook.add_format()
    red_cell.set_font_color('red')
    #black_cell - dynamic data such as SN that non need to be validated
    # ( added 'excluded_for_validation': 1 to results in report constructor)
    black_cell = workbook.add_format()
    black_cell.set_font_color('gray')
    #yellow cell in case of result is not found in master file
    orange_cell = workbook.add_format()
    orange_cell.set_font_color('orange')
    #create worksheet
    worksheet = workbook.add_worksheet(report_file_name)

    #helper to calculate and update width for column
    def toStr(val, coord):
        if val == None:
            val = ''
        try:
            curr = maxwidth[coord[0]]
            if curr < len(val):
                maxwidth[coord[0]] = len(val)
        except KeyError:
            maxwidth[coord[0]] = len(val)
        return str(val)

    if geometry == "columns":
        for i, result in enumerate(cur_report, 1):
            #extracting data values list
            res = cur_report[result]
            # #header
            coords='{}1'.format(colnum_string(i))
            worksheet.write(coords, toStr(result, coords), header_cell)
            for ind, v in enumerate(res, 2):
                coords = '{}{}'.format(colnum_string(i), ind)
                for data, valid in v.items():
                    golden = v['golden']
                    #cell coloring based on value
                    if valid == 0:
                        worksheet.write(coords, toStr('fail', coords), red_cell)
                        worksheet.write_comment(coords, '\"{}\" not equal golden setting \"{}\" '.format(data,golden))
                    elif valid == 1:
                        worksheet.write(coords, toStr('pass', coords), green_cell)
                    elif valid == 2:
                        worksheet.write(coords, toStr(data, coords), black_cell)
                    elif valid == 5:
                        worksheet.write(coords, toStr(data, coords), orange_cell)
                        worksheet.write_comment(coords, 'data not found in master, should be {}'.format(golden))

        #print(maxwidth)
    if geometry == 'rows':
        for i, result in enumerate(cur_report, 1):
            res = cur_report[result]
            #print(i, data, ascii_uppercase[i])
            for r in res:
                # header
                coords = 'A{}'.format(i)
                worksheet.write(coords, toStr(result, coords))
                # in case of multiple values data
                for ind, v in enumerate(res, 1):
                    for data, valid in v.items():
                        golden = v['golden']
                        coords = '{}{}'.format(colnum_string(ind+1), i)
                        if valid == 0:
                            worksheet.write(coords, toStr('failed', coords), red_cell)
                            worksheet.write_comment(coords, '\"{}\" not equal golden setting \"{}\" '.format(data, golden))
                        elif valid == 1:
                            worksheet.write(coords, toStr('passed', coords), green_cell)
                        elif valid == 2:
                            worksheet.write(coords, toStr(data, coords), black_cell)
                        elif valid == 5:
                            worksheet.write(coords, toStr(data, coords),orange_cell)
                            worksheet.write_comment(coords, 'data not found in master, should be {}'.format(golden))

    #sheet setup for better look
    for m in maxwidth:
        worksheet.set_column('{}:{}'.format(m, m), maxwidth[m])
    # summary = {}
    # failoverresult = {}
    #workbook.close()

#report constructor
def report(xml):
    results = []
    #probing for hwinventory by checking via getdata that request invoking a ServiceTag
    service_tag = getdata(xml, classname='DCIM_SystemView', name='ServiceTag')
    if len(service_tag) == 1 and len(service_tag[0]) == 7:
        service_tag=service_tag[0]
        print('hwinventory configuration data for {} discovered {}'.format(service_tag, xml))
        rep_type = 'hwinvent_report'
        for hwrequest in hw_collect:

            results.append({hwrequest['displayname']: getdata(xml, classname=hwrequest['classname'], name=hwrequest['name']),
                            'excluded_for_validation': hwrequest['excluded_for_validation']})
        # compare 1=to be validated, 0=without validation(data not to be validated - serial numbers, et.c.)
    #probing for configuration data
    else:
        #checking for ServiceTag directly in root attribute
        try:
            service_tag = getroot(xml).attrib['ServiceTag']
            print('configuration data for {} discovered {}'.format(service_tag, xml))
            rep_type = 'config_report'
            #possibly its configuration, trying to request ServiceTag via document root
            #implement same interface as for getdata with only difference that all data vill be invoked by
            # by looping over xml data
            configitems = getdata(xml)
            for conf in configitems:
                for param, value in conf.items():
                    #print('>>>>>',param ,value)
                    results.append({param: [value]})
        #in case of both requests failed - writing some error info
        except:
            return {'rep_type': 'error', 'service_tag': 'n/a', 'report': {'error: unsupported file:'+xml: {'data': [0], 'valid': 0}}}

    #building data structure
    resData = {}
    for r in results:
        for key in r:
            #generating entries only for data keys (not for 'excluded_for_validation' "input" key or something else)
            if key != 'excluded_for_validation':
                #in case of compare attribute not defined - adding validation to be executed
                try:
                    excluded = r['excluded_for_validation']
                except KeyError:
                    excluded = 0
                if excluded:
                    #validated = 2  to avoid further validation and make grey colored value
                    validated = 2
                else:
                    validated = 0
                resData[key] = {'data': r[key], 'valid': validated}
    resData = {'rep_type': rep_type, 'service_tag': service_tag, 'report' : resData}
    return resData

# def sendrep(sysserial):
#     try:
#         curtime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
#         fromaddr = "jade@nextra01.xiv.ibm.com"
#         toaddrs = ['IBM-IVT@tel-ad.co.il']
#         subject = "Afterloan server " + sysserial + " test ended " + str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
#         msg = MIMEMultipart()
#         msg["From"] = fromaddr
#         msg["To"] = ",".join(toaddrs)
#         msg["Subject"] = subject
#         html = """
#         <html>
#         <head>
#         <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
#         </head>
#         <body>
#         <strong><b>Afterloan server passed:</b></strong>
#         """ + sysserial + """
#         <p><u>Log files were transferred to wiki, please follow the link below:</u><br>
#         <a href="http://10.148.38.142/wiki/doku.php?id=lenovo:x3650m5:"""+ sysserial +"""">http://10.148.38.142/wiki/doku.php?id=lenovo:x3650m5:""" + sysserial + """ </a>
#         <p>Using """ + os.path.realpath(__file__) +""" <p>
#         <p>Generated at: """ + curtime + """
#         <p>Tel-Ad IVT Team.<br>
#         All Rights Reserved to Tel-Ad Electronics LTD. © 2017
#         </body></html>
#         """
#         msg.attach(MIMEText(html, 'html'))
#         server = smtplib.SMTP()
#         server.connect('localhost')
#         # server.send_message(msg)
#         text = msg.as_string()
#         server.sendmail(fromaddr, toaddrs, text)
#         server.quit()
#     except:
#        if ConnectionRefusedError():
#            print('SMTP connection error, please check network and local Sendmail server')




# if __name__ == "__main__":
#     cmdlist=[]
#     for num in range(13,17):
#         onoff='OFF '
#         outlname='Master_'
#         cmdlist.append(''.join([onoff,outlname,str(num)]))
#     print(cmdlist)
#     command(cmdlist)
        #print(''.join([onoff,outlname,str(num)]))

if __name__ == "__main__":
    main(sys.argv[1:])
