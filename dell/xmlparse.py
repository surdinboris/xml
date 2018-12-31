#old xml library
#import xml.etree.ElementTree as ET
from lxml import etree as ET
import gzip
import os
import os.path
import glob
import sys, getopt
import subprocess
import shutil
import time
# import smtplib
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# import datetime
import xlsxwriter
import nmap

#generator for AB style excell cells
def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

hardware_golden = 'HardwareInventory.golden'
configuration_golden = 'ConfigurationInventory.golden'

#additional attributes to collect for dynamic configuration data (FQDD, <!-- <Attribute Name=" ....)
additional_conf_collect = {}
additional_conf_collect.update({"Disk.Virtual.0:RAID.Integrated.1-1": ['Name', 'Size', 'StripeSize', 'SpanDepth', 'SpanLength', 'RAIDTypes', 'IncludedPhysicalDiskID']})
#additional_conf_collect.update({"iDRAC.Embedded.1": ["IPv4Static.1#Address"]})
# summary object init
summary = {}
#harware collection constructor
hw_collect=[]
hw_collect.append({'displayname': 'ServiceTag', 'classname': 'DCIM_SystemView', 'name': 'ServiceTag', 'excluded_for_validation': 2})
hw_collect.append({'displayname': 'Inventory date', 'classname': 'DCIM_SystemView', 'name': 'LastSystemInventoryTime', 'excluded_for_validation': 2})
hw_collect.append({'displayname': 'CPU model', 'classname': 'DCIM_CPUView', 'name': 'Model', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'PCI device', 'classname': 'DCIM_PCIDeviceView', 'name': 'Description', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'System memory size', 'classname': 'DCIM_SystemView', 'name': 'SysMemTotalSize', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'Memory serial', 'classname': 'DCIM_MemoryView', 'name': 'SerialNumber', 'excluded_for_validation': 2})
hw_collect.append({'displayname': 'Memory module part number', 'classname': 'DCIM_MemoryView', 'name': 'PartNumber', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'Memory slot', 'classname': 'DCIM_MemoryView', 'name': 'FQDD', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'HDD serial', 'classname': 'DCIM_PhysicalDiskView', 'name': 'SerialNumber', 'excluded_for_validation': 2})
hw_collect.append({'displayname': 'HDD model', 'classname': 'DCIM_PhysicalDiskView', 'name': 'Model', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'HDD fw', 'classname': 'DCIM_PhysicalDiskView', 'name': 'Revision', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'HDD slot population', 'classname': 'DCIM_PhysicalDiskView', 'name': 'Slot', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'PSU part number', 'classname': 'DCIM_PowerSupplyView', 'name': 'PartNumber', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'PSU serial', 'classname': 'DCIM_PowerSupplyView', 'name': 'SerialNumber', 'excluded_for_validation': 2})
hw_collect.append({'displayname': 'PSU model', 'classname': 'DCIM_PowerSupplyView', 'name': 'Model', 'excluded_for_validation': 0})
hw_collect.append({'displayname': 'PSU fw', 'classname': 'DCIM_PowerSupplyView', 'name': 'FirmwareVersion', 'excluded_for_validation': 0})


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
    # tree = ET.parse(xml)
    tree = getroot(xml)
    # for sc in tree.xpath('//SystemConfiguration'):
    # for compon in sc.xpath('//Component'):
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
        hwinventory = []
        for i in inst:
            # gathering results example: Component Classname="DCIM_ControllerView
            if i.attrib[classnameattr] == classname:
                props = i.findall('PROPERTY')
                for prop in props:
                    if prop.attrib['NAME'] == name:
                        val = prop.find('VALUE').text
                        hwinventory.append(val)
        return hwinventory
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
            # gathering results examle: FQDD="LifecycleController.Embedded.1
            props = i.findall('Attribute')
            for prop in props:
                val = prop.text
                key = prop.attrib['Name']
                confinventory.append({key: val})
        for FQDD in additional_conf_collect:
            confinventory.append(add_dynamic_attrs(FQDD, additional_conf_collect[FQDD],xml))
        return confinventory


def main(argv):
    # fallbacks - to current workdir
    temp = os.path.join(os.getcwd(), 'temp')
    arrived = os.path.join(os.getcwd(), 'arrived')
    passed = os.path.join(os.getcwd(), 'passed')
    def cleantemp(temp):
        for inputfile in os.listdir(temp):
            print('clearing',os.path.join(temp,inputfile))
            os.remove(os.path.join(temp,inputfile))
        if len(os.listdir(temp)) !=0:
           raise FileExistsError('Clearing of temporary dir failed, please check!')
    #retrieving hosts information
    def nmapscan():
        nm = nmap.PortScanner()
        nm.scan('10.48.228.1-40', '22')
        print("Found hosts:")
        for host in nm.all_hosts():
            print('-'*100)
            print('Host : %s' % host)
            print('State : %s' % nm[host].state())
        return nm.all_hosts()
    active_hosts= nmapscan()
    answer = input("Found {} hosts. Do you want to proceed?[y/n]".format(len(active_hosts)))
    if not answer or answer[0].lower() != 'y':
        print('Interrupting')
        exit(1)

    for host in active_hosts:
        print('\n'*2)
        print("Connecting to host {}".format(host))
        cleantemp(temp)
            ##get orig data via racadm - disabled implemented at the earlier stage:
            ###os.system("racadm -r {host} -u root -p calvin hwinventory export -f {fn}".format(host,os.path.join(temp,"hw_orig_tmp.xml")))
            # subprocess.run(["racadm", "-r", host, "-u", "root", "-p", "calvin", "hwinventory", "export", "-f",
            #                 "{}".format(os.path.join(temp,"hw_orig_tmp.xml"))])
            #
            # subprocess.run(["racadm", "-r", host, "-u", "root", "-p", "calvin", "--nocertwarn", "get", "-t", "xml", "-f",
            #                 "{}".format(os.path.join(temp,"conf_orig.tmp.xml"))])
            # files_processing(temp, arrived, step='arrived')
            #cleantemp(temp)
            ##applying golden template
            # print("Applying Golden configuration, please wait....")
            # subprocess.run(["racadm", "-r", host, "-u", "root", "-p", "calvin", "--nocertwarn", "set", "-f",
            #                 "{}".format(os.path.join(os.getcwd(), "ConfigurationInventory.golden")), "-t", "xml", "-b",
            #                 "graceful", "-w", "600", "-s", "on"])

    ##getting data after golden termplate enrollment:

        subprocess.run(["racadm", "-r", host, "-u", "root", "-p", "wildcat1", "hwinventory", "export", "-f",
                       "{}".format(os.path.join(temp,"hw_passed.xml"))])
        subprocess.run(["racadm", "-r", host, "-u", "root", "-p", "wildcat1", "--nocertwarn", "get", "-t", "xml", "-f",
                        "{}".format(os.path.join(temp,"conf_passed.xml"))])

        #verifying against golden template

        files_processing(temp, passed, step='golden', ip=host)
        cleantemp(temp)
    writesummary(os.path.join(os.getcwd(), 'summary_report.xlsx'), summary)

    #files_processing(os.getcwd(), os.getcwd(), ip='0.0.0.0')
#per server files processing
def files_processing(inputdir, outputdir, step=None, ip=None):
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
                report_file_name = os.path.join(outputdir, os.path.join(inputdir, inputfile)) + '_report.xlsx'
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
                summary[service_tag].append(cur_report)
                writetoxlsx(os.path.join(outputdir, "{}_{}_{}".format(service_tag, rep_type, fn+'_report.xlsx')), cur_report)

                counter += 1
                print('Passed report for {} stored in {}'.format(service_tag, filename))

            #default behavior - for testing only
            else:
                report_file_name = os.path.join(outputdir, os.path.join(inputdir,inputfile)) + '_report.xlsx'
                print('Found xml file: {} Processing...'.format(fn+ext))
                #report generation
                cur_report = report(os.path.join(inputdir, inputfile))
                service_tag = cur_report['service_tag']
                #report analysing
                cur_report = report_analyze(cur_report)
                try:
                    summary[service_tag]
                except KeyError:
                    summary[service_tag] = []
                summary[service_tag].append(cur_report)
                summary[service_tag].append({'ip': ip})
                writetoxlsx(report_file_name, cur_report)

                counter += 1

    #last execution block
                print('{} done. Processed {}, files'.format(service_tag, counter))
    #print('Summary', summary)
    #implement whole report building from summary
            # reportfile.close()
            # sendrep(sysserial)

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

def writesummary(report_file_name, summary):

    maxwidth = {}
    # creating xls file
    workbook = xlsxwriter.Workbook(report_file_name)
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
    worksheet = workbook.add_worksheet()

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

    # print(summary)

    i = 1
    for result in summary:
        print('Summary  detected for {}'.format(result))
        service_tag = result
        conf_passed = 0
        conf_error = 0
        hw_passed = 0
        hw_error = 0
        for res in summary[result]:
            #print('~'*30, '\n')
            #entering to report data
            try:
                if res['rep_type'] == 'config_report':
                    for ind, v in enumerate(res['report'], 0):
                    #coords = '{}{}'.format(colnum_string(i), ind)
                        #print(ind, v, )
                        conf_items=res['report'][v]
                        #print(ind,v ,hw_items)
                        for confitem in conf_items:
                            #print(hwitem)
                            for key in confitem:
                                if key != 'golden':
                                    #print(key,confitem[key])
                                    if confitem[key] == 1:
                                        conf_passed += 1
                                    if confitem[key] == 0:
                                        conf_error += 1
                                    elif confitem[key] == 2:
                                        pass

                if res['rep_type'] == 'hwinvent_report':
                    for ind, v in enumerate(res['report'], 0):
                        # coords = '{}{}'.format(colnum_string(i), ind)
                        #print(ind, v, )
                        hw_items=res['report'][v]
                        #print(ind,v ,hw_items)
                        for hwitem in hw_items:
                            #print(hwitem)
                            for key in hwitem:
                                if key != 'golden':
                                    #print(key, hwitem[key])
                                    if hwitem[key] == 1:
                                        hw_passed += 1
                                    if hwitem[key] == 0:
                                        hw_error += 1
                                    elif hwitem[key] == 2:
                                        pass
            except KeyError:
                ip = res['ip']
        worksheet.write('A{}'.format(i), toStr('Service Tag', 'A{}'.format(i)), header_cell)
        coords = 'B{}'.format(i)
        worksheet.write(coords, toStr(service_tag, coords))
        i += 1
        worksheet.write('A{}'.format(i), toStr('IP addres', 'A{}'.format(i)), header_cell)
        coords = 'B{}'.format(i)
        worksheet.write(coords, toStr(ip, coords))
        i += 1
        worksheet.write('A{}'.format(i), toStr('Conf passed', 'A{}'.format(i)), header_cell)
        coords = 'B{}'.format(i)
        worksheet.write(coords, toStr(conf_passed, coords))
        i += 1
        worksheet.write('A{}'.format(i), toStr('Conf error', 'A{}'.format(i)), header_cell)
        coords = 'B{}'.format(i)
        worksheet.write(coords, toStr(conf_error, coords))
        i += 1
        worksheet.write('A{}'.format(i), toStr('HW passed', 'A{}'.format(i)), header_cell)
        coords = 'B{}'.format(i)
        worksheet.write(coords, toStr(hw_passed, coords))
        i += 1
        worksheet.write('A{}'.format(i), toStr('HW error', 'A{}'.format(i)), header_cell)
        coords = 'B{}'.format(i)
        worksheet.write(coords, toStr(hw_error, coords))
        i += 2
        print(service_tag, "Config items pass: {}, Config errors:{} , HW items passed: {}, "
                           "HW errors: {}".format(conf_passed, conf_error, hw_passed, hw_error))

    for m in maxwidth:
        worksheet.set_column('{}:{}'.format(m,m), maxwidth[m])
    workbook.close()

def writetoxlsx(report_file_name, cur_report):
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
    workbook = xlsxwriter.Workbook(report_file_name)
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
    worksheet = workbook.add_worksheet()

    #helper to calculate and update width for column
    def toStr(val, coord):
        if val == None:
            val=''
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
        worksheet.set_column('{}:{}'.format(m,m), maxwidth[m])
    workbook.close()

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
            configitems=getdata(xml)
            for conf in configitems:
                for param, value in conf.items():
                    results.append({param:[value]})
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
    resData = {'rep_type': rep_type, 'service_tag': service_tag, 'report':resData}
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

if __name__ == "__main__":
    main(sys.argv[1:])
