import xml.etree.ElementTree as ET
import gzip
import os
import os.path
import glob
import sys, getopt
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime
ascii_uppercase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
import xlsxwriter
master = 'HardwareInventory.master'

#getroot helper for use directly from report (for configuration pasing)
#  and in getdata requests (for hw inventory parsing)
def getroot(xml):
    with open(xml, 'r') as x:
        data = x.read()
    root = ET.fromstring(data)
    return root

def getdata(xml,classname='', name=''):
    root = getroot(xml)
    #data collect helper
    def collect(inst, classnameattr):
        listval = []
        for i in inst:
            # gathering results example: Component Classname="DCIM_ControllerView
            if i.attrib[classnameattr] == classname:
                props = i.findall('PROPERTY')
                for prop in props:
                    if prop.attrib['NAME'] == name:
                        val = prop.find('VALUE').text
                        listval.append(val)
        return listval
    #router to use both two types of hwinventory retrieved via web interface or
    #racadmin and additional support for segregate requests configuration parsing (possibly not needed)

    #collecting hwinventory items in case of hwinventory detected
    if root.tag =='Inventory':
        inst = root.findall('Component')
        classnameattr = 'Classname'
        return collect(inst,classnameattr)

    elif root.tag =='CIM':
        inst = root.findall('MESSAGE/SIMPLEREQ/VALUE.NAMEDINSTANCE/INSTANCE')
        classnameattr = 'CLASSNAME'
        return collect(inst, classnameattr)

    #collecting hwinventory items in case of configuration parsing detected
    #and building custom structure attribute-value pairs
    elif root.tag =='SystemConfiguration':
        listval=[]
        inst = root.findall('Component')
        for i in inst:
            # gathering results examle: FQDD="LifecycleController.Embedded.1
            props = i.findall('Attribute')
            for prop in props:
                val = prop.text
                key = prop.attrib['Name']
                listval.append({key: val})
        return listval


def main(argv):
    #fallbacks - to current workdir
    inputdir = os.getcwd()
    outputdir = os.getcwd()
    try:
        opts, args = getopt.getopt(argv, "hi:o:", ["ifile=", "ofile="])
    except getopt.GetoptError:
        print('test.py -i <inputdir> -o <outputdir>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print('test.py -i <inputdir> -o <outputdir>')
            sys.exit()
        elif opt in ("-i", "--inpdir"):
            inputdir = arg
        elif opt in ("-o", "--outpdir"):
            outputdir = arg
    print('Input file:', inputdir)
    print('Report outputdir:', outputdir)
    files_processing(inputdir, outputdir)

def files_processing(inputdir, outputdir):
    master_report_hwinvent = report(os.path.join(inputdir, master))
    master_report_config = report(os.path.join(inputdir, master))
    print('Master report generated from HardwareInventory.master \n')
    for inputfile in os.listdir(inputdir):
        fn, ext = (os.path.splitext(inputfile))
        if ext == '.xml':
            report_file = os.path.join(outputdir, os.path.join(inputdir,inputfile)) + '_report.xlsx'
            print('Found xml file: <<', fn+ext, '>> Processing...')
            #report generation
            cur_report = report(os.path.join(inputdir, inputfile))
            #routing for hwinventory  or configuration

            #report analysing
            report_analyze(cur_report, master_report_hwinvent)
            cur_report=report_analyze(cur_report, master_report_hwinvent)
            writetoxlsx(report_file, cur_report, geometry='columns')
            # reportfile.close()
            # sendrep(sysserial)

def report_analyze(current,master):
    result={}
    #exp = {}
    #building up data structure as following:
    # {'Memory slot': [{'DIMM.Socket.A1': 1}, {'DIMM.Socket.A2': 1}, {'DIMM.Socket.A3': 1}, {'DIMM.Socket.A4': 1},
    #                  {'DIMM.Socket.B1': 1}, {'DIMM.Socket.B2': 1}, {'DIMM.Socket.B3': 1}, {'DIMM.Socket.B4': 1}],
    #  'PSU model': [{'PWR SPLY,750W,RDNT,DELTA      ': 1}, {'PWR SPLY,750W,RDNT,DELTA      ': 1}]}

    #print(current, ' \nversus\n', master)
    for record in current:
        #in case of record availalable in master file
        data_per = []
        if current[record]['valid'] == 2:
            for data_item in current[record]['data']:
                data_per.append({data_item:2})
            result[record]=data_per
            continue
        try:
            master_record = master[record]
            data_per = []
            for i, data_item in enumerate(current[record]['data']):
                try:
                    master_val = master[record]['data'][i]
                except IndexError:
                    master_val = 'not availalable in master configuration'
                data_per.append({data_item: int(data_item == master_val)})
                result[record] = data_per
                #for data_item, pos in enumerate(current[record]['data']):

                #print(data_item,pos)
                #data_per.append({data_item: 1})
            result[record] = data_per
               #print('unequal', master_record['data'], current[record]['data'],'\n')
                #old result[record] = {'data': current[record]['data'], 'valid': 0}

        except KeyError:
            #if failed to find whole master values branch in master file - assign specific attribute
            data_per = []
            for data_item in current[record]['data']:
                data_per.append({data_item: 5})
            result[record] = data_per
                #continue
            #result[record] = {'data':record['data'], 'valid': 5}
            #print(master_record)
    #print(result)
    #print('resulted',result)

    return result
        #
        # if validated and validated !=5:
        #validated='t'
        #result[record] = {'data': current[record]['data'], 'valid': validated}
    #print('resulted', result)

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
#columns
#to be refactored accordingly new report structure
def writetoxlsx(report_file, results, geometry='rows'):
    maxwidth = {}
    #creating xls file
    workbook = xlsxwriter.Workbook(report_file)
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
    #create worksheet
    worksheet = workbook.add_worksheet()

    #helper to calculate and update width for column
    def toStr(val, coord):
        try:
            curr = maxwidth[coord[0]]
            if curr < len(val):
                maxwidth[coord[0]] = len(val)
        except KeyError:
            maxwidth[coord[0]] = len(val)
        return str(val)

    if geometry == "columns":
        for i, result in enumerate(results, 0):
            #extracting data values list
            res = results[result]
            # #header
            coords='{}1'.format(ascii_uppercase[i])

            worksheet.write(coords, toStr(result, coords), header_cell)
            for ind, v in enumerate(res, 2):
                coords = '{}{}'.format(ascii_uppercase[i], ind)
                for key,value in v.items():
                    data = key
                    valid = value
                    #cell coloring based on value
                    if valid == 0:
                        worksheet.write(coords, toStr(data, coords), red_cell)
                    elif valid == 1:
                        worksheet.write(coords, toStr(data, coords), green_cell)
                    elif valid == 2:
                        worksheet.write(coords, toStr(data, coords), black_cell)

        #print(maxwidth)
    if geometry == 'rows':
        for i, result in enumerate(results, 1):
            res = results[result]
            #print(i, data, ascii_uppercase[i])
            for r in res:
                # header
                coords = 'A{}'.format(i)
                worksheet.write(coords, toStr(result, coords))
                # in case of multiple values data
                for ind, v in enumerate(res):
                    for key, value in v.items():
                        data = key
                        valid = value
                    # need to enumerate with letters ascii_uppercase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
                    coords = '{}{}'.format(ascii_uppercase[ind + 1], i)
                    if valid == 0:
                        worksheet.write(coords, toStr(data, coords), red_cell)
                    elif valid == 1:
                        worksheet.write(coords, toStr(data, coords), green_cell)
                    elif valid == 2:
                        worksheet.write(coords, toStr(data, coords), black_cell)

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
        print('hwinventory configuration data for {} discovered {}'.format(service_tag[0], xml))

        # compare 1=to be validated, 0=without validation(data not to be validated - serial numbers, et.c.)
        # xls - add data
        results.append(
            {'ServiceTag': getdata(xml, classname='DCIM_SystemView', name='ServiceTag'), 'excluded_for_validation': 1})
        results.append({'Inventory date': getdata(xml, classname='DCIM_SystemView', name='LastSystemInventoryTime'),
                        'excluded_for_validation': 1})
        results.append({'CPU model': getdata(xml, classname='DCIM_CPUView', name='Model')})
        # PCI
        results.append({'PCI device': getdata(xml, classname='DCIM_PCIDeviceView', name='Description')})
        # Memory
        results.append({'System memory size': getdata(xml, classname='DCIM_SystemView', name='SysMemTotalSize')})
        results.append({'Memory serial': getdata(xml, classname='DCIM_MemoryView', name='SerialNumber'),
                        'excluded_for_validation': 1})
        results.append({'Memory module part number': getdata(xml, classname='DCIM_MemoryView', name='PartNumber')})
        results.append({'Memory slot': getdata(xml, classname='DCIM_MemoryView', name='FQDD')})
        # HDD
        results.append({'HDD serial': getdata(xml, classname='DCIM_PhysicalDiskView', name='SerialNumber'),
                        'excluded_for_validation': 1})
        results.append({'HDD model': getdata(xml, classname='DCIM_PhysicalDiskView', name='Model')})
        results.append({'HDD fw': getdata(xml, classname='DCIM_PhysicalDiskView', name='Revision')})
        results.append({'HDD slot population': getdata(xml, classname='DCIM_PhysicalDiskView', name='Slot')})
        # PSU
        results.append({'PSU part number': getdata(xml, classname='DCIM_PowerSupplyView', name='PartNumber')})
        results.append({'PSU serial': getdata(xml, classname='DCIM_PowerSupplyView', name='SerialNumber'),
                        'excluded_for_validation': 1})
        results.append({'PSU model': getdata(xml, classname='DCIM_PowerSupplyView', name='Model')})
        results.append({'PSU fw': getdata(xml, classname='DCIM_PowerSupplyView', name='FirmwareVersion')})

    #probing for configuration data
    else:
        #checking for ServiceTag directly in root attribute
        try:
            service_tag = getroot(xml).attrib['ServiceTag']
            print('configuration data for {} discovered {}'.format(service_tag, xml))
            #possibly its configuration, trying to request ServiceTag via document root
            #implement same interface as for getdata with only difference that all data vill be invoked by
            # by looping over xml data
            #getdata in key-value
            #getdata(xml)
            #resData -> to build standard data object for report analyzing
            #print('for refactoring', {
            #'LCD.1#vConsoleIndication': getdata(xml, classname='System.Embedded.1', name='LCD.1#vConsoleIndication')})

        #in case of both requests failed - writing some error info
        except:
            return {'error: unsupported file:'+xml: {'data': [0], 'valid': 0}}


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
