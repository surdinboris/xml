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

def getdata(xml,classname, name, rawsearch=None):
    listval = []
    with open(xml, 'r') as x:
        data = x.read()
    root = ET.fromstring(data)
    #patch to use both two types of xml retrieved via web interface and racadmin
    inst = root.findall('Component')
    classnameattr = 'Classname'
    if len(root.findall('MESSAGE')) == 1:
        inst = root.findall('MESSAGE/SIMPLEREQ/VALUE.NAMEDINSTANCE/INSTANCE')
        classnameattr = 'CLASSNAME'
    #searching for items
    for i in inst:
       # gathering results
        if i.attrib[classnameattr] == classname:
            props=i.findall('PROPERTY')
            for prop in props:
                if prop.attrib['NAME'] == name:
                    val= prop.find('VALUE').text
                    listval.append(val)

    #return listval[0] if len(listval) == 1 else listval
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
    master_report = report(os.path.join(inputdir, master))
    print('Master report generated from HardwareInventory.master \n')
    for inputfile in os.listdir(inputdir):
        fn, ext = (os.path.splitext(inputfile))
        if ext == '.xml':
            workbook = xlsxwriter.Workbook(os.path.join(outputdir, os.path.join(inputdir,inputfile)) + '_report.xlsx')
            worksheet = workbook.add_worksheet()
            print('Found xml files:', fn)
            print('Processing files...')
            #report generation
            cur_report = report(os.path.join(inputdir, inputfile))
            #report analysing
            report_analyze(cur_report, master_report)
            cur_report=report_analyze(cur_report, master_report)
            writetoxlsx(worksheet, cur_report, geometry='columns')

            workbook.close()
            # reportfile.close()
            # sendrep(sysserial)

def report_analyze(current,master):
    result={}
    #exp = {}
    #building up data structure as following:
    # {'Memory slot': [{'DIMM.Socket.A1': 1}, {'DIMM.Socket.A2': 1}, {'DIMM.Socket.A3': 1}, {'DIMM.Socket.A4': 1},
    #                  {'DIMM.Socket.B1': 1}, {'DIMM.Socket.B2': 1}, {'DIMM.Socket.B3': 1}, {'DIMM.Socket.B4': 1}],
    #  'PSU model': [{'PWR SPLY,750W,RDNT,DELTA      ': 1}, {'PWR SPLY,750W,RDNT,DELTA      ': 1}]}

    #instead of
    # {'ServiceTag': {'data': 'C1WN2S2', 'valid': 2},
    # {'CPU model': {'data': ['Intel(R) Xeon(R) CPU E5-2630 v4 @ 2.20GHz', 'Intel(R) Xeon(R) CPU E5-2630 v4 @ 2.20GHz'],
    #                'valid': 0}


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
    print(result)
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
def writetoxlsx(worksheet, results, geometry='rows'):
    maxwidth = {}

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
            for data_item in results[result]:
                print(result,"----------->>>",data_item)
            record = results[result]
            #extracting data values list
            # data = record['data']
            # #do validation coloring!
            # #header
            # coords='{}1'.format(ascii_uppercase[i])
            # worksheet.write(coords, toStr(result, coords))
            # #in case of multiple values data
            # if type(data) == list and len(data) > 1:
            #     for ind, v in enumerate(data, 2):
            #         coords = '{}{}'.format(ascii_uppercase[i], ind)
            #         worksheet.write(coords, toStr(v, coords))
            # else:
            #     coords = '{}2'.format(ascii_uppercase[i])
            #     worksheet.write(coords, toStr(data, coords))
    if geometry == 'rows':
        for i, result in enumerate(results, 1):
            record = results[result]
            data = record['data']
            #print(i, data, ascii_uppercase[i])
            for r in data:
                # header
                coords = 'A{}'.format(i)
                worksheet.write(coords, toStr(result, coords))
                # in case of multiple values data
                if type(data) == list and len(data) > 1:
                    for ind, v in enumerate(data):
                        # need to enumerate with letters ascii_uppercase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
                        coords = '{}{}'.format(ascii_uppercase[ind + 1], i)
                        worksheet.write(coords, toStr(v, coords))
                else:
                    coords = 'B{}'.format(i)
                    worksheet.write(coords, toStr(data, coords))
    #sheet setup for better look
    for m in maxwidth:
        worksheet.set_column('{}:{}'.format(m,m), maxwidth[m])


# #rows
# def writetoxlsxRow(worksheet, results):
#     maxwidth = {}
#     #helper to calculate and update with for column
#     def toStr(val, coord):
#         try:
#             curr = maxwidth[coord]
#             if curr < len(val):
#                 maxwidth[coord] = len(val)
#         except KeyError:
#             maxwidth[coord] = len(val)
#         return str(val)
#
#     #processing results
#     for i, result in enumerate(results, 1):
#         print(i, result)
#         for r in result:
#             #header
#             coords = 'A{}'.format(i)
#             worksheet.write(coords, toStr(r,coords))
#             # in case of multiple values data
#             if type(result[r]) == list and len(result[r]) > 1 :
#                 for ind, v in enumerate(result[r]):
#                     #need to enumerate with letters ascii_uppercase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
#                     coords='{}{}'.format(ascii_uppercase[ind+1],i)
#                     worksheet.write(coords, toStr(v,coords))
#             else:
#                 coords = 'B{}'.format(i)
#                 worksheet.write(coords, toStr(result[r], coords))
#
#     #sheet setup for better look
#     for m in maxwidth:
#         worksheet.set_column('{}:{}'.format(m,m), maxwidth[m])


#report generation
def report(xml):
    #ol stuff
    #reportfile = open(os.path.join(outputdir, xml) +'_report.log', "w")
    #reportfile.write('{0}Parsing logfile {1} started{0}\n'.format('*' * 20, xml))
    #reportfile.write('System serial number: {0}\n'.format(sysserial))
    #reportfile.write('System CPUs model: {0}\n'.format(cpusmodel))
    results = []
    #workbook.add_format()
    #compare 1=to be validated, 0=without validation(data not to be validated - serial numbers, et.c.)
    #xls - add data
    results.append({'ServiceTag': getdata(xml, classname='DCIM_SystemView', name='ServiceTag'), 'excluded_for_validation': 1})
    results.append({'CPU model': getdata(xml, classname='DCIM_CPUView', name='Model')})
    #PCI
    results.append({'PCI device': getdata(xml, classname='DCIM_PCIDeviceView', name='Description')})
    #Memory
    results.append({'System memory size': getdata(xml, classname='DCIM_SystemView', name='SysMemTotalSize')})
    results.append({'Memory serial': getdata(xml, classname='DCIM_MemoryView', name='SerialNumber'), 'excluded_for_validation': 1})
    results.append({'Memory module part number': getdata(xml, classname='DCIM_MemoryView', name='PartNumber')})
    results.append({'Memory slot': getdata(xml, classname='DCIM_MemoryView', name='FQDD')})
    #HDD
    results.append({'HDD serial': getdata(xml, classname='DCIM_PhysicalDiskView', name='SerialNumber'),'excluded_for_validation': 1})
    results.append({'HDD model': getdata(xml, classname='DCIM_PhysicalDiskView', name='Model')})
    results.append({'HDD fw': getdata(xml, classname='DCIM_PhysicalDiskView', name='Revision')})
    results.append({'HDD slot population': getdata(xml, classname='DCIM_PhysicalDiskView', name='Slot')})
    #PSU
    results.append({'PSU part number': getdata(xml, classname='DCIM_PowerSupplyView', name='PartNumber')})
    results.append({'PSU serial': getdata(xml, classname='DCIM_PowerSupplyView', name='SerialNumber'), 'excluded_for_validation': 1})
    results.append({'PSU model': getdata(xml, classname='DCIM_PowerSupplyView', name='Model')})
    results.append({'PSU fw': getdata(xml, classname='DCIM_PowerSupplyView', name='FirmwareVersion')})

    #building data structure
    resData = {}
    for r in results:
        for key in r:
            #generating enty only for data keys (not for compare or something else)
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
    #print(resData)
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
