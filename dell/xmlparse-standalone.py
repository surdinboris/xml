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
       # gathering results for regular data
        if i.attrib[classnameattr] == classname:
            props=i.findall('PROPERTY')
            for prop in props:
                if prop.attrib['NAME'] == name:
                    val= prop.find('VALUE').text
                    listval.append(val)

    return(listval[0] if len(listval)==1 else listval)


def main(argv):
    inputdir = ''
    outputdir = ''
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
    filesProcessing(inputdir, outputdir)

def filesProcessing(inputdir, outputdir):
    masterRepo = report(os.path.join(inputdir, master))
    print('Master report generated from HardwareInventory.master \n', masterRepo)
    for inputfile in os.listdir(inputdir):
        fn, ext = (os.path.splitext(inputfile))
        if ext == '.xml':
            workbook = xlsxwriter.Workbook(os.path.join(outputdir, os.path.join(inputdir,inputfile)) + '_report.xlsx')
            worksheet = workbook.add_worksheet()
            print('Found xml files:', fn)
            print('Processing files...')
            #report generation
            repo = report(os.path.join(inputdir, inputfile))

            writetoxlsx(worksheet, repo, geometry='columns')
            workbook.close()
            # reportfile.close()
            # sendrep(sysserial)

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
def writetoxlsx(worksheet, results, geometry='rows'):
    maxwidth = {}
    #helper to calculate and update with for column
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
            print(i, result,ascii_uppercase[i])
            for r in result:
                #header
                coords='{}1'.format(ascii_uppercase[i])
                worksheet.write(coords, toStr(r,coords))
                #in case of multiple values data
                if type(result[r]) == list and len(result[r]) > 1 :
                    for ind, v in enumerate(result[r],2):
                        coords = '{}{}'.format(ascii_uppercase[i], ind)
                        worksheet.write(coords, toStr(v, coords))
                else:
                    coords = '{}2'.format(ascii_uppercase[i])
                    worksheet.write(coords, toStr(result[r], coords))
    if geometry == 'rows':
        for i, result in enumerate(results, 1):
            print(i, result)
            for r in result:
                # header
                coords = 'A{}'.format(i)
                worksheet.write(coords, toStr(r, coords))
                # in case of multiple values data
                if type(result[r]) == list and len(result[r]) > 1:
                    for ind, v in enumerate(result[r]):
                        # need to enumerate with letters ascii_uppercase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
                        coords = '{}{}'.format(ascii_uppercase[ind + 1], i)
                        worksheet.write(coords, toStr(v, coords))
                else:
                    coords = 'B{}'.format(i)
                    worksheet.write(coords, toStr(result[r], coords))
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
    #xls - add data
    results.append({'ServiceTag': getdata(xml, classname='DCIM_SystemView', name='ServiceTag')})
    results.append({'CPU model': getdata(xml, classname='DCIM_CPUView', name='Model')})
    #PCI
    results.append({'PCI device': getdata(xml, classname='DCIM_PCIDeviceView', name='Description')})
    #Memory
    results.append({'System memory size': getdata(xml, classname='DCIM_SystemView', name='SysMemTotalSize')})
    results.append({'Memory serial': getdata(xml, classname='DCIM_MemoryView', name='SerialNumber')})
    results.append({'Memory module part number': getdata(xml, classname='DCIM_MemoryView', name='PartNumber')})
    results.append({'Memory slot': getdata(xml, classname='DCIM_MemoryView', name='FQDD')})
    #HDD
    results.append({'HDD serial': getdata(xml, classname='DCIM_PhysicalDiskView', name='SerialNumber')})
    results.append({'HDD model': getdata(xml, classname='DCIM_PhysicalDiskView', name='Model')})
    results.append({'HDD fw': getdata(xml, classname='DCIM_PhysicalDiskView', name='Revision')})
    results.append({'HDD slot population': getdata(xml, classname='DCIM_PhysicalDiskView', name='Slot')})
    #PSU
    results.append({'PSU part number': getdata(xml, classname='DCIM_PowerSupplyView', name='PartNumber')})
    results.append({'PSU serial': getdata(xml, classname='DCIM_PowerSupplyView', name='SerialNumber')})
    results.append({'PSU model': getdata(xml, classname='DCIM_PowerSupplyView', name='Model')})
    results.append({'PSU fw': getdata(xml, classname='DCIM_PowerSupplyView', name='FirmwareVersion')})
    return results


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