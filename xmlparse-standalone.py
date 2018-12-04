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
import xlsxwriter

def getdata(xml,classname, name, rawsearch=None):
    listval = []
    with open(xml, 'r') as x:
        data = x.read()
    root = ET.fromstring(data)
    inst = root.findall('MESSAGE/SIMPLEREQ/VALUE.NAMEDINSTANCE/INSTANCE')
    for i in inst:
        print(i,i.attrib)
       # gathering results for regular data
        if i.attrib['CLASSNAME'] == classname:
            props=i.findall('PROPERTY')
            for prop in props:
                if prop.attrib['NAME'] == name:
                    val=prop.find('VALUE').text
                    listval.append(val)

    return(listval[0] if len(listval)==1 else listval)


def main(argv):
    inputfile = ''
    outputdir = ''
    try:
        opts, args = getopt.getopt(argv, "hi:o:", ["ifile=", "ofile="])
    except getopt.GetoptError:
        print('test.py -i <inputfile> -o <outputdir>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print('test.py -i <inputfile> -o <outputdir>')
            sys.exit()
        elif opt in ("-i", "--ifile"):
            inputfile = arg
        elif opt in ("-o", "--ofile"):
            outputdir = arg
    print('Input file:', inputfile)
    print('Report outputdir:', outputdir)
    filedetect(inputfile,outputdir)


def filedetect(inputdir,outputdir):
    filelist=[]
    for inputfile in os.listdir(inputdir):
        fn, ext = (os.path.splitext(inputfile))
        if ext == '.xml':
            print('Found xml files:', fn)
            print('Processing files...')
            #report generation
            report(os.path.join(inputdir,inputfile),outputdir)
        if ext == '.gz':
            print('Found archived file:', fn)
            filelist.append(os.path.join(inputdir,inputfile))
            latest_file = max(filelist, key=os.path.getctime)
            print('Latest DSA file is:', latest_file)
            report(unpack(os.path.join(inputdir,latest_file)),outputdir)

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

def writetoxlsx(worksheet, results):
    for i, result in enumerate(results, 1):
        print(i, result)
        for r in result:
            #in case of multiple values data
            worksheet.write('A{}'.format(i), str(r))
            if type(result[r]) == list and len(result[r]) > 1 :
                for ind, v in enumerate(result[r]):
                    #need to enumerate with letters ascii_uppercase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
                    #worksheet.write('B{}'.format(ind), str(v))
            else:
                worksheet.write('B{}'.format(i), str(result[r]))



def report(xml,outputdir):
    #ol stuff

    #reportfile = open(os.path.join(outputdir, xml) +'_report.log', "w")

    results=[]
    # reportfile.write('{0}Parsing logfile {1} started{0}\n'.format('*' * 20, xml))
    # reportfile.write('System serial number: {0}\n'.format(sysserial))
    # reportfile.write('System CPUs model: {0}\n'.format(cpusmodel))

    #xls - init
    workbook = xlsxwriter.Workbook(os.path.join(outputdir, xml) +'_report.xlsx')
    worksheet = workbook.add_worksheet()
    #xls - add data
    #
    results.append({'ServiceTag': getdata(xml, classname='DCIM_SystemView', name='ServiceTag')})
    results.append({'CPUs model': getdata(xml, classname='DCIM_CPUView', name='Model')})
    results.append({'HDD serials': getdata(xml, classname='DCIM_PhysicalDiskView', name='SerialNumber')})
    results.append({'HDD fw': getdata(xml, classname='DCIM_PhysicalDiskView', name='Revision')})
    results.append({'HDD slots': getdata(xml, classname='DCIM_PhysicalDiskView', name='Slot')})
    results.append({'System memory size': getdata(xml, classname='DCIM_SystemView', name='SysMemTotalSize')})
    results.append({'System memory modules': getdata(xml, classname='DCIM_MemoryView', name='PartNumber')})

    writetoxlsx(worksheet, results)
    workbook.close()
    #reportfile.close()
    #sendrep(sysserial)
    return


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
