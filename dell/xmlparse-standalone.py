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

def getdata(xml,classname, name, rawsearch=None):
    listval = []
    with open(xml, 'r') as x:
        data = x.read()
    root = ET.fromstring(data)


    inst = root.findall('DECLARATION/DECLGROUP/VALUE.OBJECT/INSTANCE')

    for i in inst:
        if name == 'RawResult':  #searching in linear PCI inventory data
            if i.attrib['CLASSNAME'] == classname:
                vals=i.findall('PROPERTY.ARRAY/VALUE.ARRAY/VALUE')
                for val in vals:
                    if val.text and val.text.find(rawsearch) != -1:
                        listval.append(val.text)

                    if val.text and val.text.find('[SN]')!= -1 and len(listval)>0 and len(listval)%2 !=0:
                        listval.append(val.text)

        if name == 'RawResults':  # gathering results for raw pci data
            if i.attrib['CLASSNAME'] == classname:
                vals=i.findall('PROPERTY.ARRAY/VALUE.ARRAY/VALUE')
                for val in vals:
                    if val.text.find(rawsearch) == 0:
                        if val.text not in listval:
                             listval.append(val.text)

        if name == 'Name' and rawsearch:  # gathering results for FRU data from two matching values inside one instance
            if i.attrib['CLASSNAME'] == classname:
                props = i.findall('PROPERTY')
                for prop in props:
                    if prop.attrib['NAME'] == name and prop.find('VALUE').text == rawsearch:
                        for prop in props:
                            if prop.attrib['NAME'] == 'SerialNumber':
                                val = prop.find('VALUE').text
                                listval.append(val)

        else:  # gathering results for regular data
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
        # if ext == '.xml':
        #     print('Found xml files:', fn)
        #     print('Processing files...')
        #     report(os.path.join(inputdir,inputfile),outputdir)
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


def report(xml,outputdir):
    sysserial = getdata(xml, classname='IBMSG_ComputerSystem', name='SerialNumber')
    dsabuild  =getdata(xml, classname='IBMSG_DsaVPD', name='BuildNumber')
    dsaver  =getdata(xml, classname='IBMSG_DsaVPD', name='Version')
    ipmibuild  =getdata(xml, classname='IBMSG_IPMIFirmwareElement', name='BuildNumber')
    ipmiver  = getdata(xml, classname='IBMSG_IPMIFirmwareElement', name='Version')
    uefibuild = getdata(xml, classname='IBMSG_SystemBIOSElement', name='BuildNumber')
    uefiver = getdata(xml, classname='IBMSG_SystemBIOSElement', name='Version')

    raidserial = getdata(xml, classname='LSIESG_PhysicalCard', name='SerialNumber')
    raidfw = getdata(xml, classname='LSIESG_FirmwarePackageIdentity', name='VersionString')

    drslot = getdata(xml, classname='LSIESG_PhysicalDrive', name='Slot_No')
    drpartnumber = getdata(xml, classname='LSIESG_PhysicalDrive', name='PartNumber')
    drserial = getdata(xml, classname='LSIESG_PhysicalDrive', name='SerialNumber')
    disklist = zip(drslot, drpartnumber, drserial)

    ethname = getdata(xml, classname='IBMSG_BcmDeviceFirmwareElement', name='Name')

    ethfw = getdata(xml, classname='IBMSG_BcmDeviceFirmwareElement', name='Version')
    ethlist = zip(ethname,ethfw)

    mellanoxlist = getdata(xml, classname='IBMSG_PCIRawData', name='RawResult', rawsearch='Mellanox Technologies Device')

    qlogicser = getdata(xml, classname='IBMSG_QLogicFibreChannelRawData', name='RawResults',rawsearch='Serial Number')
    sbserial = getdata(xml, classname='IBMSG_IPMIFRU', name='Name',rawsearch='System Board')
    psu1serial = getdata(xml, classname='IBMSG_IPMIFRU', name='Name',rawsearch='Power Supply 1')
    psu2serial = getdata(xml, classname='IBMSG_IPMIFRU', name='Name', rawsearch='Power Supply 2')
    bpserial = getdata(xml, classname='IBMSG_IPMIFRU', name='Name', rawsearch='DASD Backplane 1')
    pcilist = getdata(xml, classname='IBMSG_PCIDevice', name='Description')

    reportfile = open(os.path.join(outputdir, xml) +'_report.log', "w")

    reportfile.write('{0}Parsing logfile {1} started{0}\n'.format('*' * 20, xml))
    reportfile.write('System serial number: {0}\n'.format(sysserial))
    reportfile.write('DSA OEM: {0} {1}\n'.format(dsabuild, dsaver))
    reportfile.write('IPMI fw: {0} {1}\n'.format(ipmibuild, ipmiver))
    reportfile.write('UEFI fw: {0} {1}\n'.format(uefibuild, uefiver))
    reportfile.write('RAID serial number: {0} firmware: {1}\n'.format(raidserial, raidfw))
    reportfile.write('Board serial number: {0}\n'.format(sbserial))
    reportfile.write('PSU1 serial number: {0}\n'.format(psu1serial))
    reportfile.write('PSU2 serial number: {0}\n'.format(psu2serial))
    reportfile.write('Backplane serial number: {0}\n'.format(bpserial))


    if len(mellanoxlist)==4:
        reportfile.write('Mellanox serial numbers: {0} and {1}\n'.format(mellanoxlist[1][23:42],mellanoxlist[3][23:42]))

    for disk in disklist:
        reportfile.write('Drive slot:{0} P/N: {1} serial: {2}\n'.format(disk[0],disk[1],disk[2]))
    for eth in ethlist:
        ethd = eth[0].replace(":", "")
        reportfile.write('Ethernet device: {0} firmware: {1}\n'.format(ethd, eth[1]))
    for qlogic in qlogicser:
        reportfile.write('Qlogic serial number: {0}\n'.format(qlogic))
    reportfile.write('=' * 40+'\n')
    for pci in pcilist:
        reportfile.write('PCI device: {0}\n'.format(pci))

    reportfile.close()
    #sendrep(sysserial)
    return sysserial,(xml+ '_report.log')


def sendrep(sysserial):
    try:
        curtime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        fromaddr = "jade@nextra01.xiv.ibm.com"
        toaddrs = ['IBM-IVT@tel-ad.co.il']
        subject = "Afterloan server " + sysserial + " test ended " + str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        msg = MIMEMultipart()
        msg["From"] = fromaddr
        msg["To"] = ",".join(toaddrs)
        msg["Subject"] = subject
        html = """	
        <html>
        <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        </head>
        <body>
        <strong><b>Afterloan server passed:</b></strong>
        """ + sysserial + """ 
        <p><u>Log files were transferred to wiki, please follow the link below:</u><br>
        <a href="http://10.148.38.142/wiki/doku.php?id=lenovo:x3650m5:"""+ sysserial +"""">http://10.148.38.142/wiki/doku.php?id=lenovo:x3650m5:""" + sysserial + """ </a> 
        <p>Using """ + os.path.realpath(__file__) +""" <p>
        <p>Generated at: """ + curtime + """
        <p>Tel-Ad IVT Team.<br>
        All Rights Reserved to Tel-Ad Electronics LTD. © 2017 
        </body></html>
        """
        msg.attach(MIMEText(html, 'html'))
        server = smtplib.SMTP()
        server.connect('localhost')
        # server.send_message(msg)
        text = msg.as_string()
        server.sendmail(fromaddr, toaddrs, text)
        server.quit()
    except:
       if ConnectionRefusedError():
           print('SMTP connection error, please check network and local Sendmail server')

if __name__ == "__main__":
    main(sys.argv[1:])