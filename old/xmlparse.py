import xml.etree.ElementTree as ET
from os.path import join
import os
from os import listdir
import django
from mysite import settings
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'mysite.settings')
django.setup()


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

def report(xmlname,ident):

    xml=os.path.join(settings.MEDIA_ROOT, xmlname)
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

    reportfile = open(os.path.join(settings.MEDIA_ROOT, sysserial) + '_' +ident + '.txt', "w")

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
    return sysserial,(sysserial+ '_' +ident+'.txt')
