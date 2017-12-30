import xml.etree.ElementTree as ET
from os.path import join, isfile
from os import listdir
xmlpath='c:/Xmls/'
xmlslist=[join(xmlpath, f) for f in listdir(xmlpath)]
print('Found xml files:', xmlslist)


def getdata(xml,classname, name, rawsearch=None):
    with open(xml, 'r') as x:
        data = x.read()
    root = ET.fromstring(data)
    listval = []

    inst = root.findall('DECLARATION/DECLGROUP/VALUE.OBJECT/INSTANCE')
    for i in inst:
        if name == 'IBMSG_PCIRawData':
            if i.attrib['CLASSNAME'] == classname:
                vals=i.findall('PROPERTY.ARRAY/VALUE.ARRAY/VALUE')
                for val in vals:
                    if val.text.find(rawsearch) == 0:
                         if val.text not in listval:
                             listval.append(val.text)

        if name == 'RawResults':  # gathering results for raw pci data
            if i.attrib['CLASSNAME'] == classname:
                vals=i.findall('PROPERTY.ARRAY/VALUE.ARRAY/VALUE')
                for val in vals:
                    if val.text.find(rawsearch) == 0:
                         if val.text not in listval:
                             listval.append(val.text)

        if name == 'Name' and rawsearch:  # gathering results for FRU data from two matching vals inside one instance
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


for xml in xmlslist:
    sysserial = getdata(xml, classname='IBMSG_ComputerSystem', name='SerialNumber')
    raidserial = getdata(xml, classname='LSIESG_PhysicalCard', name='SerialNumber')
    raidfw = getdata(xml, classname='LSIESG_FirmwarePackageIdentity', name='VersionString')

    drslot = getdata(xml, classname='LSIESG_PhysicalDrive', name='Slot_No')
    drpartnumber = getdata(xml, classname='LSIESG_PhysicalDrive', name='PartNumber')
    drserial = getdata(xml, classname='LSIESG_PhysicalDrive', name='SerialNumber')
    disklist = zip(drslot, drpartnumber, drserial)

    ethname = getdata(xml, classname='IBMSG_BcmDeviceFirmwareElement', name='Name')
    ethfw = getdata(xml, classname='IBMSG_BcmDeviceFirmwareElement', name='Version')
    ethlist = zip(ethname,ethfw)

    coleto = getdata(xml, classname='IBMSG_PCIRawData', name='RawResult', rawsearch='')

    qlogicser = getdata(xml, classname='IBMSG_QLogicFibreChannelRawData', name='RawResults',rawsearch='Serial Number')
    sbserial = getdata(xml, classname='IBMSG_IPMIFRU', name='Name',rawsearch='System Board')
    psu1serial = getdata(xml, classname='IBMSG_IPMIFRU', name='Name',rawsearch='Power Supply 1')
    psu2serial = getdata(xml, classname='IBMSG_IPMIFRU', name='Name', rawsearch='Power Supply 2')
    bpserial = getdata(xml, classname='IBMSG_IPMIFRU', name='Name', rawsearch='DASD Backplane 1')
    pcilist = getdata(xml, classname='IBMSG_PCIDevice', name='Description')

    print('{0}Parsing logfile {1} started{0}'.format('*'*20,xml))
    print('System serial number: {0}'.format(sysserial))
    print('RAID serial number: {0} firmware: {1}'.format(raidserial, raidfw))
    print('Board serial number: {0}'.format(sbserial))
    print('PSU1 serial number: {0}'.format(psu1serial))
    print('PSU2 serial number: {0}'.format(psu2serial))
    print('Backplane serial number: {0}'.format(bpserial))

    for disk in disklist:
        print('Drive slot:{0} P/N: {1} serial: {2}'.format(disk[0],disk[1],disk[2]))
    for eth in ethlist:
        print('Ethernet device: {0} firmware: {1}'.format(eth[0], eth[1]))
    for qlogic in qlogicser:
        print('Qlogic serial number: {0}'.format(qlogic))
    print('='*40)
    for pci in pcilist:
        print('PCI device: {0}'.format(pci))
