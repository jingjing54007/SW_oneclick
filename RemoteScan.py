import MySQLdb,re,time,telnetlib,ping

class POWERSUPPLY():
    boolValidIpAddr = False
    boolValidComAddr = False

    def __init__(self, ip_addr = "APC_127.0.0.1"):
        p = re.compile("APC\d*_\d{1,3}\.\d{1,3}\.\d{1,3}.\d{1,3}")
        self.connection = 'NG'
        if p.match(ip_addr) is not None:
            POWERSUPPLY.boolValidIpAddr = True
            self.ip = ip_addr.split("_")[1]
            #self.outletPort = ip_addr.split("_")[2]
            print("Power Suppy IP Addr = " + self.ip)
            #print("Outlet of Power Supply = " + self.outletPort)
            try:
                if POWERSUPPLY.boolValidIpAddr:
                    tn = telnetlib.Telnet(self.ip,"23")
                    tn.read_until("User Name :",10)
                    tn.write("apc\r\n")
                    tn.read_until("Password  :",10)
                    tn.write("sf2sogo-c\r\n")
                    tn.close()
                    self.connection = 'OK'
            except Exception, e:
                print(e)
                print e
                print "---->Problem: Fail to access to APC!!!"
                
        elif 'COM' in ip_addr:
            POWERSUPPLY.boolValidComAddr = True
            self.ser = serial.Serial(ip_addr)#( ip_addr,9600,8,"N",1,"None")
            self.ser.baudrate = '9600'
            self.ser.bytesize = 8
            self.ser.parity = 'N'
            self.ser.stopbits = 1
            self.ser.timeout = 2
            self.ser.write('system:remote\r\n')
            time.sleep(1)
            self.ser.write('*IDN?\r\n')
            time.sleep(1)
            self.type=""
            if "E3631A" in self.ser.read(self.ser.inWaiting()):
                self.type = "E3631A"
            
            if  self.type == "E3631A":
                self.ser.write('appl p6v, 4.5, 5.0\r\n')
            else:
                self.ser.write('appl 4.5\r\n')
        else:
            print("Incorrect APC power address: " + ip_addr)
            print("Please correct it in sample.cfg")

    

    def __del__(self):
        if POWERSUPPLY.boolValidComAddr:
            self.ser.close()

def getAllAPC():
    try:
        db = MySQLdb.connect("cnhkg-ed-hkva17","oneclick","sierra_211","simbook" )
        cursor = db.cursor()
        cursor.execute("SELECT APC, host_pc, location, vmware FROM module_remote WHERE APC  IS NOT NULL")
        results = cursor.fetchall()
        #print results
        Apc_dict = {}
        for item in results:
            temp = item[0]
            if 'APC' not in temp:
                temp = 'APC_' + temp            
            if temp not in Apc_dict.keys():
                Apc_dict[temp] = {}
                Apc_dict[temp]['PC'] = item[1]
                Apc_dict[temp]['LOC'] = item[2]
        return Apc_dict
    except Exception, e:
        print type(e)
        print e
        return []

def getAllVM():
    try:
        db = MySQLdb.connect("cnhkg-ed-hkva17","oneclick","sierra_211","simbook" )
        cursor = db.cursor()
        cursor.execute("SELECT vmware FROM module_remote WHERE vmware IS NOT NULL")
        results = cursor.fetchall()
        #print results
        Vm_list = []
        for item in results:
            temp = item[0]            
            Vm_list.append(temp)
        return Vm_list
    except Exception, e:
        print type(e)
        print e
        return []

        

if __name__ == "__main__":
    print "\n----------------------------------------------------------------------------------"
    print "                   Scan APC"
    print "----------------------------------------------------------------------------------"
    Apc_dict = getAllAPC()
    for apc in Apc_dict.keys():
        myPower = POWERSUPPLY(apc)
        print "APC : %s, Connected to %s, Location: %s" % (myPower.connection,Apc_dict[apc]['PC'],Apc_dict[apc]['LOC'])        

    print "\n----------------------------------------------------------------------------------"
    print "                   Scan Host Machine"
    print "----------------------------------------------------------------------------------"
    for apc in Apc_dict.keys():
        print "%s at %s : " % (Apc_dict[apc]['PC'],Apc_dict[apc]['LOC'])
        result = ping.verbose_ping('%s'%Apc_dict[apc]['PC'], count=3)        
        
    print "\n----------------------------------------------------------------------------------"
    print "                   Scan Virtual Machine"
    print "----------------------------------------------------------------------------------"
    for vm in getAllVM():
        print "%s : " % (vm)
        result = ping.verbose_ping('%s'%vm, count=3)        
    
        
        
