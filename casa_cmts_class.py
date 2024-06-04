#CASA_CMTS_CLASS

#MAC Address    IP Address      US             DS           MAC         Prim RxPwr  Timing Num  BPI
#                               Intf           Intf         Status      Sid  (dBmv) Offset CPEs Enb
#AAAA.AAAA.AAAA 0.0.0.0         9/2.0/0        2/0/31       offline     0    0.0    0      0    no 
#BBBB.BBBB.BBBB 10.7.144.159    12/4.0/0       0/5/29       online(pt)  554  -0.5   2818   1    yes
class CasaScm:
    num_of_cm=0
    def __init__(self, mac, ip, us_interface,ds_interface,status,sid,rxpwr,timingoff,numcpe,bpi):
        self.mac = mac
        self.ip = ip
        self.us_interface = us_interface
        self.ds_interface = ds_interface
        self.status = status
        self.sid = sid
        self.rxpwr = rxpwr
        self.timingoff = timingoff
        self.numcpe = numcpe
        self.bpi = bpi
        
        CasaScm.num_of_cm +=1
        
    def casashort(self):
        return '{} {} {} {}'.format(self.mac, self.ip, self.us_interface, self.status)
    
    @classmethod
    def from_string_input(cls,scm_str):
        mac, ip, us_interface,ds_interface,status, side, rxpwr, timeoffset, numpe, bpi = scm_str.split()
        return cls(mac, ip, us_interface,ds_interface, status, side, rxpwr, timeoffset, numpe, bpi)
