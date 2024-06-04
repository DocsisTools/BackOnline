class CiscoScm:
    """A sample CBR8 device"""

    ph_change_var="ph_change_value"
    num_of_cm=0
    
    def __init__(self, mac, ip, interface,state,sid,rxpwr,timingoff,numcpe,dip):
        self.mac = mac
        self.ip = ip
        self.interface = interface
        self.state = state
        self.sid = sid
        self.rxpwr = rxpwr
        self.timingoff = timingoff
        self.numcpe = numcpe
        self.dip = dip
        
        CiscoScm.num_of_cm +=1
        
        
    def ciscoshort(self):
        return '{} {} {} {}'.format(self.mac, self.ip, self.interface, self.state)
    
    def ph_change(self):
        self.interface=CiscoScm.ph_change_var
     
    @classmethod   
    def var_change(cls, amount):
        cls.raise_amount=amount
        
    @classmethod
    def from_string_input(cls,scm_str):
        mac, ip, interface, state, side, rxpwr, timeoffset, numpe, dip = scm_str.split()
        return cls(mac, ip, interface, state, side, rxpwr, timeoffset, numpe, dip)
        
cm1=CiscoScm('mac12','ip13','int14','state15','sid16','rxpwr','timingoff','numcpe','dip')
cm2=CiscoScm('mac112','ip113','int114','state115','sid116','rxpwr','timingoff','numcpe','dip')

