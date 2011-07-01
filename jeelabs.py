import os, sys
from twisted.internet.serialport import SerialPort
from twisted.protocols import basic
import ConfigParser
from plugins.pluginapi import PluginAPI
from twisted.internet import reactor
from twisted.python import log

# Platform specific imports
if os.name == "nt":
    import win32service
    import win32serviceutil
    import win32event
    import win32evtlogutil

class JeelabsProtocol(basic.LineReceiver):
    '''
    This class handles the JeeLabs protocol, i.e. the wire level stuff.
    '''
    def __init__(self, wrapper):
        self.wrapper = wrapper       
           
    def lineReceived(self, line):
        if line.startswith("OK"):
            self._handle_data(line)
            
    def _handle_data(self, line):
        '''
        This function handles incoming node data, current the following sketches/node types are supported:
        - Roomnode sketch
        - Outside node sketch
        @param line: the raw line of data received.
        '''
        data = line.split(" ")
        
        if int(data[2]) == 1:
            
            # Raw data packets (information from host.tcl (JeeLabs))
            a   = int(data[3]) 
            b   = int(data[4])
            c   = int(data[5])
            d   = int(data[6])       
            node_id = data[1]
            
            light       = a 
            motion      = b & 1
            humidity    = b >> 1
            temperature = str(((256 * (d&3) + c) ^ 512) - 512)
            battery     = (d >> 2) & 1
            temperature = temperature[0:2] + '.' + temperature[-1]
            
            log.msg("Received data from rooms jeenode; channel: %s, LDR: %s, " \
                  "humidity: %s, temperature: %s, motionsensor: %s, battery: %s" % (node_id, \
                                                                                    light, 
                                                                                    humidity, 
                                                                                    temperature, 
                                                                                    motion,
                                                                                    battery))
                  
            values = {'Light': str(light), 'Humidity': str(humidity),
                      'Temperature': str(temperature), 'Motion': str(motion), 'Battery': str(battery)}
           
            self.wrapper.pluginapi.value_update(node_id, values)         
            
        # Handle outside node sketch
        elif int(data[2]) == 2:
            
            node_id = data[1]
            
            # temperature from pressure chip (16bit)
            temp = str((int(data[4]) << 8) + int(data[3]))
            temp = temp[0:2] + '.' + temp[-1]
            
            # Lux level (32bit)
            lux = str((int(data[8]) << 24) + (int(data[7]) << 16) + (int(data[6]) << 8) + int(data[5]))
            
            # barometric pressure (32bit)
            pressure = str((int(data[-1]) << 24) + (int(data[-2]) << 16) + (int(data[-3]) << 8) + int(data[-4]))
            pressure = pressure[0:4] + "." + pressure[-2:]

            log.msg("Received data from outside sketch jeenode; channel: %s, lux: %s, " \
                  "pressure: %s, temperature: %s" % (node_id, lux, pressure, temp))
            
            values = {'Lux': str(lux), 'Pressure': str(pressure), 'Temperature': str(temp)}

            self.wrapper.pluginapi.value_update(node_id, values)  

class JeelabsWrapper():

    def __init__(self):
        '''
        Load initial JeeLabs configuration from jeelabs.conf
        '''
        from utils.generic import get_configurationpath
        config_path = get_configurationpath()
        
        config = ConfigParser.RawConfigParser()
        config.read(os.path.join(config_path, 'jeelabs.conf'))
        self.port = config.get("serial", "port")

        # Get broker information (RabbitMQ)
        self.broker_host = config.get("broker", "host")
        self.broker_port = config.getint("broker", "port")
        self.broker_user = config.get("broker", "username")
        self.broker_pass = config.get("broker", "password")
        self.broker_vhost = config.get("broker", "vhost")
        
        self.logging = config.getboolean('general', 'logging')
        self.id = config.get('general', 'id')
                
    def start(self):
        '''
        Function that starts the JeeLabs plug-in. It handles the creation 
        of the plugin connection and connects to the specified serial port.
        '''
        self.pluginapi = PluginAPI(plugin_id=self.id, plugin_type='Jeelabs', logging=self.logging,
                                   broker_ip=self.broker_host, broker_port=self.broker_port,
                                   username=self.broker_user, password=self.broker_pass, vhost=self.broker_vhost)
        
        protocol = JeelabsProtocol(self) 
        myserial = SerialPort (protocol, self.port, reactor)
        myserial.setBaudRate(57600)      

        reactor.run(installSignalHandlers=0)
        return True

if os.name == "nt":    
    
    class JeelabsService(win32serviceutil.ServiceFramework):
        '''
        This class is a Windows Service handler, it's common to run
        long running tasks in the background on a Windows system, as such we
        use Windows services for HouseAgent.
        '''        
        _svc_name_ = "hajeelabs"
        _svc_display_name_ = "HouseAgent - Jeelabs Service"
        
        def __init__(self,args):
            win32serviceutil.ServiceFramework.__init__(self,args)
            self.hWaitStop=win32event.CreateEvent(None, 0, 0, None)
            self.isAlive=True
    
        def SvcStop(self):
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
            reactor.stop()
            win32event.SetEvent(self.hWaitStop)
            self.isAlive=False
    
        def SvcDoRun(self):
            import servicemanager
                   
            win32evtlogutil.ReportEvent(self._svc_name_,servicemanager.PYS_SERVICE_STARTED,0,
            servicemanager.EVENTLOG_INFORMATION_TYPE,(self._svc_name_, ''))
    
            self.timeout=1000  # In milliseconds (update every second)
            jeelabs = JeelabsWrapper()
            
            if jeelabs.start():
                win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE) 
    
            win32evtlogutil.ReportEvent(self._svc_name_,servicemanager.PYS_SERVICE_STOPPED,0,
                                        servicemanager.EVENTLOG_INFORMATION_TYPE,(self._svc_name_, ''))
    
            self.ReportServiceStatus(win32service.SERVICE_STOPPED)
    
            return

if __name__ == '__main__':
    
    if os.name == "nt":    
        
        if len(sys.argv) == 1:
            try:
    
                import servicemanager, winerror
                evtsrc_dll = os.path.abspath(servicemanager.__file__)
                servicemanager.PrepareToHostSingle(JeelabsService)
                servicemanager.Initialize('JeelabsService', evtsrc_dll)
                servicemanager.StartServiceCtrlDispatcher()
    
            except win32service.error, details:
                if details[0] == winerror.ERROR_FAILED_SERVICE_CONTROLLER_CONNECT:
                    win32serviceutil.usage()
        else:    
            win32serviceutil.HandleCommandLine(JeelabsService)
    else:
        jeelabs = JeelabsWrapper()
        jeelabs.start()