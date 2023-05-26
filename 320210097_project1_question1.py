import psutil
import socket
import win32serviceutil
import servicemanager
import win32event
import win32service
import smtplib
import schedule
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

class project1q1(win32serviceutil.ServiceFramework):
    _svc_name_ = 'workload service'
    _svc_display_name_ = 'workload service'
    _svc_description='workload service'

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        self.ReportServiceStatus(win32service.SERVICE_START_PENDING)
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,servicemanager.PYS_SERVICE_STARTED,(self._svc_name_, ''))
        self.main()


    def main(self):
        while(True):
            f=open('data.txt' ,"a")
            cpu=psutil.cpu_times()
            mem = psutil.virtual_memory()
            disk=psutil.disk_usage('/')
            network=psutil.net_io_counters()
            usage = network.bytes_sent+network.bytes_recv
            f.write(f"general data is cpu:{cpu} , memory:{mem} , hdd :{disk}\n , network usage:{usage}")
            for i in psutil.process_iter():
                info=f"name ={i.name()} , cpu ={i.cpu_percent()}, memory ={i.memory_percent()} , network usage={psutil.net_io_counters(i)}\n"
                #print(info)
                f.write(info)
            f.close()
            schedule.every(12).hours.do(self.send)
            while 1:
                schedule.run_pending()
                time.sleep(1)
            
    def send(self):
        receiver="------------------------------"
        sender="----------------------------------------"
        password="--------"
        subject="this is workload service"
        email= MIMEMultipart()
        email['From']=sender
        email['To']=receiver
        email ['Subject']=subject
        email.attach(MIMEText("this is text file", 'plain'))
        filename="data.txt"
        attachment = open('data.txt', "rb")
        p = MIMEBase('application', 'octet-stream')
        p.set_payload((attachment).read())
        encoders.encode_base64(p)
        p.add_header('Content-Disposition', "attachment", filename=filename)
        email.attach(p)
        smtp=smtplib.SMTP_SSL('smtp.gmail.com',465)
        smtp.ehlo()
        smtp.login(sender,password)
        smtp.sendmail(sender, receiver,email.as_string())
        smtp.close()
    
if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(project1q1)