import win32com.client
import pythoncom
import time

class Evhandler:
    def OnBarcodeIn(self, _1):
        print(_1)
        zip = _1

        location = 1

        if location == 1 :
            if zip[0] == '1' or zip[0] == '2' :
                if int(zip[0:3])>=100 and int(zip[0:3])<=116:
                    print('Taipei City')
                elif int(zip[0:3])>=200 and int(zip[0:3])<=206:
                    print('Keelung City')
                elif int(zip[0:3])>=207 and int(zip[0:3])<=208:
                    print('New Taipei City')
                elif int(zip[0:3])>=209 and int(zip[0:3])<=212:
                    print('Lianjiang County')
                elif int(zip[0:3])>=220 and int(zip[0:3])<=253:
                    print('New Taipei City')
                elif int(zip[0:3])>=260 and int(zip[0:3])<=290:
                    print('Yilan County')            
            elif zip[0] == '3' :
                print('Taoyuan Mail Processing Center')
            elif zip[0] == '4' or zip[0] == '5' or zip[0] == '6' :
                print('Taichung Mail Processing Center')
            elif zip[0] == '7' or zip[0] == '8' or zip[0] == '9' :
                print('Kaohsiung Mail Processing Center')

        if location == 2 :
            if zip[0] == '1' or zip[0] == '2' :
                print('Taipei Mail Processing Center')
            elif zip[0] == '3' :
                if int(zip[0:3])==300:
                    print('Hsinchu City')
                elif int(zip[0:3])>=302 and int(zip[0:3])<=315:
                    print('Hsinchu County')
                elif int(zip[0:3])>=320 and int(zip[0:3])<=338:
                    print('Taoyuan City')
                elif int(zip[0:3])>=350 and int(zip[0:3])<=369:
                    print('Miaoli County')
            elif zip[0] == '4' or zip[0] == '5' or zip[0] == '6' :
                print('Taichung Mail Processing Center')
            elif zip[0] == '7' or zip[0] == '8' or zip[0] == '9' :
                print('Kaohsiung Mail Processing Center')

        if location == 3 :
            if zip[0] == '1' or zip[0] == '2' :
                print('Taipei Mail Processing Center')
            elif zip[0] == '3' :
                print('Taoyuan Mail Processing Center')
            elif zip[0] == '4' or zip[0] == '5' or zip[0] == '6' :
                if int(zip[0:3])>=400 and int(zip[0:3])<=439:
                    print('Taichung City')
                elif int(zip[0:3])>=500 and int(zip[0:3])<=530:
                    print('Changhua County')
                elif int(zip[0:3])>=540 and int(zip[0:3])<=558:
                    print('Nantou County')
                elif int(zip[0:3])==600:
                    print('Chiayi City')
                elif int(zip[0:3])>=602 and int(zip[0:3])<=625:
                    print('Chiayi County')
                elif int(zip[0:3])>=630 and int(zip[0:3])<=655:
                    print('Yunlin County')
            elif zip[0] == '7' or zip[0] == '8' or zip[0] == '9' :
                print('Kaohsiung Mail Processing Center')

        if location == 4 :
            if zip[0] == '1' or zip[0] == '2' :
                print('Taipei Mail Processing Center')
            elif zip[0] == '3' :
                print('Taoyuan Mail Processing Center')
            elif zip[0] == '4' or zip[0] == '5' or zip[0] == '6' :
                print('Taichung Mail Processing Center')
            elif zip[0] == '7' or zip[0] == '8' or zip[0] == '9' :
                if int(zip[0:3])>=700 and int(zip[0:3])<=745:
                    print('Tainan City')
                elif int(zip[0:3])>=800 and int(zip[0:3])<=852:
                    print('Kaohsiung City')
                elif int(zip[0:3])>=880 and int(zip[0:3])<=885:
                    print('Penghu County')
                elif int(zip[0:3])>=890 and int(zip[0:3])<=896:
                    print('Jinmen County')
                elif int(zip[0:3])>=900 and int(zip[0:3])<=947:
                    print('Pingtung County')
                elif int(zip[0:3])>=950 and int(zip[0:3])<=966:
                    print('Taitung County')
                elif int(zip[0:3])>=970 and int(zip[0:3])<=983:
                    print('Hualien County')

scanner = win32com.client.DispatchWithEvents("BarcodeScanner.Reader", Evhandler)
scanner.Visible = True

while 1:
    pythoncom.PumpWaitingMessages()
    time.sleep(1)

