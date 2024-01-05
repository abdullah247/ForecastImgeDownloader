
import os
import re
import shutil
from datetime import datetime, date, timedelta
from tkinter import Tk, filedialog, messagebox
from bs4 import BeautifulSoup
import requests
import win32com
from PyQt5.QtGui import QPalette
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QMessageBox, QMainWindow, QDialog

from PyQt5 import uic, QtGui, QtCore
import sys
USERARG='Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.5060.134 Mobile Safari/537.36 Edg/103.0.1264.77'
from excelExport import ExportImagestoExcel
def numOfDays(date1, date2):
  #check which date is greater to avoid days output in -ve number

    try:
        if date2==date1:
            return 0
        if date2 > date1:
            return (date2-date1).days
        else:
            return (date1-date2).days
    except Exception as e:
        print(e)

class UI(QMainWindow):

    def __init__(self):
        super(UI ,self).__init__()
        uic.loadUi(os.path.join(os.getcwd(),"Screen.ui") ,self)

        self.btn1.clicked.connect(self.browse1)
        self.btn2.clicked.connect(self.browse2)
        self.btn3.clicked.connect(self.getImages)
        self.btn4.clicked.connect(self.importImages)
        self.btn5.clicked.connect(self.analysis)
        self.btn6.clicked.connect(self.browse3)
        self.show()

    def browse2(self):

        stylesheet = self.btn2.styleSheet()
        self.btn2.setStyleSheet( str(stylesheet)+f"\nborder:1px solid red;")
        delim=".xlsx"
        QApplication.processEvents()
        try:
            root = Tk()
            root.withdraw()
            filetypes = (('Files', delim), ('All files', '*.*'))
            self.txt2.setText(
                str(filedialog.askopenfilename(title='Select a file', filetypes=filetypes)).replace("/", "\\"))


        except:

            pass

        self.btn2.setStyleSheet(stylesheet)
        QApplication.processEvents()

    def browse3(self):

        stylesheet = self.btn6.styleSheet()
        self.btn6.setStyleSheet( str(stylesheet)+f"\nborder:1px solid red;")

        QApplication.processEvents()
        try:
            root = Tk()
            root.withdraw()
            self.txt3.setText(
                str(filedialog.askdirectory(title='Select a folder')).replace("/", "\\"))


        except:

            pass

        self.btn6.setStyleSheet(stylesheet)
        QApplication.processEvents()

    def browse1(self):

        stylesheet = self.btn1.styleSheet()
        self.btn1.setStyleSheet( str(stylesheet)+f"\nborder:1px solid red;")
        delim="*.*"
        QApplication.processEvents()
        try:
            root = Tk()
            root.withdraw()
            filetypes = (('Files', delim), ('All files', '*.*'))
            self.txt1.setText(
                str(filedialog.askopenfilename(title='Select a file', filetypes=filetypes)).replace("/", "\\"))


        except:

            pass

        self.btn1.setStyleSheet(stylesheet)
        QApplication.processEvents()


    def importImages(self):
        stylesheet = self.btn4.styleSheet()
        self.btn4.setStyleSheet(str(stylesheet) + f"\ncolor:blue;")
        QApplication.processEvents()
        try:
            ppt = win32com.client.Dispatch('PowerPoint.Application')
            ppt.visible = True

            ppres=ppt.Presentations.Open(str(self.txt1.text()).strip(),False)

            ppt.Run(f"{ os.path.basename(str(self.txt1.text()).strip())}!UpdateShape.updateFromFolder")

            ppres.SaveAs(str(self.txt1.text()).strip())
            ppres.close()
            ppt.quit()
        except Exception as e:
            root = Tk()
            root.withdraw()
            messagebox.showinfo("Error ", f"{e.__str__()}")

        self.btn4.setStyleSheet(stylesheet)
        root = Tk()
        root.withdraw()
        messagebox.showinfo("information", "Completed")
        self.progressBar.setValue(0)
        QApplication.processEvents()
        pass


    def analysis(self):
        stylesheet = self.btn5.styleSheet()
        self.btn5.setStyleSheet( str(stylesheet)+f"\ncolor:blue;")
        QApplication.processEvents()
        try:
            s=str(self.txt3.text()).strip()
            if len(s)<1:
                ExportImagestoExcel(str(self.txt1.text()).strip(),str(self.txt2.text()).strip(),os.path.join(os.getcwd(),"SubImages"))
            else:
                ExportImagestoExcel(str(self.txt1.text()).strip(), str(self.txt2.text()).strip(),s)
        except Exception as e:
            root = Tk()
            root.withdraw()
            messagebox.showinfo("Error ", f"{e.__str__()}")

        self.btn5.setStyleSheet(stylesheet)
        root = Tk()
        root.withdraw()
        messagebox.showinfo("information", "Completed")
        self.progressBar.setValue(0)
        QApplication.processEvents()



    def getImages(self):
        stylesheet = self.btn3.styleSheet()
        self.btn3.setStyleSheet( str(stylesheet)+f"\ncolor:blue;")
        QApplication.processEvents()



        try:
            myArray = self.getMothData()
            print("Moving To next")
            Allurls = self.getAllpastwxUrls()

        except Exception as e:
            print(e)
        try:
            base=os.getcwd()

            address=str(self.txt3.text()).strip()
            if len(address)<1:
                outpath = os.path.join(os.getcwd(),"DownloadedImages")
            else:
                outpath=address

            if os.path.exists(outpath):
                shutil.rmtree(outpath)

            os.makedirs(outpath)




            file1 = open(os.path.join(base,"Urls.txt"), 'r')
            Lines = file1.readlines()

            count = 0
            # Strips the newline character
            for index,line in enumerate(Lines):
                count += 1
                arr=str(line).split(",")
                if arr[0] != "Name":
                    # date1 = date(2023, 10, 24)
                    # date2 = datetime.today().date()
                    flag = True
                    day=0
                    while (flag):
                        try:
                            prog= int(((index+1)/len(Lines) ) *100)  if int(((index+1)/len(Lines) ) *100)  <=100 else 100
                            self.progressBar.setValue(prog)
                            QApplication.processEvents()
                            name=str(arr[1]).strip().replace("\n","")
                            url=""
                            try:
                                if not "/" in name:

                                    url=Allurls[name]
                                else:

                                    name=name.replace("/","")
                                    for dict in myArray:

                                       if name in dict and len(url)<1:
                                            url=dict[name]
                                            break


                            except Exception as e:
                                print("st " +url + " as "+ e)

                            # url=str(url).replace("20231024",datetime.strftime(date2, '%Y%m%d')).strip().replace("\n","")
                            print("Before Downloading",url)
                            response = requests.get(url, stream=True)



                            if response.headers.get('content-type') =="image/png":
                                # Write output to a file
                                with open(os.path.join(outpath, arr[0]), 'wb') as out_file:
                                    shutil.copyfileobj(response.raw, out_file)
                                flag=False

                                # day = day + 1
                                # date2 = (datetime.today() - timedelta(days=day)).date()
                            else:
                                print(response.headers.get('content-type'))
                            del response
                        except Exception as e:
                            print("st " +url + " Exception " + str(e))
                            # day = day+1
                            # date2 = (datetime.today() - timedelta(days = day)).date()
                            pass
                    # with open(os.path.join(outpath,arr[0]), "wb") as f:
                    #             im = requests.get(url)
        except Exception as e:
            root = Tk()
            root.withdraw()
            messagebox.showinfo("Error ", f"{e.__str__()}")

        self.btn3.setStyleSheet(stylesheet)

        root = Tk()
        root.withdraw()
        messagebox.showinfo("information", "Completed")
        self.progressBar.setValue(0)
        QApplication.processEvents()

        pass


    def getAllpastwxUrls(self):

        AllUrls = {}
        url = "https://www.worldagweather.com/pastwx/"
        session = requests.Session()
        session.trust_env = False
        print("getting All past")
        r = session.get(url, headers={'User-Agent': USERARG})
        print("was her")
        parsedWebPage = BeautifulSoup(r.content, "html.parser")
        table = parsedWebPage.find("table")
        print("Got Table")
        rows = table.find_all("tr")
        print("Getiing PastwxUrls ")
        for i in range(3, len(rows)):
            try:
                name =rows[i].find("a")["href"]
                foundurl = url + rows[i].find("a")["href"]
                # cudate = str(rows[i].find_all("td")[2].text).split(" ")[0].split("-")
                # newdate = date(int(cudate[0]), int(cudate[1]), int(cudate[2]))
                # print(url + rows[i].find("a")["href"])
                # print(str(rows[i].find("a").text).strip().replace("..>", ""))
                # print(cudate[0], cudate[1], cudate[2])
                name=re.sub("\d*.png",".png",name)
                print(name)
                AllUrls[name] = foundurl


            except Exception as e:
                print("Normal Exception " + str(e))

                pass
        return AllUrls

    def getMothData(self):
        print("Getting Month Data")
        mainurl = ["https://www.tropicaltidbits.com/analysis/models/cfs-avg/","https://www.tropicaltidbits.com/analysis/models/cfs-mon/"]

        dict1 = {}
        dict2 = {}
        dict3 = {}
        dict4 = {}
        array = [dict1, dict2, dict3,dict4]
        for counter,murl in enumerate(mainurl):
            r = requests.get(murl, headers={'User-Agent': USERARG})
            parsedWebPage = BeautifulSoup(r.content, "html.parser")
            mainanchor = parsedWebPage.find_all("a")

            maxval=5
            for i in range(1, maxval):
                url = murl + mainanchor[len(mainanchor) - i]["href"]
                print("Getting Url", url)
                r = requests.get(url, headers={'User-Agent': USERARG})
                parsedWebPage = BeautifulSoup(r.content, "html.parser")
                anchor = parsedWebPage.find_all("a")
                for a in anchor:
                    try:
                        index = i-1 if counter<1 else 3
                        if not url.endswith("/"):
                            url=url+"/"

                        array[index][a.text] = url + a["href"]
                        print(array[index][a.text])
                    except Exception as e:
                        print(e)

        return array

#
# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app = QApplication(sys.argv)
    app.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)

    UIWindow = UI()

    try:
        sys.exit(app.exec_())
    except SystemExit:
        print("Closing")








        # print(rows[i].find("td")[3].text)

#
# # See PyCharm help at https://www.jetbrains.com/help/pycharm/
#
#
#
#
#
# folder = "PPTIMAGES"
#
#
# def getImages(base,OutputFolder):
#
#
#
#
#
#     # #
#     # #
#     # #
#     # #
#
#     options = webdriver.ChromeOptions();
#     options.add_argument('window-size=1200x600');
#     # options.add_argument("--headless")  #comment this line to make visible
#
#     driver = webdriver.Chrome(options=options)
#     #
#
#     if not os.path.exists(os.path.join(base, folder)):
#         os.makedirs(os.path.join(base, folder))
#
#     os.chdir(os.path.join(base, folder))
#
#     for data in mydata:
#         url = data["url"]
#         driver.get(url)
#         time.sleep(4)
#         while len(driver.find_elements(By.CSS_SELECTOR, "#" + data["id"])) < 1:
#             print(driver.find_elements(By.CSS_SELECTOR, "#" + data["id"]), data["id"],
#                   len(driver.find_elements(By.CSS_SELECTOR, "#" + data["id"])))
#             time.sleep(2)
#         while "loading.gif" in driver.find_element(By.ID, data["id"]).get_attribute("src"):
#             print(driver.find_element(By.ID, data["id"]).get_attribute("src"), "waiting for loading to go away")
#             time.sleep(2)
#         if "gif" in data["name"]:
#             print(driver.find_element(By.NAME, data["id"]).get_attribute("src").replace("ce1",
#                                                                                         "ce" + str(data["name"][3])))
#         else:
#             print(driver.find_element(By.ID, data["id"]).get_attribute("src"))
#
#         name = data["name"]
#         with open(name, "wb") as f:
#             if "gif" in data["name"]:
#                 im = requests.get(driver.find_element(By.NAME, data["id"]).get_attribute("src").replace("ce1",
#                                                                                                         "ce" + str(data[
#                                                                                                                        "name"][
#                                                                                                                        3])))
#             else:
#                 im = requests.get(driver.find_element(By.ID, data["id"]).get_attribute("src"))
#             f.write(im.content)
#
#     driver.close()
#
# # if __name__=="__main__":
# #     base = os.getcwd()
# #     #
# #     print("yes")
# #     OutputFolder = "Output"
# #     if not os.path.exists(os.path.join(base, OutputFolder)):
# #         os.makedirs(os.path.join(base, OutputFolder))
# #
# #
# #     outAddress = os.path.join(base, OutputFolder)
# #     getImages(base, OutputFolder)
# #     # cropit(os.path.join(base,folder))
# #
# #     # url="http://www.worldagweather.com/fcstwx/pcp_ens_anom_q50_af_2137.png"
# #     # response = requests.get(url, stream=True)
# #     # with open(os.path.join(base,"img.jpg"), 'wb') as out_file:
# #     #     shutil.copyfileobj(response.raw, out_file)
# #     # del response
# #     # with open(os.path.join(base,"img.jpg"), "wb") as f:
# #     #         im = requests.get(url)
# #
# #     ppt = win32com.client.Dispatch('PowerPoint.Application')
# #     ppt.visible = True
# #     ppres=ppt.Presentations.Open(os.path.join(base,pptfile),False)
# #     ppt.Run('updateFromFolder')
# #
# #     ppres.SaveAs(os.path.join(base,pptfile))
# #     ppres.close()
# #     ppt.quit()
#     print("Complete")