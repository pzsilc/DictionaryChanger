import sys
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QPushButton
from PyQt5.QtGui import QIcon
from pandas import *
import pyodbc, requests, json
from colorama import Fore, Back, Style, init
init()




class Product:
    def __init__(self, headers, attrs): #product has only 2 props - code and attrs (attrs are all values of next dictionaires)
        self.code = attrs[0]
        self.attrs = {}
        for index, header in enumerate(headers):
            if header == 'Kod': continue
            self.attrs[header] = attrs[index]



class Changer:

    def __init__(self):
        self.__guid = None #key for webapi of symfonia
        self.__session = '' #session token
        with open('config.json') as config: #load guid
            c = json.loads(config.read())
            self.__guid = c['guid']
        self.__options = requests.get('http://192.168.0.234/dictionaires?name=') #load all dictionaires
        if self.__options.status_code > 400:
            raise Exception('Unable to load options from db')
        else:
            self.__options = json.loads(self.__options.text)


    def get_id_from_value(self, dimension, value):
        if isinstance(value, int) and value < 10: value = '0' + str(value) #i assume each of numeric values under 10 has to contains '0' char front of
        for i in self.__options:
            if i['dict_name'] == str(dimension) and i['val_name'] == str(value): #searching suitable dictionaire base on data in parameters
                    return i['val_id'] #return dictionaire-value-id
        return None #not found
    

    def login(self):
        res = requests.get('http://192.168.70.70:80/api/Sessions/OpenNewSession?deviceName=test', headers={
            'Authorization': 'Application {' + self.__guid + '}'
        })
        self.__session = res.text.replace('"', '')
        if '{' in self.__session:
            print(Fore.RED + "User's limit is exceeded. Try later.")
            sys.exit()
        else:
            print(Fore.CYAN + '\n' + self.__session + '\n')


    def logout(self):
        res = requests.get('http://192.168.70.70:80/api/Sessions/CloseSession', headers={
            'Authorization': 'Session {' + self.__session + '}'
        })
        print(Fore.CYAN + '\n\nFinish', res.text)
        print(Style.RESET_ALL)


    def run(self, products):
        errors = [] #list of products didn't pass
        for product in products:
            print(Style.RESET_ALL)
            data = [] #body of request to api
            for key in product.attrs:
                data.append({
                    "Code": key,
                    "Value": self.get_id_from_value(key, product.attrs[key])
                })
            res = requests.put(
                f'http://192.168.70.70:80/api/ProductDimensions/UpdateList?productCode={product.code}',
                data=json.dumps(data),
                headers={
                    'Authorization': 'Session {' + self.__session + '}',
                    'Content-Type': 'application/json'               
                }
            )
            if res.status_code < 400: 
                print(f'{product.code} został zaktualizowany') #no errors
            else:
                errors.append(product)    

        if len(errors): #listing errors (if any exists)
            print(Fore.RED + 'Something went wrong with:')
            for index, product in enumerate(errors):
                print(index+1, product.code)




class App(QWidget):

    def __init__(self):
        super().__init__()
        self.title = 'Changer'
        self.left = 300 #properties of window
        self.top = 300
        self.width = 200
        self.height = 100
        self.filename = None #path to choosed file
        self.products = list() #list of Product instances
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.attach_btn = QPushButton('Attach file (xlsx)', self) #button 1
        self.attach_btn.clicked.connect(self.attach)
        self.attach_btn.resize(160, 30)
        self.attach_btn.move(20, 20)
        self.submit_btn = QPushButton('Export', self) #button 2
        self.submit_btn.clicked.connect(self.submit)
        self.submit_btn.resize(160, 30)
        self.submit_btn.move(20, 60)
        self.show()

    def attach(self):
        self.open_file()
        if self.filename:
            xlsx = ExcelFile(self.filename) #parse choosed 
            df = xlsx.parse(xlsx.sheet_names[0]) #file
            df = df.to_dict() #to py dict
            if next(iter(df)) != 'Kod': #uncorrect format of file
                print(Fore.RED + 'Dane są w nieodpowiednim formacie')
                self.products.clear()
            else:
                headers = list(df.keys())
                for i in df['Kod']:
                    attrs = [ df[key][i] for key in df ]
                    self.products.append(Product(headers, attrs)) #save data from file to list of products

    def open_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","All Files (*);;Python Files (*.py)", options=options)
        self.filename = fileName or None

    def submit(self):
        if len(self.products) == 0:
            print(Fore.BLUE + 'You have to select file')
            return
        changer = Changer()
        changer.login()
        changer.run(self.products)
        changer.logout()



#start app
if __name__ == '__main__':
    print(Style.RESET_ALL)
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())