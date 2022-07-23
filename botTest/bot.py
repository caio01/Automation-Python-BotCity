from time import sleep
from turtle import left
from typing_extensions import Self
from botcity.core import DesktopBot
from botcity.plugins.excel import BotExcelPlugin

# Instantiate the plugin
bot_excel = BotExcelPlugin()

numLista = 7
numberOfNames = 492  # 1 of 493


class Bot(DesktopBot):
    def action(self, execution=None):
        
        bot_excel.read('botTest\LISTA CLIENTES HMB 04072022.xlsx')
        listaNomes = bot_excel.get_column('C')                
        column_table = 1
        line = 'D'
        
        for i in listaNomes:  
            self.paste(i)
        
            if not self.find( "btnAtualizar", matching=0.97, waiting_time=10000):
                self.not_found("btnAtualizar")
            self.click()
            
            if not self.find( "btnLupa", matching=0.97, waiting_time=10000):
                self.not_found("btnLupa")
            self.click()
            
            if not self.find( "cmpCodigo", matching=0.97, waiting_time=2000):
                self.not_found("cmpCodigo")

            self.right_click_at(180, 345)
            
            self.type_down(wait=50)
            self.type_down(wait=50)
            self.type_down(wait=50)
            self.type_down(wait=50)
            self.enter()
            
            email = self.get_clipboard()  #função que recupera o que foi copiado (CRTL + C) anteriormente
            print(email)
            bot_excel.set_cell('D',column_table,email)
            column_table += 1
            
            bot_excel.write('botTest\LISTA' + str(numLista) + '.xlsx')
            
            #sleep(20)
            
            if not self.find( "btnFechar", matching=0.97, waiting_time=10000):
                self.not_found("btnFechar")
            self.click()
            
            if not self.find( "btnAtualizar", matching=0.97, waiting_time=10000):
                self.not_found("btnAtualizar")
            
            sleep(1)
    
            self.tab(wait=200)
            self.tab(wait=200)
            self.tab(wait=200)
            self.tab(wait=200)
                
        for i in listaNomes:
            print(str(i) + " ")

        #bot_excel.set_cell('F',3,'TESTE')
        
        #bot_excel.write('botTest\LISTA' + str(numLista) + '.xlsx')
        
        

    def not_found(self, label):
        print(f"Element not found: {label}")

if __name__ == '__main__':
    Bot.main()







