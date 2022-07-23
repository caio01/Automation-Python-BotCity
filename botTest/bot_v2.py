from time import sleep
from turtle import left
from typing_extensions import Self
from botcity.core import DesktopBot
from botcity.plugins.excel import BotExcelPlugin

# Instantiate the plugin
bot_excel = BotExcelPlugin()

numLista = 18
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
            
            sleep(0.1)

            if not self.find( "cmpTel2", matching=0.97, waiting_time=10000):
                self.not_found("cmpTel2")
            self.click()
            
            if not self.find( "cmpDDD", matching=0.97, waiting_time=10000):
                self.not_found("cmpDDD")
            self.click_at(175,263)
            self.click_at(175,263)
            
            self.control_c()        #copia o DDD
            ddd = self.get_clipboard()  #função que recupera o que foi copiado (CRTL + C) anteriormente
            bot_excel.set_cell('E',column_table,ddd)            
            
            if not self.find( "cmpTelefone", matching=0.97, waiting_time=10000):
                self.not_found("cmpTelefone")
            self.click_at(230,260)
            self.click_at(230,260)
            
            #sleep(30)
            
            self.control_c()        #copia o Telefone
            tel = self.get_clipboard()  #função que recupera o que foi copiado (CRTL + C) anteriormente
            print(str(column_table) + " : " + ddd + " " + tel)
            bot_excel.set_cell('F',column_table,tel)
            column_table += 1
            
            bot_excel.write('botTest\LISTA' + str(numLista) + '.xlsx')
            
            if not self.find( "btnFechar", matching=0.97, waiting_time=10000):
                self.not_found("btnFechar")
            self.click()
            
            if not self.find( "btnAtualizar", matching=0.97, waiting_time=10000):
                self.not_found("btnAtualizar")
            
            sleep(0.5)
    
            self.tab(wait=200)
            self.tab(wait=200)
            self.tab(wait=200)
            self.tab(wait=200)
                
        for i in listaNomes:
            print(str(i) + " ")       
        

    def not_found(self, label):
        print(f"Element not found: {label}")

if __name__ == '__main__':
    Bot.main()







