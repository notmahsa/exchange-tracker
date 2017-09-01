import pygame as pg
import pandas as pd
import sys, os
from pygame.locals import *
from openpyxl import load_workbook
import xlsxwriter
from time import gmtime, strftime
import pytz
from datetime import datetime
from Tkinter import *
import numpy

pg.init()

class Checkbox:
    def __init__(self, surface, x, y, caption, font_color, color=(230, 230, 230), outline_color=(0, 0, 0),check_color=(0, 0, 0), font_size=22, text_offset=(28, 1)):
        self.surface = surface
        self.x = x
        self.y = y
        self.color = color
        self.caption = caption
        self.oc = outline_color
        self.cc = check_color
        self.fs = font_size
        self.fc = font_color
        self.to = text_offset
        # checkbox object
        self.checkbox_obj = pg.Rect(self.x, self.y, 12, 12)
        self.checkbox_outline = self.checkbox_obj.copy()
        # variables to test the different states of the checkbox
        self.checked = False
        self.active = False
        self.unchecked = True
        self.click = False

    def _draw_button_text(self):
        self.font = pg.font.Font(None, self.fs)
        self.font_surf = self.font.render(self.caption, True, self.fc)
        w, h = self.font.size(self.caption)
        self.font_pos = (self.x + 12 / 2 - w / 2 + self.to[0], self.y + 12 / 2 - h / 2 + self.to[1])
        self.surface.blit(self.font_surf, self.font_pos)

    def render_checkbox(self):
        if self.checked:
            pg.draw.rect(self.surface, self.color, self.checkbox_obj)
            pg.draw.rect(self.surface, self.oc, self.checkbox_outline, 1)
            pg.draw.circle(self.surface, self.cc, (self.x + 6, self.y + 6), 4)

        elif self.unchecked:
            pg.draw.rect(self.surface, self.color, self.checkbox_obj)
            pg.draw.rect(self.surface, self.oc, self.checkbox_outline, 1)
        self._draw_button_text()

    def _update(self, event_object):
        x, y = event_object.pos
        # self.x, self.y, 12, 12
        px, py, w, h = self.checkbox_obj  # getting check box dimensions
        if px < x < px + w and py < y < py + h:
            self.active = True
        else:
            self.active = False

    def _mouse_up(self):
        if self.active and not self.checked and self.click:
            self.checked = True
        elif self.active and self.checked:
            self.checked = False
            self.unchecked = True

        if self.click is True and self.active is False:
            if self.checked:
                self.checked = True
            if self.unchecked:
                self.unchecked = True
            self.active = False

    def update_checkbox(self, event_object):
        if event_object.type == pg.MOUSEBUTTONDOWN:
            self.click = True
            # self._mouse_down()
        if event_object.type == pg.MOUSEBUTTONUP:
            self._mouse_up()
        if event_object.type == pg.MOUSEMOTION:
            self._update(event_object)

    def is_checked(self):
        if self.checked is True:
            return True
        else:
            return False

    def is_unchecked(self):
        if self.checked is False:
            return True
        else:
            return False

    def uncheck(self):
        self.checked = False
        self.unchecked = True
def main():
    gmt = pytz.timezone('GMT')
    eastern = pytz.timezone('US/Eastern')
    time = strftime('%a, %d %b %Y %H:%M:%S GMT', gmtime())

    date = datetime.strptime(time, '%a, %d %b %Y %H:%M:%S GMT')
    
    dategmt = gmt.localize(date)
   
    dateeastern = dategmt.astimezone(eastern)
    
    time = str(dateeastern)
    day = time[:10]
    try:
        xl = pd.ExcelFile(day+".xlsx")
    except:
        writer = pd.ExcelWriter(day+'.xlsx', engine='xlsxwriter')
        
        df = pd.DataFrame(columns=["Time","CAD","USD","EUR"])
        df.to_excel(writer, startrow=0, index=False)
        df = pd.DataFrame(columns=["Current Amount","0","0","0"])
        df.to_excel(writer, startrow=1, index=False)
        df = pd.DataFrame(columns=["","","",""])
        df.to_excel(writer, startrow=2, index=False)
                         
        writer.save()
    pg.display.set_caption('Ashena Arz')
    WIDTH = 520
    HEIGHT = 320
    display = pg.display.set_mode((WIDTH, HEIGHT))

    basicfont = pg.font.SysFont("FreeSansBold.ttf", 22)
    tTypeText = basicfont.render('Transaction Type:', True, (0, 0, 0))
    currencyText = basicfont.render('Currency:', True, (0, 0, 0))
    amountText = basicfont.render('Amount:', True, (0, 0, 0))
    saveText = basicfont.render('Save', True, (0, 0, 0))
    
    typeTran = []
    typeTran.append(Checkbox(display, 150,80, "BUY", (0,0,0)))
    typeTran.append(Checkbox(display, 150,100, "SELL", (0,0,0)))
    typeTranLabel = ["Buy", "Sell"]

    curr = []
    curr.append(Checkbox(display, 310,80, "CAD", (0,0,200)))
    curr.append(Checkbox(display, 310,100, "USD", (200,0,0)))
    curr.append(Checkbox(display, 310,120, "EUR", (0,99,33)))
    currLabel = ["CAD", "USD", "EUR"]

    
    key = ""
    message = ""
    running = True
    timeMsg = ""
    while running:
        
        for event in pg.event.get():
            
            mouse = pg.mouse.get_pos()
            if event.type == pg.QUIT:
                running = False
                pg.quit()
                quit()
            if event.type == KEYDOWN and event.key >= K_0 and event.key <= K_9 and len(key) < 10:
                key += str(event.key - K_0)
            if event.type == KEYDOWN and event.key == K_BACKSPACE:
                key = key[:-1]
            if event.type == pg.MOUSEBUTTONDOWN and 310<mouse[0] <360 and 185<mouse[1]<205:
                color = (200, 200, 200);
            else:
                color = (0,0,0)
            if event.type == pg.MOUSEBUTTONUP and 310<mouse[0] <360 and 185<mouse[1]<205:
                currList = []
                typeList = []

                gmt = pytz.timezone('GMT')
                eastern = pytz.timezone('US/Eastern')
                time = strftime('%a, %d %b %Y %H:%M:%S GMT', gmtime())

                date = datetime.strptime(time, '%a, %d %b %Y %H:%M:%S GMT')
                
                dategmt = gmt.localize(date)
               
                dateeastern = dategmt.astimezone(eastern)
                
                time = str(dateeastern)
                day = time[:10]
                
                for i in typeTran:
                    if i.is_checked():
                        typeList.append(i.caption)
                for i in curr:
                    if i.is_checked():
                        currList.append(i.caption)
                
                try:
                    xl = pd.ExcelFile(day+".xlsx")
                except:
                    writer = pd.ExcelWriter(day+'.xlsx', engine='xlsxwriter')
                    
                    df = pd.DataFrame(columns=["Time","CAD","USD","EUR"])
                    df.to_excel(writer, startrow=0, index=False)
                    df = pd.DataFrame(columns=["Current Amount","0","0","0"])
                    df.to_excel(writer, startrow=1, index=False)
                    df = pd.DataFrame(columns=["","","",""])
                    df.to_excel(writer, startrow=2, index=False)
                                     
                    writer.save()

                xl = pd.ExcelFile(day+".xlsx")
                writer = pd.ExcelWriter(day+'.xlsx', engine='openpyxl')
                book = load_workbook(day+'.xlsx')
                
                writer.book = book
                writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                worksheet0 = book.worksheets[0]
                xldf = xl.parse(xl.sheet_names[0])
                rowStart = len(xldf)+1

                df = pd.DataFrame(columns =[""])
                message = ""
                timeMsg = ""
                try:
                    message = "Saved: "+typeList[0]+" "+currList[0]+" "+str(key)
                    timeMsg = "at: "+time
                    if typeList[0] == "BUY":
                        if currList[0] == "CAD":
                            df = pd.DataFrame(columns=[time,float(key),0, 0])
                        elif currList[0] == "USD":
                            df = pd.DataFrame(columns=[time,0, float(key), 0])
                        elif currList[0] == "EUR":
                            df = pd.DataFrame(columns=[time,0,0, float(key)])
                    else:
                        key = -1 * float(key)
                        if currList[0] == "CAD":
                            df = pd.DataFrame(columns=[time,float(key),0, 0])
                        elif currList[0] == "USD":
                            df = pd.DataFrame(columns=[time,0, float(key), 0])
                        elif currList[0] == "EUR":
                            df = pd.DataFrame(columns=[time,0,0, float(key)])
                except:
                    message = "ERROR! Please enter all entries"
                sumRow = rowStart
                if sumRow < 3:
                    sumRow = 3
                dfSum = pd.DataFrame(columns=["=SUM(B2:B"+str(sumRow)+")", "=SUM(C2:C"+str(sumRow)+")","=SUM(D2:D"+str(sumRow)+")"])
                
                
                
                if rowStart < 4:
                    rowStart = 4
                    
                try:
                    df.to_excel(writer, startrow=rowStart-1, index=False)
                    dfSum.to_excel(writer, startrow=rowStart, index_label='TOTAL')

                except:
                    print("LOL you dun messed up")

                # Close the Pandas Excel writer and output the Excel file.
                writer.save()
                for i in typeTran:
                    i.uncheck()
                for i in curr:
                    i.uncheck()
                key = ""
                            
            for i in typeTran:
                i.update_checkbox(event)
            for i in curr:
                i.update_checkbox(event)
        
        
        display.fill((200,200, 200))
        msgText = basicfont.render(message, True, (0, 0, 0))
        pg.draw.rect(display, color, pg.Rect(147, 182, 99,26))
        pg.draw.rect(display, (255,255,255), pg.Rect(150, 185, 93,20))
        pg.draw.rect(display, color, pg.Rect(310, 185, 50,20))
        pg.draw.rect(display, (0,255,0), pg.Rect(312, 187, 46,16))
        amountNum = basicfont.render(key, True, (0,0, 0))
        timeMsgText = basicfont.render(timeMsg, True, (0,0, 0))
        display.blit(tTypeText, (150, 55))
        display.blit(currencyText, (310, 55))
        display.blit(amountText, (150, 160))
        display.blit(amountNum, (153, 188))
        display.blit(saveText, (317, 188))
        display.blit(msgText, (150, 240))
        display.blit(timeMsgText, (150, 260))
        
        for i in typeTran:
            i.render_checkbox()
        for i in curr:
            i.render_checkbox()
            
        pg.display.flip()
    pg.display.update()
if __name__ == '__main__':
    main()
