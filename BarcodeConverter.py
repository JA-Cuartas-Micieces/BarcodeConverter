# -*- coding: utf-8 -*-
"""
Created on Wed Mar 17 20:47:46 2021

MIT License

Copyright (c) 2021 Javier Alejandro Cuartas Micieces
https://github.com/JA-Cuartas-Micieces

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

"""

class BarcodeConverter:
    def __init__(self,*args,**kwargs):
        self.root=Tk()
        self.root.withdraw()
        self.cd = os.path.realpath(__file__)[:-19]
        try:
            ksd=sorted([el for el in os.listdir() if el.endswith(".xlsx")])
            rn=self.cd+ksd[0]
            inputWorkbook=xlrd.open_workbook(rn, on_demand = True)
            bk=inputWorkbook.sheet_by_index(0)
            columnsdf=list()
            for il in range(0,1):
                for ik in range(0,bk.ncols):
                    val=bk.cell(il, ik).value 
                    columnsdf.append(val)
            d1=pd.DataFrame(columns=columnsdf)
            for il in range(1,bk.nrows):
                for ik in range(0,bk.ncols):
                    if type(bk.cell(il, ik).value)==str: 
                        d1.loc[il-1,columnsdf[ik]]=bk.cell(il, ik).value.strip()
                    else:
                        d1.loc[il-1,columnsdf[ik]]=bk.cell(il, ik).value
            inputWorkbook.release_resources()
            
            os.chdir(self.cd+"Files\\")
            skipf=int(d1.loc[0,columnsdf[0]])
            skipll=int(d1.loc[0,columnsdf[len(columnsdf)-1]])
            filechain=[fel for fel in os.listdir() if fel.lower().endswith(".csv")]
            for el in filechain:
                with open(el) as fp:
                    Lines = fp.readlines()
                    Linesf=list()
                    subscol=d1[columnsdf[1]].unique().tolist()
                    for il,line in enumerate(Lines):
                        lf=line.split(",")
                        l=[bel.replace("\"","") for bel in lf]
                        if il<=skipf-1:
                            Linesf.append(line)
                        elif all([il>skipf-1,il<len(Lines)-skipll]):
                            rt=0
                            for gdel in subscol:
                                if l[int(gdel)-1] in d1[d1[columnsdf[1]]==int(gdel)][columnsdf[2]].values.tolist():
                                    rt=1
                                    l[int(gdel)-1]="\""+d1[(d1[columnsdf[1]]==int(gdel))&(d1[columnsdf[2]]==l[int(gdel)-1])][columnsdf[3]].values[0]+"\""
                                    Linesf.append(self.record_strings(l,lf))
                                    break
                            if rt==0:
                                Linesf.append(line)
                        elif il>=len(Lines)-skipll:
                            Linesf.append(line)
                        
                    with open("converted_"+el[:-4]+".CSV", "w") as fw:
                        fw.writelines(Linesf)
        except:
            self.error_W()
        self.root.destroy()
        
    def record_strings(self,*args):
        lg=[bel.replace("\"","") for bel in args[0]]
        for iel,bel in enumerate(lg):
            try:
                float(bel)
            except:
                if "\n" in bel:
                    bel="\""+bel[:-1]+"\""+bel[-1:]
                elif all([args[1][iel]=='',args[1][iel]!="\"\""]):
                    bel=args[0][iel]
                else:
                    bel="\""+bel+"\""
            lg[iel]=bel
        resl=",".join(lg)
        return resl

    def error_W(self):
        self.top = tk.Toplevel()
        self.top.tk.call('wm', 'iconphoto', self.top._w,tk.PhotoImage(file=self.cd+"ICON.png"))
        self.top.resizable(0,0)
        self.top.title('Error: Entrada Incorrecta')
        msg=tk.Message(self.top,text="La estructura de la tabla de entrada .xlsx no es la esperada por el programa, no hay correspondencia de las columnas especificadas en el .xlsx con las del archivo de entrada que debe ser .CSV, las última y primera columna del .xlsx no tienen todos los valores iguales, o algún otro elemento está incorrectamente definido.",width=500)
        msg.grid(padx=0,pady=0)
        self.top.mainloop()
	
    def error_importing(self):
        self.top = tk.Toplevel()
        self.top.tk.call('wm', 'iconphoto', self.top._w,tk.PhotoImage(file=self.cd+"ICON.png"))
        self.top.resizable(0,0)
        self.top.title('Error de importación')
        msg=tk.Message(self.top,text="Por favor, instale las librerías de python pandas, tkinter y xlrd para que el script pueda funcionar correctamente.",width=500)
        msg.grid(padx=0,pady=0)
        self.top.mainloop()
try:
    import os
    import tkinter as tk
    from tkinter import Tk
    import xlrd
    import pandas as pd
except:
    tr=BarcodeConverter("").error_importing()

if __name__=="__main__":
    app=BarcodeConverter()

