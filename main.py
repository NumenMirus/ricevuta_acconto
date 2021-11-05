#!/usr/bin/env python

import csv
import click
import os
import time
import docx
from docx import Document
from docx.text.paragraph import Paragraph
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from docx.shared import RGBColor
from docx.shared import Mm
from generator import create_receipt
from datetime import date

path = "/home/numen/Downloads/"

class StoredItem:
    quantity = int()
    id = int()

    def __init__(self, q, id):
        self.quantity = q
        self.id = id

    def getQuantity(self):
        return self.quantity

    def getId(self):
        return self.id

def title():
    print("▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄\n"
          "██ ▄▀▄ █ ▄▄▀█▄ ▄█ ▄▄▀██ ▄▀▄ █ ▄▄▀█▄ ▄█ ▄▄▀██\n"
          "██ █ █ █ ▀▀ ██ ██ ▀▀ ██ █ █ █ ▀▀ ██ ██ ▀▀ ██\n"
          "██ ███ █▄██▄██▄██▄██▄██ ███ █▄██▄██▄██▄██▄██\n"
          "▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀\n")

def getWarehouse():
    items = {}

    with open('data.dat', mode='r') as file:
        f = csv.reader(file)

        for line in f:
            items.update({int(line[0]) : (line[1], float(line[2]))})
    
    return items

def storeItems(items):
    with open("data.dat", mode="w") as file:
        for i in items:
            file.write("%d,%s,%.2f\n" %(i, items[i][0], items[i][1]))
            
def itemInReceipt(item, filename):
    with open(filename, mode="a") as file:
        file.write("%d,%d\n" %(item.getQuantity(), item.getId()))

def addItemToWarehouse(items):
    while True:
        click.clear()
        print("Non scrivere niente per tornare indietro\n")
        name = input("Nome prodotto da aggiungere: ")
        res = None

        if name == "":
            return

        for key in items:
            if items[key][0] == name:
                res = True
        if res:
            print("Prodotto già esistente!")
            time.sleep("2")
        
        else:
            break
        
    price = float(input("\nPrezzo: "))
    
    items.update({len(items)+1 : (name, price)})
    storeItems(items)

def visualizeItemsInWarehouse(items, wait=False):
    print("Prodotti:\n")
    for i in items:
        print("%d) %s" %(i, items[i][0]))
    if wait:
        input("\n\nPremi invio per tornare indietro...")

def addItemToRecipe(items, name):
    print("Prodotti:\n")
    for i in items:
        print("%d) %s" %(i, items[i][0]))
    print("\n\n0) Indietro")
    
    try:
        id = int(input("\nSeleziona prodotto da inserire: "))
        if id == 0:
            pass
        
        if id >= 1 and id <= len(items):
            q = int(input("Quantità: "))

            item = StoredItem(q, id)
            itemInReceipt(item, "./receipts/"+name)
    except ValueError:
        return

def createReceipt(name):
    if name == "":
        return
    try:
        receipt = open("./receipts/"+name+".dat", mode="x")
    except FileExistsError:
        print("\nRicevuta già esistente!")
        time.sleep(2)

def visualizeReceipts():
    l = os.listdir("receipts")
    if not len(l):
        print("\nNon sono presenti ricevute")
    else:
        i = 0
        for r in l:
            i = i+1
            print("%d ..... " %i+ r[:-4] )

    input("\n\nPremi invio per tornare indietro...")

def deleteReceipt(name):
    if os.path.exists("./receipts/"+name):
        os.remove("./receipts/"+name)
        print("\nRicevuta eliminata!")
        time.sleep(2)
    else:
        print("\nLa ricevuta non esiste")
        time.sleep(2)

def deleteItemFromWarehouse(items):
    
    visualizeItemsInWarehouse(items, False)

    print("\n\n0) Indietro")
    
    try:
        id = input("\nSeleziona prodotto da eliminare: ")
        id = int(id)
        if id == 0:
            return
        elif id >= 1 and id <= len(items):
            items.pop(id)
            storeItems(items)
            items = getWarehouse()
    except ValueError:
        return

def selectReceipt():
    l = os.listdir("receipts")
    if not len(l):
        print("\nNon sono presenti ricevute")
    else:
        i = 0
        for r in l:
            i = i+1
            print("%d)" %i+ r[:-4] )
        print("\n\n0) Indietro")
    try:
        res = int(input("\n\nNr. della ricevuta da aprire: "))

        if res == 0:
            return -2
        elif res <= len(l) and res > 0:
            return res
        else:
            return -1

    except ValueError:
        return

def printRecipe(items, file, name):
    f = open("./receipts/"+file, mode="r")
    file = csv.reader(f)
    item_list = []

    for line in file:
        it = []
        it.append(line[0])

        price = int(items[int(line[1])][1])
        item_name = items[int(line[1])][0]

        it.append(price)
        it.append(item_name)
        item_list.append(it)
    
    numero = input("Numero della ricevuta: ")

    causale = input("Causale del pagamento: ")

    
    d = create_receipt(item_list, numero, causale)

    if name == "":
        name = "ricevuta_" + date.today().strftime("%d-%m-%Y")
    d.save(path+name+".docx")

    print("Stampata!")
    time.sleep(2)

def main():
    items = getWarehouse()
    choice1 = None
    choice2 = None
    choice3 = None
    choice4 = None

    while choice1 != 0:
        choice1 = None

        click.clear()
        title()
        print("Azioni disponibili:\n")
        print("1)Gestione Ricevute\n2)Gestione magazzino\n3)Apri ricevuta\n\n0) Esci\n")
        try:
            choice1 = int(input("Selezione: "))
        except ValueError:
            pass

    
        if choice1 == 1:
            choice2 = None
            while choice2 != 0:
                click.clear()
                title()
                print("Gestore ricevute - azioni disponibili:\n")
                print("1) Crea nuova ricevuta\n2) Visualizza ricevute\n3) Elimina ricevuta\n\n0) Indietro\n")
                try:
                    choice2 = int(input("Selezione: "))
                except ValueError:
                    pass

                if choice2 == 1:
                    click.clear()
                    name = input("Nome ricevuta: ")
                    createReceipt(name)
                elif choice2 == 2:
                    click.clear()
                    visualizeReceipts()
                elif choice2 == 3:
                    click.clear()
                    try:
                        name = input("0) Indietro\n\nNome della ricevuta da eliminare: ")
                        if int(name) == 0:
                            continue
                        deleteReceipt(name)
                        break
                    except ValueError:
                        continue
                else:
                    pass

        elif choice1 == 2:
            choice3 = None
            while choice3 != 0:
                click.clear()
                title()
                print("Gestore magazzino - azioni disponibili:\n")
                print("1) Aggiungi prodotto\n2) Elimina prodotto\n3) Visualizza prodotti\n\n0) Indietro\n")
                try:
                    choice3 = int(input("Selezione: "))
                except ValueError:
                    pass

                if choice3 == 1:
                    click.clear()
                    addItemToWarehouse(items)
                    items = getWarehouse()
        
                if choice3 == 2:
                    click.clear()
                    deleteItemFromWarehouse(items)
                    items = getWarehouse()

                if choice3 == 3:
                    click.clear()
                    visualizeItemsInWarehouse(items, True)
                else:
                    pass

        if choice1 == 3:
            choice4 = None
            while choice4 != 0:
                title()
                click.clear()
                print("Seleziona la ricevuta da aprire: \n")

                r = selectReceipt()
                if not isinstance(r, int):
                    continue
                
                if r == -1:
                    click.clear()
                    print("Ricevuta non esistente!")
                    time.sleep(2)
                    pass
                elif r == -2:
                    break

                else:
                    click.clear()
                    l = os.listdir("receipts")

                    recipt = l[r-1]
                    
                    click.clear()
                    print(recipt, end="\n\n")

                    while True:
                        click.clear()
                        try:
                            choice5 = int(input("1) Aggiungi prodotto\n2) Stampa ricevuta\n3) Elimina ricevuta\n\n0) Indietro\n\nSelezione: "))
                        except ValueError:
                            pass

                        if choice5 == 0:
                            break

                        elif choice5 == 1:
                            click.clear()
                            items = getWarehouse()
                            addItemToRecipe(items, recipt)

                        elif choice5 == 2:
                            name = input("Nome per il salvataggio: ")
                            printRecipe(items, recipt, name)
                        elif choice5 == 3:
                            deleteReceipt(recipt)
                            break
                        else:
                            pass
                        

        else:
            pass


main()