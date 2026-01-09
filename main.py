import json
from datetime import datetime
from InquirerPy import inquirer
import os
import xlsxwriter

def setupHeaders(worksheet,cell_format):
        worksheet.write(0,0,"#")

        worksheet.write(0,1,"Gennaio",cell_format)
        worksheet.write(0,2,"")

        worksheet.write(0,3,"Febbraio",cell_format)
        worksheet.write(0,4,"")

        worksheet.write(0,5,"Marzo",cell_format)
        worksheet.write(0,6,"")

        worksheet.write(0,7,"Aprile",cell_format)
        worksheet.write(0,8,"")

        worksheet.write(0,9,"Maggio",cell_format)
        worksheet.write(0,10,"")

        worksheet.write(0,11,"Giugno",cell_format)
        worksheet.write(0,12,"")

        worksheet.write(0,13,"Luglio",cell_format)
        worksheet.write(0,14,"")

        worksheet.write(0,15,"Agosto",cell_format)
        worksheet.write(0,16,"")

        worksheet.write(0,17,"Settembre",cell_format)
        worksheet.write(0,18,"")

        worksheet.write(0,19,"Ottobre",cell_format)
        worksheet.write(0,20,"")

        worksheet.write(0,21,"Novembre",cell_format)
        worksheet.write(0,22,"")

        worksheet.write(0,23,"Dicembre",cell_format)

        for i in range(1,32):
            worksheet.write(i,0,str(i),cell_format)

def fillXlsData(anno, worksheet):
    
    with open("anni/"+anno+".json", "r", encoding="utf-8") as json_file:
        data = json.load(json_file)

    with open('category.json', encoding='utf-8') as json_file:
        cat = json.load(json_file)

    CATEGORIES = cat["CATEGORIES"]
    totAnno = {cat: 0 for cat in CATEGORIES}

    for date in data:
        day = int(date.split("-")[0])
        month = int(date.split("-")[1])
        monthIndex = month*2-1
        text = "";
        for category in data[date]:
            totAnno[category] += data[date][category]
            text += str(data[date][category]) + " in " + category + "\n"
        
        worksheet.write(day,monthIndex," "+text)

    worksheet.write(33,1,"TOTALE ANNUO PER CATEGORIA:")
    XCat = 34;
    for c in totAnno:
        worksheet.write(XCat,1,c)
        worksheet.write(XCat,2,totAnno[c])
        XCat+=1;
        
def xlsRecap():
    current_day = str(datetime.now().day)
    current_month = str(datetime.now().month)
    current_year = str(datetime.now().year)
    workbook = xlsxwriter.Workbook("yearly_recap_"+current_day+"-"+current_month+"-"+current_year+".xlsx")
    anni = [f.replace(".json","") for f in os.listdir("./anni") if f.endswith(".json")]
    cell_format = workbook.add_format({'bold': True})

    for anno in anni:
        worksheet = workbook.add_worksheet(anno)
        setupHeaders(worksheet,cell_format)
        fillXlsData(anno, worksheet)

    workbook.close()
    print("\nReport excel generato")

def addPayment():
    with open('category.json', encoding='utf-8') as json_file:
        data = json.load(json_file)

    CATEGORIES = data["CATEGORIES"]

    if len(CATEGORIES) == 0:
        print("\nNon ci sono categorie disponibili, aggiungi una categoria prima di procedere")
        return
    
    value = input("\nInserisci il valore:")
    if (value.isdigit):
        value = float(value)
    else:
        print("Non è un numero")
        return
    current_day = str(datetime.now().day)
    current_month = str(datetime.now().month)
    current_year = str(datetime.now().year)


    
    category = inquirer.select(
    message="Scegli la categoria:",
    choices=CATEGORIES
    ).execute()


    with open("anni/"+current_year+".json", "r", encoding="utf-8") as json_file:
        data = json.load(json_file)

    if current_day+"-"+current_month in data:
        category_dict = data[current_day+"-"+current_month]
    else:
        category_dict = {}

    if category not in category_dict:
        category_dict[category] = 0
    category_dict[category] += value
    
    data[current_day+"-"+current_month] = category_dict

    with open("anni/"+current_year+".json", "w", encoding="utf-8") as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=4)
    return

def addCategory():
    value = input("\nScrivi il nome della nuova categoria: ")
    with open('category.json','r', encoding='utf-8') as json_file:
        data = json.load(json_file)
    data["CATEGORIES"].append(value.upper())

    with open("category.json", "w", encoding="utf-8") as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=4)
    return

def closeProgram():
    return True;

def main():
    os.makedirs("anni", exist_ok=True)
    anno_corrente = str(datetime.now().year)
    nome_file =  os.path.join("anni", anno_corrente +".json")

    if not os.path.exists(nome_file):
        with open(nome_file, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=4)
     
    if not os.path.exists("category.json"):
        CATEGORIES = {"CATEGORIES" :[]}
        with open("category.json", "w", encoding="utf-8") as f:
            json.dump(CATEGORIES, f, ensure_ascii=False, indent=4)
     
    scelta = inquirer.select(
    message="Scegli un'opzione:",
    choices=["Aggiungi spesa o guadagno",
            "Aggiungi categoria",
            "Vedi spese",
            "Esci"
            ],
    ).execute()

    funzioni = {"Aggiungi spesa o guadagno": addPayment,"Aggiungi categoria": addCategory,"Vedi spese": xlsRecap,"Esci": closeProgram}

    if (funzioni[scelta]()):
        return
    main()

if __name__ == "__main__":
    main() 