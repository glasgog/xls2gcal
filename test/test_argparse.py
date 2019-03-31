import argparse

parser = argparse.ArgumentParser(description="xls2gcal permette di popolare " + \
        "automaticamente il proprio Google Calendar " + \
        "con i turni forniti in un file excel situato in un dato URL. " + \
        "Per far questo necessita che sia presente un file di configurazione ini " + \
        "ed un file contenente l'URL del file excel. " + \
        "Per eseguire la configurazione guidata del file ini " + \
        "o per informazioni riguardo la configurazione manuale eseguire: " + \
        "python xls2gcal.py --conf")
parser.add_argument("--conf", help="Configure", action="store_true")
args = parser.parse_args()

print("L'argomento conf è presente:", args.conf)

if args.conf:
    print("Configurazione in corso.")
    answer = input("Scegliere y per effettuare la configurazione guidata," + \
        "o n per indicazioni sulla configurazione manuale [y/n]: ")
    if answer == "y":
        print("Creazione file ini in corso.")
        print("Creazione file URL xls in corso")
    elif answer == "n":
        print("\nIstruzioni confugurazione manuale file ini.")
        print("Il file conf.ini dovrà essere situato nella cartella data, " + \
            "e dovrà avere la seguente struttura (tab esclusi):")
        print("")
        print("    [EXCEL]")
        print("    sheet_number = {numero foglio xls in cui sono presenti i turni}")
        print("    first_useful_row = {numero riga in cui iniziano le date}")
        print("    date_column = {lettera colonna delle date}")
        print("    worker_column = {lettera colonna dei turni del lavoratore}")
        print("    ")
        print("    [CALENDAR]")
        print("    calendar_name = {nome del Google Calendar}")
        print("    days_to_read = {giorni da leggere dal calendario, consigliati 60}")
        print("\nIstruzioni configurazione manuale file URL xls.")
        print("Il file xls.url dovrà essere situato nella cartella data, " + \
            "e dovrà contenere esclusivamente l'URL del file excel.")
        print("\nNella cartella data/example sono presenti esempi della struttura " + \
            "dei file excel, ini e url")
