import configparser

#TODO:modificare applicazione principale integrando la lettura del file ini,
#aggiungere --help per illustrare come compilare il file ini
#e aggiungere tool --makeconf per creazione guidata file ini
#o a creazone file ini vuoto

def main():
    config = configparser.ConfigParser()
    config.read("conf.ini")
    print(config.sections())
    print("file name:", config["EXCEL"]["FILE_NAME"])

    FILE_NAME = config["EXCEL"]["file_name"]
    XLS_SHEET = config["EXCEL"]["sheet_number"]
    XLS_FIRST_ROW = config["EXCEL"]["first_useful_row"]
    XLS_DATE = config["EXCEL"]["date_column"]
    XLS_WORKER = config["EXCEL"]["worker_column"]

    SHIFT_DURATION = 9
    CALENDAR = config["CALENDAR"]["calendar_name"]
    DAYS_TO_READ = config["CALENDAR"]["days_to_read"]

    DEBUG_EXCEL_ONLY = config["DEBUG"]["debug_excel_only"]
    DEBUG_READ_ONLY = config["DEBUG"]["debug_read_cal_only"]
    

if __name__ == '__main__':
    main()
