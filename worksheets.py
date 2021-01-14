import logging
import sys
import re
import math
import gspread
from oauth2client.service_account import ServiceAccountCredentials

#Function to get the student status in a semester
def grades_and_absences(worksheet_list, total_perc, enrol, absc, p1, p2, p3):
    answer = []

    for student in worksheet_list:
        logging.info("Checking student status: enrollment number {}".format(student[enrol.col-1]))
        try:
            grade    = math.ceil((int(student[p1]) + int(student[p2]) + int(student[p3])) / 3) #Round the grade UP 
            absences = int(student[absc])  #Get the absences number
        except ValueError:
                logging.error("String to Int conversion")
                return []

        #Verify the student situation, and insert into the list
        if absences < total_perc:
            answer.append(["Reprovado por Falta",0])
        elif grade >= 70 :
            answer.append(["Aprovado",0])
        elif grade >= 50:
            answer.append(["Exame Final",(100 - grade)])
        else:
            answer.append(["Reprovado por Nota",0])
    
    return answer

#Function to get the classes total number in a semester
def get_class_total(worksheet_list):
    time_total = -1
    position   = 0

    logging.info("Finding the Number of Classes in a semester")
    while (position < len(worksheet_list)) and (time_total == -1):
        if [text for text in worksheet_list[position] if text.startswith('Total de aulas no semestre: ')]: #Get the number of classes total in a semester
            try:
                time_total = int(re.findall("\d+", worksheet_list[position][0])[0])
            except ValueError:
                logging.error("String to Int conversion")
                return 0
        position += 1

    if time_total == -1:
        logging.error("No number of classes specified")
        return 0
    
    return time_total
         

def main(scope_f = None, sheet_name_f = None):
    #Verify if is missing data to obtain the worksheet
    if scope_f is None:
        logging.error("No scope specified")
        return False
    elif sheet_name_f is None:
        logging.error("No results specified")
        return False
    
    logging.info("Loading Sheet")

    credentials_f = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope_f)
    gc_f = gspread.authorize(credentials_f)
    ss_f = gc_f.open(sheet_name_f)

    try:
        worksheet_f = ss_f.sheet1
    except gspread.exceptions.WorksheetNotFound:
        logging.error("Worksheet not found")
        return False
    
    logging.info("Sheet Loaded")
    list_f = worksheet_f.get_all_values() #Get a list with all worksheet values
    
    #Get the number of classes 
    time_total_f = get_class_total(list_f) 
    if time_total_f == 0:
        return False

    #Get the cells coordinates
    p1_position         = worksheet_f.find("P1")
    p2_position         = worksheet_f.find("P2")
    p3_position         = worksheet_f.find("P3")
    enrol_position      = worksheet_f.find("Matricula")
    abscences_position  = worksheet_f.find("Faltas")
    situation_position  = worksheet_f.find("Situação")
    naf_position        = worksheet_f.find("Nota para Aprovação Final")

    #Get a list with the student situation
    list_f = grades_and_absences(list_f[abscences_position.row:],
                                 time_total_f*(0.25), #Get the answer list
                                 enrol_position,
                                 abscences_position.col-1,
                                 p1_position.col-1,
                                 p2_position.col-1,
                                 p3_position.col-1)

    #if the returned list is Empty
    if not list_f:
        return False
    
    #Writing the updated data in the worksheet
    for position in range(len(list_f)):
        logging.info("Writing data in the datasheet: enrollment number {}".format(position+1))
        worksheet_f.update_cell(position+situation_position.row+1,
                                situation_position.col, 
                                list_f[position][0])
        worksheet_f.update_cell(position+naf_position.row+1,
                                naf_position.col,
                                list_f[position][1])

    return True

log_format = ('[%(asctime)s] %(levelname)-8s %(name)-12s %(message)s')
logfile = 'aplication.log'

logging.basicConfig(level=logging.INFO, 
                        format=log_format, 
                        # Declare handlers
                        handlers=[
                            logging.FileHandler(logfile),
                            logging.StreamHandler(sys.stdout),
                        ]
                   )

scope       = ['https://www.googleapis.com/auth/drive']
sheet_name  = 'Cópia de Engenharia de Software - Desafio [Sammuel Ramos]'

if main(scope, sheet_name):
    logging.info("Closing Application: SUCESS")
