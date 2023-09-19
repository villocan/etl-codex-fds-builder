#st-----Imports-------
import pandas as pd
from docxtpl import DocxTemplate
#-------Imports-----nd

#st-----Definitions-------
#-------Definitions-----nd

#st-----Configurations-------
Template_Path = r"C:\Users\ve042\OneDrive\github\internship-letter-buildup-tool\Formato\letter_template.docx"
nombre_decano = "Dr. Cezar G"
cedula_decano = "V-1"
#-------Configurations-----nd

#st-----Main-------
#Definition
def main():
    data_alumnos = pd.read_csv("data.csv")
    for k in range(0,len(data_alumnos)):
        #data
        nombre_bachiller = data_alumnos['Primer Nombre'][k]
        cedula_bachiller = "V-20072347"
        nombre_empresa = "Brainly Marketing"
        fecha_inicio = "05/Nov/1990"
        fecha_fin = "13/Ene/2023"
        horario_dias = "De Lunes a Viernes"
        horario_bloque1 = "Desde 08:00am hasta 12:00m"
        horario_bloque2 = "Desde 01:00pm hasta 05:00pm"
        horas_cursadas = "1000"
        nombre_tutor_industrial = "Cesar Villalobos"
        nombre_tutor_academico = "Gladys Quevedo"
        calificacion = "VEINTE (20)"
        Instance_Path = r"C:\Users\ve042\OneDrive\github\internship-letter-buildup-tool\Planillas\Planilla Pasantia "+nombre_bachiller+" "+cedula_bachiller+".docx"
        if horario_bloque2!="":
            horario_bloque2=" y "+horario_bloque2
        #reporting
        Instance_Letter = DocxTemplate(Template_Path)
        context = {
            'nombre_decano': nombre_decano,
            'cedula_decano': cedula_decano,
            'nombre_bachiller': nombre_bachiller,
            'cedula_bachiller': cedula_bachiller,
            'nombre_empresa': nombre_empresa,
            'fecha_inicio': fecha_inicio,
            'fecha_fin': fecha_fin,
            'horario_dias': horario_dias,
            'horario_bloque1': horario_bloque1,
            'horario_bloque2': horario_bloque2,
            'horas_cursadas': horas_cursadas,
            'nombre_tutor_industrial': nombre_tutor_industrial,
            'nombre_tutor_academico': nombre_tutor_academico,
            'calificacion': calificacion
        }
        Instance_Letter.render(context)
        Instance_Letter.save(Instance_Path)
#execute
if __name__ == '__main__':
    main()
#-------Main-----nd