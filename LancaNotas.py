'''
Programa para lançamento de notas no sistema IdUFF automatizado 
Desenvolvido por Gabriel Nascimento em 02/03/2022
Última revisão em 01/08/2022

Bibliotecas necessárias:
    - selenium
    - openpyxl
    - pandas 
    - tkinter
'''

# === INPUTS === #
FileOutputSiteData = 'NotasNoSite.xlsx' # arquivo para salvar os valores atuais armazenados no sistema (.xlsx ou .csv)
FileInputGrades = 'NotasParaLancar.xlsx' # arquivo com as notas a serem lançadas (.xlsx ou .csv)
                                              # deve ter uma coluna 'Matricula' (sem acento), com as matrículas, 
                                              # e uma coluna 'Nota', com a nota a ser lançadas
                                              # outras colunas serão ignoradas

nDecimalPlaces = 1 # quantidade de casas decimais
tPauseBetween = 0.1 # tempo, sem segundos, para pausa entre lançamento de notas
ChromeDriverPath = 'chromedriver.exe'   # driver para acesso automatizado ao Google Chrome
                                        # baixe em https://chromedriver.chromium.org/, 
                                        # conforme a versão do Chrome instalado e o sistema operacional
HtmlAdress = 'https://app.uff.br/graduacao/lancamentodenotas'   # endereço do sistema de lançamento de notas
NumberOfStudents_Xpath = '/html/body/div/div[2]/div[2]/div/div[1]/div[1]/div/div[2]/div/div[1]/h1'

# === LIBS === #
from selenium import webdriver #
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd
from math import ceil

# === AUX === #
def InitiateBrowser():
    driver = webdriver.Chrome(ChromeDriverPath)
    driver.get(HtmlAdress)
    return driver

def GetGrade(GradeList, id):
    selecao = GradeList.loc[GradeList['Matricula'] == id]
    if len(selecao) > 1:
        print(f'\n\nERRO! Mais de uma linha encontrada no arquivo de notas com a mesma matricula ({id})')
        return -1
    if len(selecao) == 0:
        print(f'\n\nERRO! Não foi encontrada uma linha no aruqivo de notas para matrícula {id}.')
        return -1
    grade = selecao.iloc[0]['Nota']
    grade = ceil(grade*10**nDecimalPlaces)/(10**nDecimalPlaces) #round(grade, nDecimalPlaces)
    return grade

def GetListFromSite(driver):
    global StudentList, nStudents    

    try:        
        nStudents = int(driver.find_element(By.XPATH, NumberOfStudents_Xpath).text)        
    except:
        answer = input('Erro ao obter a quantidade de alunos. Deseja entrar manualmente (S ou N)? ')     
        if answer.upper() == 'S':
            answer = input('Digite a quantidade total de alunos: ')
            nStudents = int(answer)
        else:
            return 'N'       

    print(f"\n{nStudents} alunos")

    list = []
    for i in range(nStudents):
        xpath = f'/html/body/div/div[2]/div[2]/div/div[2]/form/table/tbody/tr[{i+1}]/td[1]'
        StudentFullDescription = driver.find_element(By.XPATH, xpath).text
        mat = int(StudentFullDescription.split()[2])
        name = StudentFullDescription.split()[4] + ' ' + StudentFullDescription.split()[5]
        try:
            xpath = f'/html/body/div/div[2]/div[2]/div/div[2]/form/table/tbody/tr[{i+1}]/td[2]/div/input'
            element = driver.find_element(By.XPATH, xpath)
            grade = element.get_attribute('value')
        except:
            grade = '-'
        list.append([mat, name, grade])

    StudentList = pd.DataFrame(list, columns = ['Matricula', 'Nome', 'Nota Site'])
    StudentList.to_excel(FileOutputSiteData)

    print("\n\n\n\n========================================")
    print("Lista de alunos obtidas do IdUFF:\n")
    print(StudentList)    
    answer = input('\nContinuar (S ou N)? ')
    return answer.upper()


def GetNewGrades():    
    StudentList['Nova Nota'] = pd.NaT # add a column
    NewGradeList = pd.read_excel(FileInputGrades)
    for i, row in StudentList.iterrows():
        mat = row['Matricula']
        newgrade = GetGrade(NewGradeList, mat)
        StudentList.at[i, 'Nova Nota'] = newgrade
    print("\n\n\n\n========================================")
    print('Lista atualizada:')
    print(StudentList)
    answer = input('\nObs.: os alunos com problemas (nota -1) não serão lançados.\n'+
                   '\nLançar as notas (S ou N)? ')
    return answer.upper()


def WriteGrades(driver):
    for i in range(nStudents):
        id = StudentList.iloc[i]['Matricula']
        name = StudentList.iloc[i]['Nome']
        grade = StudentList.iloc[i]['Nova Nota']
        gradetxt = str(grade)
        if grade >= 0:
            try:
                xpath = f'/html/body/div/div[2]/div[2]/div/div[2]/form/table/tbody/tr[{i+1}]/td[2]/div/input'
                element = driver.find_element(By.XPATH, xpath)
                element.clear()
                element.send_keys(gradetxt)
            except:
                print(f'Erro ao lançar a nota ({grade}) do aluno {id} - {name}.')
            else:
                print(f'Lançada nota {grade} do aluno {id} - {name}.')
        
        sleep(tPauseBetween)


# === MAIN === #
driver = InitiateBrowser()
input("\n\n\nEntre com o login e senha e selecione a turma para lancamento de notas. Então, pressione enter.\n")

go = GetListFromSite(driver)
if go == 'S': go = GetNewGrades()
if go == 'S': WriteGrades(driver)

input("\n\nPressione enter para finalizar.\n\n")
driver.quit()
