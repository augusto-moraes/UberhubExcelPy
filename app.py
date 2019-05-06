import xlsxwriter, xlrd, openpyxl, os
import pandas as pd
import numpy as np

print("Tenha certeza de que a lista com todos os alunos se encontra na pasta [in] !!")

#variaveis
local   = 'Qual sua preferência de local de aula?'
periodo = 'Qual sua preferência de horário aos Sábados?'
nivel   = 'Selecione a melhor opção sobre seu conhecimento em lógica de programação:'
age     = 'Qual é a sua idade?'
name    = 'Nome Completo'
uri     = 'Caso tenha usuário no URI, favor informar o ID.'

#Read in and out
file_a  = input("Nome da lista [in]: ")
file_b  = input("Nome da lista [out]: ")

#Converts Excel (in) to DataFrame
try:
    df1 = pd.read_excel("in/" + file_a)
except:
    df1 = pd.read_excel("in/" + file_a + ".xls")


#local
pit = df1[df1[local] == 'PITÁGORAS']
ufu = df1[df1[local] == 'UFU']
uni = df1[df1[local] == 'UNIUBE']

#periodo
tarde_pit = pit[pit[periodo] == 'Tarde - 13h30 às 16h30']
tarde_ufu = ufu[ufu[periodo] == 'Tarde - 13h30 às 16h30']
tarde_uni = uni[uni[periodo] == 'Tarde - 13h30 às 16h30']

manha_pit = df1[df1[periodo] == 'Manhã - 8h30 às 11h30']
manha_ufu = ufu[ufu[periodo] == 'Manhã - 8h30 às 11h30']
manha_uni = uni[uni[periodo] == 'Manhã - 8h30 às 11h30']

#nivel tarde
iniciante1_tarde_pit = tarde_pit[tarde_pit[nivel] == 'Não sei nada, mas quero aprender!']
iniciante1_tarde_ufu = tarde_ufu[tarde_ufu[nivel] == 'Não sei nada, mas quero aprender!']
iniciante1_tarde_uni = tarde_uni[tarde_uni[nivel] == 'Não sei nada, mas quero aprender!']

iniciante2_tarde_pit = tarde_pit[tarde_pit[nivel] == 'Sei o que é algoritmos, printf e scanf']
iniciante2_tarde_ufu = tarde_ufu[tarde_ufu[nivel] == 'Sei o que é algoritmos, printf e scanf']
iniciante2_tarde_uni = tarde_uni[tarde_uni[nivel] == 'Sei o que é algoritmos, printf e scanf']

intermediario_tarde_pit = tarde_pit[tarde_pit[nivel] == 'Sei o que é string, vetor e matriz']
intermediario_tarde_ufu = tarde_ufu[tarde_ufu[nivel] == 'Sei o que é string, vetor e matriz']
intermediario_tarde_uni = tarde_uni[tarde_uni[nivel] == 'Sei o que é string, vetor e matriz']

avancado_tarde_pit = tarde_pit[tarde_pit[nivel] == 'Sei o que é ordenação']
avancado_tarde_ufu = tarde_ufu[tarde_ufu[nivel] == 'Sei o que é ordenação']
avancado_tarde_uni = tarde_uni[tarde_uni[nivel] == 'Sei o que é ordenação']

#nivel manha
iniciante1_manha_pit = manha_pit[manha_pit[nivel] == 'Não sei nada, mas quero aprender!']
iniciante1_manha_ufu = manha_ufu[manha_ufu[nivel] == 'Não sei nada, mas quero aprender!']
iniciante1_manha_uni = manha_uni[manha_uni[nivel] == 'Não sei nada, mas quero aprender!']

iniciante2_manha_pit = manha_pit[manha_pit[nivel] == 'Sei o que é algoritmos, printf e scanf']
iniciante2_manha_ufu = manha_ufu[manha_ufu[nivel] == 'Sei o que é algoritmos, printf e scanf']
iniciante2_manha_uni = manha_uni[manha_uni[nivel] == 'Sei o que é algoritmos, printf e scanf']

intermediario_manha_pit = manha_pit[manha_pit[nivel] == 'Sei o que é string, vetor e matriz']
intermediario_manha_ufu = manha_ufu[manha_ufu[nivel] == 'Sei o que é string, vetor e matriz']
intermediario_manha_uni = manha_uni[manha_uni[nivel] == 'Sei o que é string, vetor e matriz']

avancado_manha_pit = manha_pit[manha_pit[nivel] == 'Sei o que é ordenação']
avancado_manha_ufu = manha_ufu[manha_ufu[nivel] == 'Sei o que é ordenação']
avancado_manha_uni = manha_uni[manha_uni[nivel] == 'Sei o que é ordenação']

#idade m1
m1_iniciante1_tarde_pit = iniciante1_tarde_pit[iniciante1_tarde_pit[age].isin(['Menos de 13', 14])]
m1_iniciante1_tarde_ufu = iniciante1_tarde_ufu[iniciante1_tarde_ufu[age].isin(['Menos de 13', 14])]
m1_iniciante1_tarde_uni = iniciante1_tarde_uni[iniciante1_tarde_uni[age].isin(['Menos de 13', 14])]
m1_iniciante2_tarde_pit = iniciante2_tarde_pit[iniciante2_tarde_pit[age].isin(['Menos de 13', 14])]
m1_iniciante2_tarde_ufu = iniciante2_tarde_ufu[iniciante2_tarde_ufu[age].isin(['Menos de 13', 14])]
m1_iniciante2_tarde_uni = iniciante2_tarde_uni[iniciante2_tarde_uni[age].isin(['Menos de 13', 14])]
m1_intermediario_tarde_pit = intermediario_tarde_pit[intermediario_tarde_pit[age].isin(['Menos de 13', 14])]
m1_intermediario_tarde_ufu = intermediario_tarde_ufu[intermediario_tarde_ufu[age].isin(['Menos de 13', 14])]
m1_intermediario_tarde_uni = intermediario_tarde_uni[intermediario_tarde_uni[age].isin(['Menos de 13', 14])]
m1_avancado_tarde_pit = avancado_tarde_pit[avancado_tarde_pit[age].isin(['Menos de 13', 14])]
m1_avancado_tarde_ufu = avancado_tarde_ufu[avancado_tarde_ufu[age].isin(['Menos de 13', 14])]
m1_avancado_tarde_uni = avancado_tarde_uni[avancado_tarde_uni[age].isin(['Menos de 13', 14])]

m1_iniciante1_manha_pit = iniciante1_manha_pit[iniciante1_manha_pit[age].isin(['Menos de 13', 14])]
m1_iniciante1_manha_ufu = iniciante1_manha_ufu[iniciante1_manha_ufu[age].isin(['Menos de 13', 14])]
m1_iniciante1_manha_uni = iniciante1_manha_uni[iniciante1_manha_uni[age].isin(['Menos de 13', 14])]
m1_iniciante2_manha_pit = iniciante2_manha_pit[iniciante2_manha_pit[age].isin(['Menos de 13', 14])]
m1_iniciante2_manha_ufu = iniciante2_manha_ufu[iniciante2_manha_ufu[age].isin(['Menos de 13', 14])]
m1_iniciante2_manha_uni = iniciante2_manha_uni[iniciante2_manha_uni[age].isin(['Menos de 13', 14])]
m1_intermediario_manha_pit = intermediario_manha_pit[intermediario_manha_pit[age].isin(['Menos de 13', 14])]
m1_intermediario_manha_ufu = intermediario_manha_ufu[intermediario_manha_ufu[age].isin(['Menos de 13', 14])]
m1_intermediario_manha_uni = intermediario_manha_uni[intermediario_manha_uni[age].isin(['Menos de 13', 14])]
m1_avancado_manha_pit = avancado_manha_pit[avancado_manha_pit[age].isin(['Menos de 13', 14])]
m1_avancado_manha_ufu = avancado_manha_ufu[avancado_manha_ufu[age].isin(['Mpitenos de 13', 14])]
m1_avancado_manha_uni = avancado_manha_uni[avancado_manha_uni[age].isin(['Menos de 13', 14])]

#idade m2
m2_iniciante1_tarde_pit = iniciante1_tarde_pit[iniciante1_tarde_pit[age].isin([15, 16])]
m2_iniciante1_tarde_ufu = iniciante1_tarde_ufu[iniciante1_tarde_ufu[age].isin([15, 16])]
m2_iniciante1_tarde_uni = iniciante1_tarde_uni[iniciante1_tarde_uni[age].isin([15, 16])]
m2_iniciante2_tarde_pit = iniciante2_tarde_pit[iniciante2_tarde_pit[age].isin([15, 16])]
m2_iniciante2_tarde_ufu = iniciante2_tarde_ufu[iniciante2_tarde_ufu[age].isin([15, 16])]
m2_iniciante2_tarde_uni = iniciante2_tarde_uni[iniciante2_tarde_uni[age].isin([15, 16])]
m2_intermediario_tarde_pit = intermediario_tarde_pit[intermediario_tarde_pit[age].isin([15, 16])]
m2_intermediario_tarde_ufu = intermediario_tarde_ufu[intermediario_tarde_ufu[age].isin([15, 16])]
m2_intermediario_tarde_uni = intermediario_tarde_uni[intermediario_tarde_uni[age].isin([15, 16])]
m2_avancado_tarde_pit = avancado_tarde_pit[avancado_tarde_pit[age].isin([15, 16])]
m2_avancado_tarde_ufu = avancado_tarde_ufu[avancado_tarde_ufu[age].isin([15, 16])]
m2_avancado_tarde_uni = avancado_tarde_uni[avancado_tarde_uni[age].isin([15, 16])]

m2_iniciante1_manha_pit = iniciante1_manha_pit[iniciante1_manha_pit[age].isin([15, 16])]
m2_iniciante1_manha_ufu = iniciante1_manha_ufu[iniciante1_manha_ufu[age].isin([15, 16])]
m2_iniciante1_manha_uni = iniciante1_manha_uni[iniciante1_manha_uni[age].isin([15, 16])]
m2_iniciante2_manha_pit = iniciante2_manha_pit[iniciante2_manha_pit[age].isin([15, 16])]
m2_iniciante2_manha_ufu = iniciante2_manha_ufu[iniciante2_manha_ufu[age].isin([15, 16])]
m2_iniciante2_manha_uni = iniciante2_manha_uni[iniciante2_manha_uni[age].isin([15, 16])]
m2_intermediario_manha_pit = intermediario_manha_pit[intermediario_manha_pit[age].isin([15, 16])]
m2_intermediario_manha_ufu = intermediario_manha_ufu[intermediario_manha_ufu[age].isin([15, 16])]
m2_intermediario_manha_uni = intermediario_manha_uni[intermediario_manha_uni[age].isin([15, 16])]
m2_avancado_manha_pit = avancado_manha_pit[avancado_manha_pit[age].isin([15, 16])]
m2_avancado_manha_ufu = avancado_manha_ufu[avancado_manha_ufu[age].isin([15, 16])]
m2_avancado_manha_uni = avancado_manha_uni[avancado_manha_uni[age].isin([15, 16])]

#idade m3
m3_iniciante1_tarde_pit = iniciante1_tarde_pit[iniciante1_tarde_pit[age].isin([17, 'Mais de 18'])]
m3_iniciante1_tarde_ufu = iniciante1_tarde_ufu[iniciante1_tarde_ufu[age].isin([17, 'Mais de 18'])]
m3_iniciante1_tarde_uni = iniciante1_tarde_uni[iniciante1_tarde_uni[age].isin([17, 'Mais de 18'])]
m3_iniciante2_tarde_pit = iniciante2_tarde_pit[iniciante2_tarde_pit[age].isin([17, 'Mais de 18'])]
m3_iniciante2_tarde_ufu = iniciante2_tarde_ufu[iniciante2_tarde_ufu[age].isin([17, 'Mais de 18'])]
m3_iniciante2_tarde_uni = iniciante2_tarde_uni[iniciante2_tarde_uni[age].isin([17, 'Mais de 18'])]
m3_intermediario_tarde_pit = intermediario_tarde_pit[intermediario_tarde_pit[age].isin([17, 'Mais de 18'])]
m3_intermediario_tarde_ufu = intermediario_tarde_ufu[intermediario_tarde_ufu[age].isin([17, 'Mais de 18'])]
m3_intermediario_tarde_uni = intermediario_tarde_uni[intermediario_tarde_uni[age].isin([17, 'Mais de 18'])]
m3_avancado_tarde_pit = avancado_tarde_pit[avancado_tarde_pit[age].isin([17, 'Mais de 18'])]
m3_avancado_tarde_ufu = avancado_tarde_ufu[avancado_tarde_ufu[age].isin([17, 'Mais de 18'])]
m3_avancado_tarde_uni = avancado_tarde_uni[avancado_tarde_uni[age].isin([17, 'Mais de 18'])]

m3_iniciante1_manha_pit = iniciante1_manha_pit[iniciante1_manha_pit[age].isin([17, 'Mais de 18'])]
m3_iniciante1_manha_ufu = iniciante1_manha_ufu[iniciante1_manha_ufu[age].isin([17, 'Mais de 18'])]
m3_iniciante1_manha_uni = iniciante1_manha_uni[iniciante1_manha_uni[age].isin([17, 'Mais de 18'])]
m3_iniciante2_manha_pit = iniciante2_manha_pit[iniciante2_manha_pit[age].isin([17, 'Mais de 18'])]
m3_iniciante2_manha_ufu = iniciante2_manha_ufu[iniciante2_manha_ufu[age].isin([17, 'Mais de 18'])]
m3_iniciante2_manha_uni = iniciante2_manha_uni[iniciante2_manha_uni[age].isin([17, 'Mais de 18'])]
m3_intermediario_manha_pit = intermediario_manha_pit[intermediario_manha_pit[age].isin([17, 'Mais de 18'])]
m3_intermediario_manha_ufu = intermediario_manha_ufu[intermediario_manha_ufu[age].isin([17, 'Mais de 18'])]
m3_intermediario_manha_uni = intermediario_manha_uni[intermediario_manha_uni[age].isin([17, 'Mais de 18'])]
m3_avancado_manha_pit = avancado_manha_pit[avancado_manha_pit[age].isin([17, 'Mais de 18'])]
m3_avancado_manha_ufu = avancado_manha_ufu[avancado_manha_ufu[age].isin([17, 'Mais de 18'])]
m3_avancado_manha_uni = avancado_manha_uni[avancado_manha_uni[age].isin([17, 'Mais de 18'])]


#sort_values
m1_iniciante1_tarde_pit.sort_values(name)
m1_iniciante1_tarde_ufu.sort_values(name)
m1_iniciante1_tarde_uni.sort_values(name)
m1_iniciante2_tarde_pit.sort_values(name)
m1_iniciante2_tarde_ufu.sort_values(name)
m1_iniciante2_tarde_uni.sort_values(name)
m1_intermediario_tarde_pit.sort_values(name)
m1_intermediario_tarde_ufu.sort_values(name)
m1_intermediario_tarde_uni.sort_values(name)
m1_avancado_tarde_pit.sort_values(name)
m1_avancado_tarde_ufu.sort_values(name)
m1_avancado_tarde_uni.sort_values(name)
m1_iniciante1_manha_pit.sort_values(name)
m1_iniciante1_manha_ufu.sort_values(name)
m1_iniciante1_manha_uni.sort_values(name)
m1_iniciante2_manha_pit.sort_values(name)
m1_iniciante2_manha_ufu.sort_values(name)
m1_iniciante2_manha_uni.sort_values(name)
m1_intermediario_manha_pit.sort_values(name)
m1_intermediario_manha_ufu.sort_values(name)
m1_intermediario_manha_uni.sort_values(name)
m1_avancado_manha_pit.sort_values(name)
m1_avancado_manha_ufu.sort_values(name)
m1_avancado_manha_uni.sort_values(name)
m2_iniciante1_tarde_pit.sort_values(name)
m2_iniciante1_tarde_ufu.sort_values(name)
m2_iniciante1_tarde_uni.sort_values(name)
m2_iniciante2_tarde_pit.sort_values(name)
m2_iniciante2_tarde_ufu.sort_values(name)
m2_iniciante2_tarde_uni.sort_values(name)
m2_intermediario_tarde_pit.sort_values(name)
m2_intermediario_tarde_ufu.sort_values(name)
m2_intermediario_tarde_uni.sort_values(name)
m2_avancado_tarde_pit.sort_values(name)
m2_avancado_tarde_ufu.sort_values(name)
m2_avancado_tarde_uni.sort_values(name)
m2_iniciante1_manha_pit.sort_values(name)
m2_iniciante1_manha_ufu.sort_values(name)
m2_iniciante1_manha_uni.sort_values(name)
m2_iniciante2_manha_pit.sort_values(name)
m2_iniciante2_manha_ufu.sort_values(name)
m2_iniciante2_manha_uni.sort_values(name)
m2_intermediario_manha_pit.sort_values(name)
m2_intermediario_manha_ufu.sort_values(name)
m2_intermediario_manha_uni.sort_values(name)
m2_avancado_manha_pit.sort_values(name)
m2_avancado_manha_ufu.sort_values(name)
m2_avancado_manha_uni.sort_values(name)
m3_iniciante1_tarde_pit.sort_values(name)
m3_iniciante1_tarde_ufu.sort_values(name)
m3_iniciante1_tarde_uni.sort_values(name)
m3_iniciante2_tarde_pit.sort_values(name)
m3_iniciante2_tarde_ufu.sort_values(name)
m3_iniciante2_tarde_uni.sort_values(name)
m3_intermediario_tarde_pit.sort_values(name)
m3_intermediario_tarde_ufu.sort_values(name)
m3_intermediario_tarde_uni.sort_values(name)
m3_avancado_tarde_pit.sort_values(name)
m3_avancado_tarde_ufu.sort_values(name)
m3_avancado_tarde_uni.sort_values(name)
m3_iniciante1_manha_pit.sort_values(name)
m3_iniciante1_manha_ufu.sort_values(name)
m3_iniciante1_manha_uni.sort_values(name)
m3_iniciante2_manha_pit.sort_values(name)
m3_iniciante2_manha_ufu.sort_values(name)
m3_iniciante2_manha_uni.sort_values(name)
m3_intermediario_manha_pit.sort_values(name)
m3_intermediario_manha_ufu.sort_values(name)
m3_intermediario_manha_uni.sort_values(name)
m3_avancado_manha_pit.sort_values(name)
m3_avancado_manha_ufu.sort_values(name)
m3_avancado_manha_uni.sort_values(name)


#Creates the new excel file
writer = pd.ExcelWriter("out/" + file_b + ".xls", engine='xlsxwriter')

m1_iniciante1_tarde_pit[name].to_excel(writer, sheet_name='Pitagoras tarde ini_1 <13')
m1_iniciante1_tarde_ufu[name].to_excel(writer, sheet_name='UFU tarde ini_1 <13')
m1_iniciante1_tarde_uni[name].to_excel(writer, sheet_name='Uniube tarde ini_1 <13')
m1_iniciante2_tarde_pit[name].to_excel(writer, sheet_name='Pitagoras tarde ini_2 <13')
m1_iniciante2_tarde_ufu[name].to_excel(writer, sheet_name='UFU tarde ini_2 <13')
m1_iniciante2_tarde_uni[name].to_excel(writer, sheet_name='Uniube tarde ini_2 <13')
m1_intermediario_tarde_pit[name].to_excel(writer, sheet_name='Pitagoras tarde inter <13')
m1_intermediario_tarde_ufu[name].to_excel(writer, sheet_name='UFU tarde inter <13')
m1_intermediario_tarde_uni[name].to_excel(writer, sheet_name='Uniube tarde inter <13')
m1_avancado_tarde_pit[name].to_excel(writer, sheet_name='Pitagoras tarde avan <13')
m1_avancado_tarde_ufu[name].to_excel(writer, sheet_name='UFU tarde avan <13')
m1_avancado_tarde_uni[name].to_excel(writer, sheet_name='Uniube tarde avan <13')
m1_iniciante1_manha_pit[name].to_excel(writer, sheet_name='Pitagoras manha ini_1 <13')
m1_iniciante1_manha_ufu[name].to_excel(writer, sheet_name='UFU manha ini_1 <13')
m1_iniciante1_manha_uni[name].to_excel(writer, sheet_name='Uniube manha ini_1 <13')
m1_iniciante2_manha_pit[name].to_excel(writer, sheet_name='Pitagoras manha ini_2 <13')
m1_iniciante2_manha_ufu[name].to_excel(writer, sheet_name='UFU manha ini_2 <13')
m1_iniciante2_manha_uni[name].to_excel(writer, sheet_name='Uniube manha ini_2 <13')
m1_intermediario_manha_pit[name].to_excel(writer, sheet_name='Pitagoras manha inter <13')
m1_intermediario_manha_ufu[name].to_excel(writer, sheet_name='UFU  manha inter <13')
m1_intermediario_manha_uni[name].to_excel(writer, sheet_name='Uniube manha inter <13')
m1_avancado_manha_pit[name].to_excel(writer, sheet_name='Pitagoras manha avan <13')
m1_avancado_manha_ufu[name].to_excel(writer, sheet_name='UFU manha avan <13')
m1_avancado_manha_uni[name].to_excel(writer, sheet_name='Uniube manha avan <13')
m2_iniciante1_tarde_pit[name].to_excel(writer, sheet_name='Pitagoras tarde ini_1')
m2_iniciante1_tarde_ufu[name].to_excel(writer, sheet_name='UFU tarde ini_1')
m2_iniciante1_tarde_uni[name].to_excel(writer, sheet_name='Uniube tarde ini_1')
m2_iniciante2_tarde_pit[name].to_excel(writer, sheet_name='Pitagoras tarde ini_2')
m2_iniciante2_tarde_ufu[name].to_excel(writer, sheet_name='UFU tarde ini_2')
m2_iniciante2_tarde_uni[name].to_excel(writer, sheet_name='Uniube tarde ini_2')
m2_intermediario_tarde_pit[name].to_excel(writer, sheet_name='Pitagoras tarde inter')
m2_intermediario_tarde_ufu[name].to_excel(writer, sheet_name='UFU tarde inter')
m2_intermediario_tarde_uni[name].to_excel(writer, sheet_name='Uniube tarde inter')
m2_avancado_tarde_pit[name].to_excel(writer, sheet_name='Pitagoras tarde avan')
m2_avancado_tarde_ufu[name].to_excel(writer, sheet_name='UFU tarde avan')
m2_avancado_tarde_uni[name].to_excel(writer, sheet_name='Uniube tarde avan')
m2_iniciante1_manha_pit[name].to_excel(writer, sheet_name='Pitagoras manha ini_1')
m2_iniciante1_manha_ufu[name].to_excel(writer, sheet_name='UFU manha ini_1')
m2_iniciante1_manha_uni[name].to_excel(writer, sheet_name='Uniube manha ini_1')
m2_iniciante2_manha_pit[name].to_excel(writer, sheet_name='Pitagoras manha ini_2')
m2_iniciante2_manha_ufu[name].to_excel(writer, sheet_name='UFU manha ini_2')
m2_iniciante2_manha_uni[name].to_excel(writer, sheet_name='Uniube manha ini_2')
m2_intermediario_manha_pit[name].to_excel(writer, sheet_name='Pitagoras manha inter')
m2_intermediario_manha_ufu[name].to_excel(writer, sheet_name='UFU manha inter')
m2_intermediario_manha_uni[name].to_excel(writer, sheet_name='Uniube manha inter')
m2_avancado_manha_pit[name].to_excel(writer, sheet_name='Pitagoras manha avan')
m2_avancado_manha_ufu[name].to_excel(writer, sheet_name='UFU manha avan')
m2_avancado_manha_uni[name].to_excel(writer, sheet_name='Uniube manha avan')
m3_iniciante1_tarde_pit[name].to_excel(writer, sheet_name='Pitagoras tarde ini_1 >18')
m3_iniciante1_tarde_ufu[name].to_excel(writer, sheet_name='UFU tarde ini_1 >18')
m3_iniciante1_tarde_uni[name].to_excel(writer, sheet_name='Uniube tarde ini_1 >18')
m3_iniciante2_tarde_pit[name].to_excel(writer, sheet_name='Pitagoras tarde ini_2 >18')
m3_iniciante2_tarde_ufu[name].to_excel(writer, sheet_name='UFU tarde ini_2 >18')
m3_iniciante2_tarde_uni[name].to_excel(writer, sheet_name='Uniube tarde ini_2 >18')
m3_intermediario_tarde_pit[name].to_excel(writer, sheet_name='Pitagoras tarde inter >18')
m3_intermediario_tarde_ufu[name].to_excel(writer, sheet_name='UFU tarde inter >18')
m3_intermediario_tarde_uni[name].to_excel(writer, sheet_name='Uniube tarde inter >18')
m3_avancado_tarde_pit[name].to_excel(writer, sheet_name='Pitagoras tarde avan >18')
m3_avancado_tarde_ufu[name].to_excel(writer, sheet_name='UFU tarde avan >18')
m3_avancado_tarde_uni[name].to_excel(writer, sheet_name='Uniube tarde avan >18')
m3_iniciante1_manha_pit[name].to_excel(writer, sheet_name='Pitagoras manha ini_1 >18')
m3_iniciante1_manha_ufu[name].to_excel(writer, sheet_name='UFU manha ini_1 >18')
m3_iniciante1_manha_uni[name].to_excel(writer, sheet_name='Uniube manha ini_1 >18')
m3_iniciante2_manha_pit[name].to_excel(writer, sheet_name='Pitagoras manha ini_2 >18')
m3_iniciante2_manha_ufu[name].to_excel(writer, sheet_name='UFU manha ini_2 >18')
m3_iniciante2_manha_uni[name].to_excel(writer, sheet_name='Uniube manha ini_2 >18')
m3_intermediario_manha_pit[name].to_excel(writer, sheet_name='Pitagoras  manha inter >18')
m3_intermediario_manha_ufu[name].to_excel(writer, sheet_name='UFU manha inter >18')
m3_intermediario_manha_uni[name].to_excel(writer, sheet_name='Uniube manha inter >18')
m3_avancado_manha_pit[name].to_excel(writer, sheet_name='Pitagoras manha avan >18')
m3_avancado_manha_ufu[name].to_excel(writer, sheet_name='UFU manha avan >18')
m3_avancado_manha_uni[name].to_excel(writer, sheet_name='Uniube manha avan >18')

writer.save()

print("Sucesso!! Arquivo direcionado para [out]")
