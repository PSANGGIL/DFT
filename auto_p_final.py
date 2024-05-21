# -*- coding: utf-8 -*-

import re
import os, string, sys, shutil
import fileinput
from contextlib import closing
import pandas as pd

file_list = os.listdir(os.getcwd())
file_list_out = [file for file in file_list if file.endswith(".out")]
error_list = []
Imaginary_list = []


for lo in file_list_out:
    with open("{}".format(lo), "rt") as read_file:
               lines = read_file.readlines()
               a = ''.join(lines)
               b = a.count("Thank you very much for using Q-Chem.")
               c = a.count("   This Molecule has  0 Imaginary Frequencies")
               d = a.count("$rem")
               read_file.close()

               if a.count('Imaginary Frequencies') == 1:
                   if int(d) == b and c == 1:
                       print('{}'.format(lo), 'pass')


                   else:
                       print('{}'.format(lo), 'error')
                       error_list.append('{}'.format(lo))

               else:
                    print('{}'.format(lo), 'Imaginary Frequencies')
                    Imaginary_list.append('{}'.format(lo))

# directory
pwd = os.getcwd() + "/"
error_f = os.getcwd()+"/error/"
Imaginary_f = os.getcwd()+"/Imaginary/"
os.mkdir(error_f)
os.mkdir(Imaginary_f)


for ef in error_list:
    efn = "{}".format(ef)
    shutil.move(pwd+efn, error_f+efn)

for imf in Imaginary_list:
    imfn = "{}".format(imf)
    shutil.move(pwd+imfn, Imaginary_f+imfn)

# excel

file_list = os.listdir(os.getcwd())
file_list_out = [file for file in file_list if file.endswith(".out")]

df = pd.DataFrame({'file_name' : [], 'total energy(hartree)' :[], 'total enthalpy(kcal/mol)' : [], 'total entropy(cal/mol.K)' : [], 'HOMO(eV)' : [], 'LUMO(eV)' : [])

a = pd.DataFrame(data=[[0,0,0,0,0,0]], columns = ['file_name', 'total energy(hartree)', 'total enthalpy(kcal/mol)', 'total entropy(cal/mol.K)', 'HOMO(eV)', 'LUMO(eV)'])

for lo in file_list_out:
    with open("{}".format(lo), "rt") as read_file:
               lines = read_file.readlines()
               a = ''.join(lines)
               p = re.compile('^ Total energy in the final basis set =.*\n',re.MULTILINE)
               b = p.findall(a)
               c = '\n'.join(b)
               d = list(c.replace(' Total energy in the final basis set =',''.format(lo)).replace('.out','').split('\n'))

               a1 = ''.join(lines)
               p1 = re.compile('^   Total Enthalpy:.*\n',re.MULTILINE)
               b1 = p1.findall(a1)
               c1 = '\n'.join(b1)
               d1 = list(c1.replace('   Total Enthalpy:',''.format(lo)).replace('.out','').replace(' kcal/mol','').split('\n'))

               a2 = ''.join(lines)
               p2 = re.compile('^   Total Entropy:.*\n',re.MULTILINE)
               b2 = p2.findall(a2)
               c2 = '\n'.join(b2)
               d2 = list(c2.replace('   Total Entropy:',''.format(lo)).replace('.out','').replace('  cal/mol.K','').split('\n'))


               a_h = ''.join(lines)
               p_h = re.compile('^ -- Occupied --.*? -- Virtual --',re.MULTILINE | re.DOTALL)
               b_h = p_h.findall(a_h)
               c_h = '\n'.join(b_h)
               d_h = list(c_h.split())
               print("{}".format(lo),d_h[-4])


               a_l = ''.join(lines)
               p_l = re.compile('^ -- Virtual --.{9}',re.MULTILINE | re.DOTALL)
               b_l = p_l.findall(a_l)
               c_l = '\n'.join(b_l)
               d_l = list(c_l.split())
               print("{}".format(lo),d_l[-1])

               ad0 = pd.DataFrame(data=[['{}'.format(lo).replace('.out',''), d[-2], d1[0], d2[0], d_h[-4], d_l[-1]]], columns = ['file_name', 'total energy(hartree)', 'total enthalpy(kcal/mol)', 'total entropy(cal/mol.K)', 'HOMO(eV)', 'LUMO(eV)'])
               df = df.append(ad0)

df = df.reset_index(drop = True)

df.to_excel('scratch.xlsx', sheet_name = 'Sheet1')

df1 = pd.read_excel('scratch.xlsx', sheet_name = 'Sheet1')
df1['Gibbs energy(kcal/mol)'] = (627.5095*df1['total energy(hartree)'] +df1['total enthalpy(kcal/mol)'] - (298.15*df1['total entropy(cal/mol.K)']/1000))
df1['HOMO(eV)'] = (27.2114*df1['HOMO(eV)'])
df1['LUMO(eV)'] = (27.2114*df1['LUMO(eV)'])

df2 = df1.set_index('file_name')

df2.to_excel('result.xlsx', sheet_name = 'Sheet1')
os.remove('scratch.xlsx')
