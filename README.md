# Course-project
import numpy as np
import openpyxl
from openpyxl import load_workbook
from gams import *
import os
import sys
import copy

DNA_SIZE = 10            
POP_SIZE = 50          
CROSS_RATE = 0.8         
MUTATION_RATE = 0.1    
N_GENERATIONS = 500
X_BOUND1 = [0, 16]         
X_BOUND2 = [0, 31]

# 最初始的bus schedule
file_origin = 'C:\\Users\\13404\\Desktop\\data.xlsx'

def translateDNA(pop):
    temp = copy.deepcopy(pop)
    for i in range(0,POP_SIZE):
        for i_ in range(1,6):
            temp[i][0] = int(pop[i][i_].dot(2 ** np.arange(DNA_SIZE)[::-1]) / float(2**DNA_SIZE-1) * X_BOUND1[1])
            temp[i][i_] = int(pop[i][i_].dot(2 ** np.arange(DNA_SIZE)[::-1]) / float(2**DNA_SIZE-1) * X_BOUND2[1])
    return (temp)

def get_fitness(pred):
    F_value = []
    for i in range(0,POP_SIZE):    
        wb = load_workbook(filename= file_origin)
        sheet_ranges = wb['Sheet1']
        ws = wb['Sheet1'] 
        for row in ws.iter_rows(min_col=2, max_col=56, min_row=2,max_row=5):
            for cell in  row:
                cell.value += int(pred[i][0])
                #print(cell.value)
        for row in ws.iter_rows(min_col=2, max_col=31, min_row=6,max_row=9):
            for cell in  row:
                cell.value += int(pred[i][1])
                #print(cell.value)
        for row in ws.iter_rows(min_col=2, max_col=19, min_row=10,max_row=13):
            for cell in  row:
                cell.value += int(pred[i][2])
                #print(cell.value)
        for row in ws.iter_rows(min_col=2, max_col=26, min_row=14,max_row=17):
            for cell in  row:
                cell.value += int(pred[i][3])
                #print(cell.value)
        for row in ws.iter_rows(min_col=2, max_col=28, min_row=18,max_row=21):
            for cell in  row:
                cell.value += int(pred[i][4])
                #print(cell.value)
        for row in ws.iter_rows(min_col=2, max_col=24, min_row=22,max_row=25):
            for cell in  row:
                cell.value += int(pred[i][5])
                #print(cell.value)
        wb.save('data.xlsx')    
        ws = GamsWorkspace(system_directory="E:/gams/34",working_directory="E:/gams/34")
        t1 = ws.add_job_from_file("example.gms")
        t1.run()
        for rec in t1.out_db["obj"]:
            obj_value = float(rec.level)
        F_value.append(obj_value)   
    return (F_value)
    
def select(pop, fitness):    
    fitness_arr = np.array(fitness)
    pop_new = []
    idx = np.random.choice(np.arange(POP_SIZE-1), size=POP_SIZE-1, replace=True,
                           p=(1 - fitness_arr/fitness_arr.sum())/(POP_SIZE-2))
    for i in idx:
        pop_new.append(pop[i])
    return (pop_new) 

def crossover(parent, pop):     
    parent_new = copy.deepcopy(parent)
    if np.random.rand() < CROSS_RATE:                                           
        i_ = np.random.randint(0, POP_SIZE-1, size=1)                             
        cross_points = np.random.randint(0, 2, size=DNA_SIZE).astype(np.bool)   
        a = np.array(pop[int(i_)])
        for i__ in range(0,6):
            parent_new[i__][cross_points] = a[i__, cross_points]
    return parent_new

def mutate(child):
    for point in range(DNA_SIZE):
        if np.random.rand() < MUTATION_RATE:
            for i__ in range(0,6):
                child[i__][point] = 1 if child[i__][point] == 0 else 0
    return child

pop1 = np.random.randint(2, size=(POP_SIZE, DNA_SIZE))
pop2 = np.random.randint(2, size=(POP_SIZE, DNA_SIZE))
pop3 = np.random.randint(2, size=(POP_SIZE, DNA_SIZE))
pop4 = np.random.randint(2, size=(POP_SIZE, DNA_SIZE))
pop5 = np.random.randint(2, size=(POP_SIZE, DNA_SIZE))
pop6 = np.random.randint(2, size=(POP_SIZE, DNA_SIZE))
pop_total = [pop1,pop2,pop3,pop4,pop5,pop6]

pop = []
w = []
i_ = 0
while i_ < POP_SIZE:
    for i in range(0,6):
        w.append(pop_total[i][i_])
    #print (w)
    pop.append(w)
    i_ = i_ + 1
    w = []
    
for _ in range(N_GENERATIONS):
    fitness = get_fitness(translateDNA(pop)) 
    min_index = np.argmin(fitness)
    print("Most fitted DNA: ", translateDNA(pop)[min_index])
    print("Obj_value: ", [float(fitness[min_index])])
    optimal = pop[min_index]
    print(fitness)
    del pop[min_index]
    del fitness[min_index]

    pop = select(pop, fitness)
    pop_copy = copy.deepcopy(pop)
    for i in range(len(pop)):
        child = crossover(pop[i], pop_copy)
        child = mutate(child)
        pop[i] = child       

    pop.append(optimal)
