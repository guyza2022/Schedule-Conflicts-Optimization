from chromosomes import chromosome as chms
from constraints import constraints as cst
import pandas as pd
import numpy as np
import datetime as dt
import random as r
from copy import deepcopy
import math as m
import regex as re
import matplotlib.pyplot as plt
import matplotlib
import statistics
from sklearn import preprocessing
from collections import OrderedDict
import sys
import tkinter as tk
import threading
import time as t
from collections import Counter

class genetic:

    # Population
    population: list = []

    def __init__(self, n_chromosomes: int, n_generations: int, constraints_sheet_name: str,\
                 conflict_punish_const: float, available_threshold: float, available_punish_const: float):
        # Check validity of parameters
        assert n_chromosomes > 0
        assert n_generations > 0
        assert constraints_sheet_name.endswith('.xlsx'), 'Wrong File Type'
        assert conflict_punish_const >= 0 and conflict_punish_const <= 1
        assert available_threshold >= 0 and available_threshold <= 1
        assert available_punish_const >= 0 and available_punish_const <= 1
        print('All Parameters Passed the Requirements')
        # Die age
        self.die_age = 10#m.ceil(0.0001*n_generations)
        # Number of population and generations
        self.n_chromosomes = n_chromosomes
        self.n_generations = n_generations
        # Constraint sheet name
        self.constraints_sheet_name = constraints_sheet_name
        # Fitness constraints
        self.conflict_punish_score_const = conflict_punish_const
        self.available_threshold = available_threshold
        self.available_punish_const = available_punish_const
        # Generate constraints
        self.constraints = self.generateConstraints(self.constraints_sheet_name)
        # Setup punishment constants
        chms.setupPunishmentConstants(conflict_punish_const, available_threshold, available_punish_const, self.constraints)

        return None
    
    def generateConstraints(self, constraint_sheet_name: str):
        constraints = cst(constraint_sheet_name)
        return constraints
        
    def initializePopulation(self, n_chromosomes: int):
        # Randomly generate chomosome
        self.population = []
        for i in range(n_chromosomes):
            self.population.append(chms(mode = 'Random'))
            #print(f'Initialized Chromosome Number: {i+1}/{n_chromosomes}')
        print('\nInitialization Completed\n')

    def generateNewChromosomes(self, n_chromosome):
        for i in range(n_chromosome):
            self.population.append(chms(mode = 'Random'))
            #print(f'Generating Chromosome Number: {i+1}/{n_chromosome}')
        print('\nGeneration Completed\n')
        
    def sortPopulation(self):
        print(f'Max age: {max([p.age for p in self.population])}')
        self.population = [p for p in self.population if p.age <= self.die_age]
        if len(self.population) < self.n_chromosomes:
            print(f'{self.n_chromosomes-len(self.population)} chromosomes died this round')
        population_with_score_list = [(p,p.fitness_score) for p in self.population]
        self.population = [p[0] for p in sorted(population_with_score_list, key=lambda x:x[1], reverse=True)]
        return None

    def selectBestPopulation(self, ratio: int):
        start_index = m.floor(ratio*len(self.population))
        del self.population[start_index:]
        return None
    
    def crossingOver(self, ratio):
        for _ in range(m.floor(ratio*self.n_chromosomes)):
            first_parent: chms
            second_parent: chms
            first_parent = deepcopy(self.population[r.randint(0,len(self.population)-1)])
            first_parent_id = first_parent.id
            second_parent_candidates = [x for x in self.population if x.id != first_parent_id]
            #print(second_parent_candidates)
            second_parent = r.choice(second_parent_candidates)
            selected_branch_index = r.randint(0,len(self.constraints.room_names)-1)
            second_parent_branch = deepcopy(second_parent.branches[selected_branch_index])
            first_parent.branches[selected_branch_index] = second_parent_branch
            chms._increaseID()
            first_parent.id = chms.current_id
            #first_parent.age = 0
            self.population.append(deepcopy(first_parent))
            del second_parent_branch
            del first_parent
        return None

    def generateNewPopulation(self):
        missing_amount = self.n_chromosomes - len(self.population)
        self.generateNewChromosomes(missing_amount)
        return None
    
    def showMaxFitness(self):
        print(f'Current max fitness is: {self.population[0].fitness_score}')
    
    def showAverageFitness(self):
        print(f'Current average fitness is: {np.average([p.fitness_score for p in self.population])}')
        
