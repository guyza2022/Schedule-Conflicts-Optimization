from constraints import constraints as cst
from gene import gene, geneChunk, geneBranch
from copy import deepcopy
import numpy as np
import random as r
import pandas as pd
import math as m

class chromosome:

    # Punishement constants
    conflict_punish_const: float = 0
    available_threshold: float = 0
    available_punish_const: float = 0
    const: cst = None # Constraints and pre-calculated data

    # Properties
    current_id: int = 0
    mutate_modes = ['Swap','Shift']

    # mutate possibility
    branch_swap_chance = 0.5 # 50 percent chance that each branch mutate with swap mode

    solution: pd.DataFrame
    solution_sheet = pd.read_excel('Solution2.xlsx').replace(np.nan,'_')
    solution = {}
    
    def __init__(self, mode: str):
        self.id = self.current_id
        self._increaseID()
        self.fitness_score = 0
        self.age = 0
        # Assign value to solution chromosome
        for room in self.const.room_names:
            self.solution[room] = list(self.solution_sheet.loc[:,room].values)
        if mode == 'Random':
            branches_template = deepcopy(self.const.chmsTemplate)
            self.branches = self._fillBranches(branches_template,self.const.request_dict)
            self.fitness_score = self.fitness(self.listEncode(self.branches))
    
    @classmethod
    def setupPunishmentConstants(cls, conflict_punish_const: float, available_threshold: float, available_punish_const: float, constraints: cst):
        # Assign values
        cls.conflict_punish_const = conflict_punish_const
        cls.available_threshold = available_threshold
        cls.available_punish_const = available_punish_const
        cls.const = constraints
        print('\nSetup Punishment Constants Successfully\n')

        return None
    
    def tupleEncode(self, branches: list):
        tuple_dict = {}
        # Assign list of tuple into dict with room name as a key
        for room, tpl in [branch.tupleEncode() for branch in branches]:
            tuple_dict[room] = tpl
        # Fill blank task with nan index
        for room in list(set([x[1] for x in self.const.nan_index])):
            if room in tuple_dict.keys():
                room_nan_index = [x[0] for x in self.const.nan_index if x[1] == room]
                current_index = 0
                passed_task = []
                to_add = []
                for task, duration in tuple_dict[room]:
                        current_index += duration
                        passed_task.append(task)
                        if current_index in room_nan_index:
                            count = 1
                            current_index += 1
                            while(current_index in room_nan_index):
                                count += 1
                                current_index += 1
                            to_add.append((len(passed_task),('_',count)))
                            count = 0
                count = 0
                for index, item in to_add:
                    tuple_dict[room].insert(index+count,item)
                    count += 1
        return tuple_dict
                    
    def listEncode(self, branches: list):
        list_dict = {}
        # Assign string into dict with room name as a key
        for room,lst in [branch.listEncode() for branch in branches]:
            list_dict[room] = lst
        # Fill blank task with nan index
        for room in list(set([x[1] for x in self.const.nan_index])):
            if room in list_dict.keys():
                room_nan_indexes = [x[0] for x in self.const.nan_index if x[1] == room]
                for index in room_nan_indexes:
                    list_dict[room].insert(index,'_')
            
        return list_dict # Dictionary format

    def dataFrameEncode(self):
        pass

    def fitness(self, list_encoded_dict: dict):
        previous_score = deepcopy(self.fitness_score)
        list_encoded = list(list_encoded_dict.values())
        zipped_list = zip(*list_encoded)
        score_list: list = []
        for index, slot in enumerate(zipped_list):
            score_list.append(self._calculateSlotScore(index, slot))
        total_score = np.average(score_list)
        # Age
        if previous_score >= total_score:
            self.age += 1
        else:
            self.age -= 1
        del previous_score
        return total_score

    def memorize(self): # Should be a class itself
        pass

    def mutate(self):
        mode = r.choice(self.mutate_modes)
        if mode == 'Swap':
            selected_branches = r.choices(self.branches, weights=np.full((len(self.branches),),self.branch_swap_chance))
            branch: geneBranch
            for branch in selected_branches:
                duration = r.randint(1,int(self.const.room_chunks_dict[branch.room_name]/2))
                gene_candidates = branch.getSwapCandidates(duration)
                if len(gene_candidates) >= 2:
                    scores = [self._calculateCandidateScore(candidate) for candidate in gene_candidates]
                    weights = self._softmax(scores, mode = 'Least')
                    from_data = r.choices(gene_candidates, weights = weights, k=1)[0]
                    gene_candidates = [gc for gc in gene_candidates if gc[0] != from_data[0]]
                    scores = [self._calculateCandidateScore(candidate) for candidate in gene_candidates]
                    weights = self._softmax(scores, mode = 'Max')
                    to_data = r.choices(gene_candidates, weights = weights, k=1)[0]
                    branch.swap(from_data, to_data)
                else:
                    print('Unswapable')
                
        elif mode == 'Shift':
            selected_branches = r.choices(self.branches, weights=np.full((len(self.branches),),self.branch_swap_chance))
            branch: geneBranch
            for branch in selected_branches:
                branch.shift()
        self.fitness_score = self.fitness(self.listEncode(self.branches))
        
    @classmethod
    def _increaseID(cls):
        cls.current_id += 1

    def _fillBranches(self, template: list, request_dict: dict):
        branch: geneBranch
        for branch in template:
            branch_name = branch.room_name
            branch_request = sorted(request_dict[branch_name].items(),key=  lambda x:x[1][1], reverse = True)
            branch.fillBranch(branch_request)

        return template
    
    # def __repr__(self) -> str:
        
    #     return ''

    def _calculateSlotScore(self, index: int, slot: tuple): # Horizontal-wise
        score: float = 0
        # Availability Score
        score_list = [self.const.score_dict[x][index] for x in slot if x != '_']
        gated_score_list = [self._gate(s, self.available_threshold) for s in score_list]
        available_score = np.average(gated_score_list)
        score = available_score
        score -= self._calculateConflictScore(slot)
        return score
    
    def _calculateConflictScore(self, slot: tuple):
        members_list: list = [self.const.task_member_dict[x] for x in slot if x!= '_'] 
        if len(members_list) > 1:
            members_count_list: list = [len(x) for x in members_list]
            average_members_count = np.average(members_count_list)
            mutual_members = list(set.intersection(*[set(list) for list in members_list]))
            return len(mutual_members)/average_members_count
        else:
            return 0
        
    def _calculateCandidateScore(self, candidate: list):
        sequence, genes = candidate # Unpack
        n_indexes = sum([g.duration for g in genes])
        scores = [self._calculateGeneScore(g) for g in genes]
        return sum(scores)/n_indexes # Average
        
    def _calculateGeneScore(self, gene: gene):
        if gene.duration > 1:
            return sum(self.const.score_dict[gene.task][i] for i in gene.global_index)
        else:
            return self.const.score_dict[gene.task][gene.global_index[0]]
    
    def _inverseScore(self, score_list: list):
        inversed_score_list = [self.const.slot_max_score-score for score in score_list]
        return inversed_score_list
    
    def _softmax(self, x: list, mode: str = 'Max'):
        if mode == 'Least':
            x = self._inverseScore(x)
        if not isinstance(x, np.ndarray):
            x = np.array(x)
        # Subtracting the maximum value for numerical stability
        x = x - np.max(x)
        # Compute the exponential of all elements
        exp_x = np.exp(x)
        # Compute the sum of exponential values
        sum_exp_x = np.sum(exp_x)
        # Calculate softmax values by dividing each element by the sum
        softmax_x = exp_x / sum_exp_x
        return softmax_x

    def _gate(self, value: float, threshold: float, mode: str = 'Circle', base: float = 0):
        if mode == 'Circle':
            if value <= threshold:
                return -m.sqrt(value*(-value+2))+1
            else:
                return value

    


            


        
            

            