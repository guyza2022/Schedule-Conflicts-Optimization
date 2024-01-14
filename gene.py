import random as r
import functools
import operator
from copy import deepcopy
import numpy as np

class gene:
    
    def __init__(self, task: str, duration: int, sequence: int, local_index: tuple, global_index: tuple) -> None:
        self.task = task
        self.duration = duration
        self.sequence = sequence
        self.local_index = local_index # Tuple of gene local indexes
        self.global_index = global_index # Tuple of gene global indexes
        
    def updateLocalIndex(self, new_local_index: tuple):
        self.local_index = new_local_index

    def updateGlobalIndex(self, new_global_index: tuple):
        self.global_index = new_global_index

    def __repr__(self) -> str:
        return f'gene#{self.sequence}({self.task},{self.duration})'

class geneChunk:

    def __init__(self, sequence, size: int, max_chunk_size: int, local_nan_index: list) -> None:
        self.genes = []
        self.sequence = sequence
        self.size = size
        self.current_size = 0
        self.available_size = size
        self.max_chunk_size = max_chunk_size
        self.local_nan_index = local_nan_index # Local index of nan(s)
        self.nan_size = len(self.local_nan_index) # Number of nan(s)
        self.current_gene_sequence = 0 # Keep tracking gene sequence

    def addGene(self, task, duration) -> None:
        if not self.isFull() and self.current_size + duration <= self.size:
            local_index = tuple(range(self.current_size, self.current_size+duration))
            global_index = tuple(range(self.sequence*self.max_chunk_size+self.current_size,self.sequence*self.max_chunk_size+self.current_size+duration))
            self.genes.append(gene(task, duration, self.current_gene_sequence, local_index, global_index))
            self._increaseGeneSequence() # Increase current gene sequence by 1
            self.current_size += duration
            self.available_size -= duration
        elif self.current_size + duration > self.size:
            print('Adding size is exceeded')
        else:
            print('Gene is already full')

    def findGeneCandidates(self, duration):
        durations_list = [g.duration for g in self.genes]
        candidates = self._findGeneSequencesOf(durations_list, duration)
        if len(candidates) > 0:
            return [(self.sequence,gene_sequence) for gene_sequence in candidates]
        else:
            return None
    
    def shift(self):
        if len(self.genes) >= 3:
            selected_gene_index = r.randint(0,len(self.genes)-1)
            if selected_gene_index == 0:
                target_gene_index = 1
            elif selected_gene_index == len(self.genes)-1:
                target_gene_index = len(self.genes)-2
            else:
                target_gene_index = selected_gene_index + r.choice([1,-1])
            selected_gene = deepcopy(self.genes[selected_gene_index])
            target_gene = deepcopy(self.genes[target_gene_index])
            self.genes[selected_gene_index] = deepcopy(target_gene)
            self.genes[target_gene_index] = deepcopy(selected_gene)
            del selected_gene
            del target_gene
        elif len(self.genes) == 2:
            selected_gene = deepcopy(self.genes[0])
            target_gene = deepcopy(self.genes[1])
            self.genes[0] = deepcopy(target_gene)
            self.genes[1] = deepcopy(selected_gene)
            del selected_gene
            del target_gene
        else:
            print('Unshiftable')

    def isFull(self):
        if self.current_size == self.size:
            return True
        else:
            return False
    
    def tupleEncode(self):
        return [(g.task,g.duration) for g in self.genes]
    
    def listEncode(self):
        list_genes = []
        g: gene
        for g in self.genes:
            for _ in range(g.duration):
                list_genes.append(g.task)
        return list_genes
    
    def _findGeneSequencesOf(self, durations_list: list, duration: int):
        start_index = 0
        stop_index = 0
        sequence_sum = 0
        all_sequences = []
        while(stop_index<=len(durations_list)-1):
            sequence_sum += durations_list[stop_index]
            if sequence_sum >= duration:
                if sequence_sum == duration:
                    all_sequences.append([self.genes[x] for x in range(start_index,stop_index+1)])
                start_index += 1
                stop_index = start_index
                sequence_sum = 0
            else:
                stop_index += 1
        return all_sequences
    
    def shuffleGenes(self):
        r.shuffle(self.genes)
        self.reCalculateGeneIndex()
        return None
    
    def reCalculateGeneIndex(self):
        local_index_tracker = 0
        global_index_tracker = self.sequence*self.max_chunk_size
        gene_sequence_tracker = 0
        for gene in self.genes:
            local_index = tuple(range(local_index_tracker,local_index_tracker+gene.duration))
            global_index = tuple(range(global_index_tracker, global_index_tracker+gene.duration))
            gene.updateLocalIndex(local_index)
            gene.updateGlobalIndex(global_index)
            gene.sequence = gene_sequence_tracker
            local_index_tracker += gene.duration
            global_index_tracker += gene.duration
            gene_sequence_tracker += 1

    def deleteGenes(self, start_index: int, stop_index: int = None):
        # Offset is used when there are more than 1 deletion on the same chunk (Sequences index are inaccurate)
        if stop_index is not None:
            del self.genes[start_index:stop_index+1]
        else:
            del self.genes[start_index]
        
    
    def _increaseGeneSequence(self):
        self.current_gene_sequence += 1
        return None
    
    def __repr__(self) -> str:
        return f'chunk#{self.sequence}{[x.__repr__() for x in self.genes]}'
        
class geneBranch:

    nan_index: list

    def __init__(self, room_name: str, n_chunks: int, chunk_size: int, max_chunk_size: int, branch_nan_index: list):
        self.chunks = []
        self.room_name = room_name
        self.n_chunks = n_chunks
        self.chunk_size = chunk_size # Branch chunk size
        self.max_chunk_size = max_chunk_size # Max chunk size amoung all branch
        self.branch_nan_index = branch_nan_index
        # Auto generate chunks when created
        for i in range(n_chunks):
            local_nan_index = [x for x in branch_nan_index if x >= i*self.max_chunk_size and x <= (i+1)*self.max_chunk_size]
            self.chunks.append(geneChunk(i,chunk_size, max_chunk_size, local_nan_index))
        return None

    def fillBranch(self, branch_request)->None:
        chunk: geneChunk
        for task, request in branch_request:
            times, duration = request
            for _ in range(int(times)):
                chunk = self.findChunkAvailable(duration)
                if chunk is not None:
                    chunk.addGene(task,duration)
                else:
                    print(f'Can\'t find chunk for {task} duration {duration} ')
        # Shuffle elements in chunks
        chunk: geneChunk
        for chunk in self.chunks:
            chunk.shuffleGenes()
    
    def getSwapCandidates(self, duration: int):
        candidates = []
        chunk: geneChunk
        for chunk in self.chunks:
            gene_candidates = chunk.findGeneCandidates(duration)
            if gene_candidates is not None:
                candidates.append(gene_candidates)
        total_gene_candidates = functools.reduce(operator.iconcat, candidates, []) # List flattening
        return total_gene_candidates
    
    def swap(self, from_data: tuple, to_data: tuple):
        from_chunk_sequence, from_genes_list = deepcopy(from_data)
        from_genes_sequence_list = [g.sequence for g in from_genes_list]
        to_chunk_sequence, to_genes_list = deepcopy(to_data)
        to_genes_sequence_list = [g.sequence for g in to_genes_list]
        if len(from_genes_list) > 1:
            self.chunks[from_chunk_sequence].deleteGenes(from_genes_sequence_list[0], from_genes_sequence_list[-1])
        else:
            self.chunks[from_chunk_sequence].deleteGenes(from_genes_sequence_list[0])
        if len(to_genes_list) > 1:
            self.chunks[to_chunk_sequence].deleteGenes(to_genes_sequence_list[0],to_genes_sequence_list[-1])
        else:
            self.chunks[to_chunk_sequence].deleteGenes(to_genes_sequence_list[0])
        self.chunks[from_chunk_sequence].genes[from_genes_sequence_list[0]:from_genes_sequence_list[0]] =  to_genes_list
        self.chunks[to_chunk_sequence].genes[to_genes_sequence_list[0]:to_genes_sequence_list[0]] = from_genes_list
        self.chunks[from_chunk_sequence].reCalculateGeneIndex()
        self.chunks[to_chunk_sequence].reCalculateGeneIndex()
        del from_chunk_sequence
        del to_chunk_sequence


    def shift(self):
        chunk: geneChunk
        for chunk in self.chunks:
            chunk.shift()
            chunk.reCalculateGeneIndex()
        return None
        

    def findChunkAvailable(self, duration):
        availables = []
        chunk: geneChunk
        for chunk in self.chunks:
            if chunk.available_size >= duration:
                availables.append(chunk.sequence)
        if len(availables) > 0:
            return self.chunks[r.choice(availables)]
        else:
            return None
    
    def tupleEncode(self):
        tuple_chunk = [chunk.tupleEncode() for chunk in self.chunks]
        tuple_chunk = sum(tuple_chunk,[])
        return (self.room_name, tuple_chunk)
    
    def listEncode(self):
        list_chunks = [chunk.listEncode() for chunk in self.chunks]
        list_chunks = sum(list_chunks,[])
        return (self.room_name, list_chunks)

            

        

