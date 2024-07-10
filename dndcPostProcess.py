# %%

# import libraries
import json
import numpy as np
import statistics
import psycopg2

# Read and process DNDC data
class Process:
    
    def __init__(self, projectName, dndcData = None):

        self.filename = None
        self.projectName = projectName
        self.fieldNames = []

        self.dsocBaselineDistributions = {}
        self.directn2oBaselineDistributions = {}
        self.indirectn2oBaseline = {}
        self.ch4Baseline = {}

        self.dsocPracticeDistributions = {}
        self.directn2oPracticeDistributions = {}
        self.indirectn2oPractice = {}
        self.ch4Practice = {}

        self.baseline = {}
        self.practice = {}
        self.project = {}
        self.outcomes = {}
        
        self.base = {}
        self.prac = {}
        
        self.dsocBaselineP50 = []
        self.dsocPracticeP50 = []
        self.directn2oBaselineP50 = []
        self.directn2oPracticeP50 = []
        
        self.dsocAggregatedBaselineDistribution = None
        self.directn2oAggregatedBaselineDistribution = None
        self.dsocAggregatedPracticeDistribution = None
        self.directn2oAggregatedPracticeDistribution = None
        
        if dndcData:
            self.processData(dndcData)
        

#%%
    def processData(self, dndcJson):
        
        for field in dndcJson['data']['session_uncertainties']:
            self.fieldNames.append(field)

            # Baseline
            # Retrieve dsoc and directn2o distributions
            # Retrieve indirectn2o and ch4 values
            # Generate dsco and directn2o aggregated distributions
            self.dsocBaselineDistributions[field] = dndcJson['data']['session_uncertainties'][field]['scenarios_uncertainties']['baseline']['dsoc']['distribution']
            if self.dsocAggregatedBaselineDistribution:
                self.dsocAggregatedBaselineDistribution = [x + y for x, y in zip(self.dsocAggregatedBaselineDistribution, self.dsocBaselineDistributions[field])]
            else:
                self.dsocAggregatedBaselineDistribution = self.dsocBaselineDistributions[field]
            self.directn2oBaselineDistributions[field] = dndcJson['data']['session_uncertainties'][field]['scenarios_uncertainties']['baseline']['direct_n2o']['distribution']
            if self.directn2oAggregatedBaselineDistribution:
                self.directn2oAggregatedBaselineDistribution = [x + y for x, y in zip(self.directn2oAggregatedBaselineDistribution, self.directn2oBaselineDistributions[field])]
            else:
                self.directn2oAggregatedBaselineDistribution = self.directn2oBaselineDistributions[field]

            # Practice change
            # Retrieve dsoc and directn2o distributions
            # Retrieve indirectn2o and ch4 values
            # Generate dsco and directn2o aggregated distributions
            self.dsocPracticeDistributions[field] = dndcJson['data']['session_uncertainties'][field]['scenarios_uncertainties']['practice_change']['dsoc']['distribution']
            if self.dsocAggregatedPracticeDistribution:
                self.dsocAggregatedPracticeDistribution = [x + y for x, y in zip(self.dsocAggregatedPracticeDistribution, self.dsocPracticeDistributions[field])]
            else:
                self.dsocAggregatedPracticeDistribution = self.dsocPracticeDistributions[field]
            self.directn2oPracticeDistributions[field] = dndcJson['data']['session_uncertainties'][field]['scenarios_uncertainties']['practice_change']['direct_n2o']['distribution']
            if self.directn2oAggregatedPracticeDistribution:
                self.directn2oAggregatedPracticeDistribution = [x + y for x, y in zip(self.directn2oAggregatedPracticeDistribution, self.directn2oPracticeDistributions[field])]
            else:
                self.directn2oAggregatedPracticeDistribution = self.directn2oPracticeDistributions[field]
            
            # Calculate the P50 dsoc and directn2o values for Baseline and Practice Change
            self.base[field] = {'dsocP50':np.percentile(self.dsocBaselineDistributions[field],50),
                                'directn2oP50':np.percentile(self.directn2oBaselineDistributions[field],50),
                                'indirectn2o':dndcJson['data']['session_uncertainties'][field]['scenarios_uncertainties']['baseline']['indirect_n2o'],
                                'ch4':dndcJson['data']['session_uncertainties'][field]['scenarios_uncertainties']['baseline']['ch4']}
            self.prac[field] = {'dsocP50':np.percentile(self.dsocPracticeDistributions[field],50),
                                'directn2oP50':np.percentile(self.directn2oPracticeDistributions[field],50),
                                'indirectn2o':dndcJson['data']['session_uncertainties'][field]['scenarios_uncertainties']['practice_change']['indirect_n2o'],
                                'ch4':dndcJson['data']['session_uncertainties'][field]['scenarios_uncertainties']['practice_change']['ch4']}
            
            self.dsocBaselineP50.append(self.base[field]['dsocP50'])
            self.directn2oBaselineP50.append(self.base[field]['directn2oP50'])
            self.dsocPracticeP50.append(self.prac[field]['dsocP50'])
            self.directn2oPracticeP50.append(self.prac[field]['directn2oP50'])
        
        self.baseline['fields'] = self.base
        self.practice['fields'] = self.prac
            
        # Calculate aggregate dsoc and directn2o based on aggregated distributions
        self.baseline["dsocAggregation"] = np.percentile(self.dsocAggregatedBaselineDistribution,52.5)
        self.baseline["dsocAggregationStandardDeviation"] = statistics.stdev(self.dsocAggregatedBaselineDistribution)
        self.baseline["directn2oAggregation"] = np.percentile(self.directn2oAggregatedBaselineDistribution,47.5)
        self.baseline["directn2oAggregationStandardDeviation"] = statistics.stdev(self.directn2oAggregatedBaselineDistribution)
        
        self.practice["dsocAggregation"] = np.percentile(self.dsocAggregatedPracticeDistribution,47.5)
        self.practice["dsocAggregationStandardDeviation"] = statistics.stdev(self.dsocAggregatedPracticeDistribution)
        self.practice["directn2oAggregation"] = np.percentile(self.directn2oAggregatedPracticeDistribution,52.5)
        self.practice["directn2oAggregationStandardDeviation"] = statistics.stdev(self.directn2oAggregatedPracticeDistribution)
        
        # Calculate dsoc and directn2o adjusted values
        self.dsocBaselineAdjustedSum = 0
        self.directn2oBaselineAdjustedSum = 0
        self.dsocPracticeAdjustedSum = 0
        self.directn2oPracticeAdjustedSum = 0
        for field in self.baseline['fields']:
            self.baseline['fields'][field]['dsocAdjusted'] = self.baseline['fields'][field]['dsocP50']+abs(min(min(self.dsocBaselineP50),0))
            self.baseline['fields'][field]['directn2oAdjusted'] = self.baseline['fields'][field]['directn2oP50']+abs(min(min(self.directn2oBaselineP50),0))
            self.dsocBaselineAdjustedSum += self.baseline['fields'][field]['dsocAdjusted']
            self.directn2oBaselineAdjustedSum += self.baseline['fields'][field]['directn2oAdjusted']
                        
            self.practice['fields'][field]['dsocAdjusted'] = self.practice['fields'][field]['dsocP50']+abs(min(min(self.dsocPracticeP50),0))
            self.practice['fields'][field]['directn2oAdjusted'] = self.practice['fields'][field]['directn2oP50']+abs(min(min(self.directn2oPracticeP50),0))
            self.dsocPracticeAdjustedSum += self.practice['fields'][field]['dsocAdjusted']
            self.directn2oPracticeAdjustedSum += self.practice['fields'][field]['directn2oAdjusted']
        
        # Caclculate dsoc and directn2o final values
        for field in self.baseline['fields']:
            self.baseline['fields'][field]['dsocFinal'] = self.baseline["dsocAggregation"] * self.baseline['fields'][field]['dsocAdjusted'] / self.dsocBaselineAdjustedSum
            self.baseline['fields'][field]['directn2oFinal'] = self.baseline["directn2oAggregation"] * self.baseline['fields'][field]['directn2oAdjusted'] / self.directn2oBaselineAdjustedSum
            self.practice['fields'][field]['dsocFinal'] = self.practice["dsocAggregation"] * self.practice['fields'][field]['dsocAdjusted'] / self.dsocPracticeAdjustedSum
            self.practice['fields'][field]['directn2oFinal'] = self.practice["directn2oAggregation"] * self.practice['fields'][field]['directn2oAdjusted'] / self.directn2oPracticeAdjustedSum
            
            self.outcomes[field] = {'dsocOutcome':self.practice['fields'][field]['dsocFinal'] - self.baseline['fields'][field]['dsocFinal'],
                                    'directn2oOutcome':self.baseline['fields'][field]['directn2oFinal'] - self.practice['fields'][field]['directn2oFinal']}
        
# %%
    def processDataFromFile(self, dndcFilename):
        
        try:
            with open(dndcFilename) as file:
                dndcJson = json.load(file)
                self.filename = dndcFilename
                
                self.processData(dndcJson)
                
        except:
            raise Exception("Invalid DNDC JSON format")


# %%
    def printOutcomes(self):
        
        print("Baseline results")
        print(json.dumps(self.baseline,indent=4,sort_keys=True))
        
        print("\nPractice change results")
        print(json.dumps(self.practice,indent=4,sort_keys=True))
        
        print("\nOutcomes")
        print(json.dumps(self.outcomes,indent=4,sort_keys=True))

#%%

    def saveResults(self, curr):
        print('Writing outcomes')
        self.saveResultsToDb(self.baseline['fields'], True, curr)
        self.saveResultsToDb(self.practice['fields'], False, curr)
            

#%%
    def saveResultsToDb(self, data, isBaseline, curr):
        
        assert(curr)
        
        for field in data:
            curr.execute(f"INSERT INTO esmc.model_dndc_result (model_dndc__id, is_baseline, n2o_direct, n2o_indirect, methane, dsoc) SELECT md.id as model_dndc__id, {isBaseline} as is_baseline, {data[field]['directn2oFinal']} as n2o_direct, {data[field]['indirectn2o']} as n2o_indirect, {data[field]['ch4']} as methane, {data[field]['dsocFinal']} as dsoc FROM esmc.model_dndc md WHERE \'{field}\' = md.session_name and \'{self.projectName}\' = md.project_name;")
