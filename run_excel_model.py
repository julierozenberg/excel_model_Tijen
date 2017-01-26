import sys
sys.path.insert(0, 'C:\\Users\\julierozenberg\\Documents\\GitHub\\EMAworkbench\\src')

from expWorkbench import ModelEnsemble, ParameterUncertainty,\
                         Outcome, ema_logging
from connectors.excel import ExcelModelStructureInterface
from analysis.plotting import lines
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np

class ExcelModel(ExcelModelStructureInterface):
    
    uncertainties = [ParameterUncertainty((0.3, 2),"COSTAG"), 
                     ParameterUncertainty((-0.01,0.02),"Adgr"), 
                     ParameterUncertainty((0,0.02),"AVAWO"),
                     ParameterUncertainty((0.1,1),"ITD"),
                     ParameterUncertainty((0.1,1),"INTANGD"),
                     ParameterUncertainty((0,3),"SLRindex"),
                     ParameterUncertainty((0,0.12),"D"),
                     ParameterUncertainty((-0.005,0.005),"POPwo"),
                     ParameterUncertainty((0,0.01),"Adpopgr"),
                     ]
    
    #specification of the outcomes
    outcomes = [Outcome("IRR", time=False),  
                Outcome("NPV", time=False),
                Outcome("BC", time=False),
                Outcome("PV_EAAP", time=False),
                Outcome("PV_EAAP_C", time=False),
                ] 
    
    #name of the sheet
    sheet = "Econ_analysis"
    
    #relative path to the Excel file
    workbook = r'/CB Analysis1 CFNAMES_012317_OPTION 8.xlsx'
    
    def model_init(self, policy, kwargs):
        '''initializes the model'''
        
        try:
            self.workbook = policy['file']
        except KeyError:
            ema_logging.warning("key 'file' not found in policy")
        super(ExcelModel, self).model_init(policy, kwargs)

    

 
ema_logging.log_to_stderr(level=ema_logging.INFO)

model = ExcelModel(r"excel", "marshalls")

ensemble = ModelEnsemble()
ensemble.set_model_structure(model)

#add policies
policies = [{'name': 'option 7',
             'file': r'\CB Analysis1 CFNAMES_012217_OPTION 7.xlsx'},
            {'name': 'option 8',
             'file': r'\CB Analysis1 CFNAMES_012317_OPTION 8.xlsx'},
            ]
ensemble.add_policies(policies)


ensemble.parallel = False #turn on parallel computing
ensemble.processes = 2 #using only 2 cores 

#run 100 experiments
nr_experiments = 2000
(inputs,outputs) = ensemble.perform_experiments(nr_experiments)

results = pd.DataFrame(outputs,columns=outputs.keys())
uncertainties = pd.DataFrame(inputs,columns=inputs.dtype.names)
uncertainties.SLRindex = np.ceil(uncertainties.SLRindex)

all_results = pd.concat([uncertainties,results],axis=1)

uncertainties.to_csv("uncertainties_2options_round2.csv",index=False)
results.to_csv("results_2options_round2.csv",index=False)


