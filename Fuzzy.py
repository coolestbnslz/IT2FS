import numpy as np
import skfuzzy as fuzz
from skfuzzy import control as ctrl

quality = ctrl.Antecedent(np.arange(0, 11, 1), 'quality')
service = ctrl.Antecedent(np.arange(0, 11, 1), 'service')
rating = ctrl.Consequent(np.arange(0, 11, 1), 'rating')


quality['poor'] = fuzz.trimf(quality.universe, [0, 0, 5])
quality['medium']=fuzz.trimf(quality.universe,[2, 5, 8])
quality['medium1'] = fuzz.trimf(quality.universe, [0, 5, 10])
quality['excellent'] = fuzz.trimf(quality.universe, [5, 10,10 ])

service['poor'] = fuzz.trimf(service.universe, [0, 0, 5])
service['medium']=fuzz.trimf(quality.universe,[2, 5, 8])
service['medium1'] = fuzz.trimf(service.universe, [0, 5, 10])
service['excellent'] = fuzz.trimf(service.universe, [5, 10,10 ])


rating['poor'] = fuzz.trimf(rating.universe, [0, 0, 5])
rating['average']=fuzz.trimf(quality.universe,[2, 5, 8])
rating['average1'] = fuzz.trimf(rating.universe, [0, 5, 10])
rating['great'] = fuzz.trimf(rating.universe, [5, 10,10 ])
rating.view()
quality.view()
service.view()
rule1 = ctrl.Rule(quality['poor'] | service['poor'], rating['poor'])
rule2 = ctrl.Rule(service['medium'],rating['average'])
rule4=ctrl.Rule(service['medium1'],rating['average1'])
rule5=ctrl.Rule(quality['medium1'],rating['average1'])
rule3 = ctrl.Rule(service['excellent'] | quality['excellent'], rating['great'])

rating_ctrl = ctrl.ControlSystem([rule4, rule2,rule1,rule3,rule5])
tipping = ctrl.ControlSystemSimulation(rating_ctrl)
tipping.input['quality'] = 3
tipping.input['service'] = 7
tipping.compute()
print (tipping.output['rating'])
rating.view(sim=tipping)