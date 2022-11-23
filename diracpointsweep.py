import coppermountainvna as vna
from keithley2600 import Keithley2600
import time
import numpy as np
import matplotlib.pyplot as plt

"""
Diver for GPIB-USB-HS form National Instuments that connects the device to a
USB port must be installed first as has the Python moduel keithley 26000 which 
can be found on GitHub. Other softwhere may be requied to be installed.
If drivers for the instument itself are needed they can be found on the
Tektronix website.
"""

class GateVoltageMeasurment:
    
    def __init__(self, address, parameters, peak_width = 40000000, compute_Q_db = False, compute_Q_fit = False, plot_Q = False):
        
        # initialise keithley 2636B souce meter
        self.adress = address
        self.keith = Keithley2600(address, visa_library='')
        self.vna_measurment = vna.MultiMeasurment(parameters, peak_width, compute_Q_db, compute_Q_fit, plot_Q)

        # results
        self.voltage = []
        self.results_list = []
        self.Q_factor_voltage = []
        self.Q_factor_voltage_sigma = []

    def take_dirac_point(self, V_low, V_heigh, n_points, vna_repeats):
        
        for voltage in np.linspace(V_low, V_heigh, num = n_points):
            
            self.voltage.append(voltage)
            
            self.keith.apply_voltage(self.keith.smub, voltage)
            time.sleep(2)
            
            self.vna_measurment.clear_results()            
            vna_results = self.vna_measurment.multi_measurment(vna_repeats)
                        
            self.results_list.append([vna_results])
            mean_Q = self.vna_measurment.mean_Q_fit()
            self.Q_factor_voltage.append(mean_Q[0])
            self.Q_factor_voltage_sigma.append(mean_Q[1])
        
        return self.Q_factor_voltage, self.Q_factor_voltage_sigma,  self.results_list

    def plot_dirac_point(self):
        
        plt.errorbar(self.voltage, self.Q_factor_voltage, yerr = self.Q_factor_voltage_sigma, fmt = 'o', ecolor = 'black', elinewidth = 0.5, capsize = 5)
        plt.xlabel("Voltage / V")
        plt.ylabel("Q Factor")