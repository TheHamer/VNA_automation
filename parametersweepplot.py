import coppermountainvna
from coppermountainvna import MultiMeasurment
import numpy as np
import matplotlib.pyplot as plt


class PlotUncertaintyWithPoints:
    
    def __init__(self):
        
        self.mean_Q = []
        self.uncertainty = []
        self.points = []
        
    def get_points(self, start, end, number):
        
        for points_float in np.linspace(start, end, number):
            points = int(points_float)
            data = coppermountainvna.PickleObject.load_object(f"24_10_2022_repeats_100_no_sample_IFBandwidth_10k_points_{points}_width_40000000_number_points_long_run")
            mean_Q = data.mean_Q_fit()
            self.points.append(points)
            self.mean_Q.append(mean_Q[0])
            self.uncertainty.append(mean_Q[1])
            
    def plot_data(self):
        
        plt.plot(self.points, self.uncertainty)
        plt.xlabel("number of points")
        plt.ylabel("uncertainty")
        
        
class PlotUncertaintyWithWidth:
    
    def __init__(self):
        
        self.mean_Q = []
        self.uncertainty = []
        self.width = []
        
    def get_points(self, start, end, number):
        
        for width_float in np.linspace(start, end, number):
            width = int(width_float)
            data = coppermountainvna.PickleObject.load_object(f"19_10_2022_repeats_100_no_sample_IFBandwidth_10k_points_5000_width_{width}_long_run")
            mean_Q = data.mean_Q_fit()
            self.width.append(width)
            self.mean_Q.append(mean_Q[0])
            self.uncertainty.append(mean_Q[1])
            
    def plot_data(self):
        
        plt.plot(self.width, self.uncertainty)
        plt.xlabel("width / Hz")
        plt.ylabel("uncertainty")

class PlotUncertaintyWithIFBandwidth:
    
    def __init__(self):
        
        self.mean_Q = []
        self.uncertainty = []
        self.width = []
        
    def get_points(self, start, end, number):
        
        for IF_bandwidth_float in np.linspace(start, end, number):
            IF_bandwidth = int(IF_bandwidth_float)
            data = coppermountainvna.PickleObject.load_object(f"24_10_2022_repeats_100_no_sample_IFBandwidth_{IF_bandwidth}_points_5000_width_40000000_number_points_long_run")
            mean_Q = data.mean_Q_fit()
            self.width.append(IF_bandwidth)
            self.mean_Q.append(mean_Q[0])
            self.uncertainty.append(mean_Q[1])
            
    def plot_data(self):
        
        plt.plot(self.width, self.uncertainty)
        plt.xlabel("IF Bandwidth / Hz")
        plt.ylabel("Uncertainty")