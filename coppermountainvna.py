#%matplotlib qt

import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit
import pickle
import time
from scipy.special import erf, erfc
from scipy.optimize import fmin, fsolve
from lmfit.models import LinearModel, SplitLorentzianModel, ConstantModel, PseudoVoigtModel, SkewedGaussianModel
from matplotlib.animation import FuncAnimation
import pep8

# Allows communication via COM interface
try:
    import win32com.client
except:
    print("You will first need to import the pywin32 extension")
    print("to get COM interface support.")
    print("Try http://sourceforge.net/projects/pywin32/files/ )?")
    input("\nPress Enter to Exit Program\n")
    exit()


class Measurment:

    def __init__(self, parameters=None):

        # initialise instrument
        if parameters is None:
            parameters = Measurment.request_parameters()

        self.parameters = parameters
        self.app = self.instantiate_COM_client(parameters[0])

        # initialise measurment parameters
        self.F = []
        self.Y = []

    def request_parameters():
        # Prompt for user's input

        instrlist = ['R54 or R140',
                     'TR1300, TR5048, or TR7530',
                     'S5048, S7530, Planar804, or Planar304',
                     'S8081']
        familylist = ['RVNA',
                      'TRVNA',
                      'S2VNA',
                      'S4VNA']
        print('\n', '0 - ', instrlist[0], '\n', '1 - ', instrlist[1],
              '\n', '2 - ', instrlist[2], '\n', '3 - ', instrlist[3], '\n')

        # choose the instrument
        instrument = familylist[int(
            input('Please select your instrument(only enter the first number):'))]

        # choose frequency type, 0 for Start/Stop Frequency, 1 for Center/Span Frequency
        use_center_and_span = int(input(
            '\nPlease enter whether 0 - Start/Stop Frequency \t 1-Center/Span Frequency:'))

        # power level
        power_level_dbm = float(input('\nPlease enter power level(dbm):'))

        # fstart=400e6 or center, as per above, in Hz
        f1_hz = int(input('\nPlease enter start/center frequency (Hz):'))

        # fstop=600e6 or span, as per above, in Hz
        f2_hz = int(input('\nPlease enter stop/span frequency (Hz):'))

        # number of measurement points
        num_points = int(input('\nPlease enter number of measurement points:'))
        parameter = input(
            '\nPlease enter the parameter (e.g. S11, S21, S12, S22, etc):')
        # "S21", "S11", "S12", etc. R54/140 must use
        # "S11"; TR devices must use "S11" or "S21";
        #  Ports 3 and 4 available for S8081 only
        
        #IF Bandwidth
        IF_Bandwidth = float(input('Please enter IF bandwidth:'))

        #"mlog" or "phase" or"smith chart"
        format = input('\nPlease enter the format (e.g. mlog, phase, smith):')

        return instrument, use_center_and_span, power_level_dbm, f1_hz, f2_hz, num_points, IF_Bandwidth, parameter, format

    def instantiate_COM_client(self, instrument=None):
        # Instantiate COM client

        if instrument is None:
            instrument = self.parameters[0]

        try:
            app = win32com.client.Dispatch(instrument + ".application")
        except:
            print("Error establishing COM server connection to " + instrument + ".")
            print("Check that the VNA application COM server was registered")
            print("at the time of software installation.")
            print("This is described in the VNA programming manual.")
            input("\nPress Enter to Exit Program\n")
            exit()

        # Wait up to 20 seconds for instrument to be ready
        if app.Ready == 0:
            print("Instrument not ready! Waiting...")
            for k in range(1, 21):
                time.sleep(1)
                if app.Ready != 0:
                    break
                print("%d" % k)

        # If the software is still not ready, cancel the program
        if app.Ready == 0:
            print("Error, timeout waiting for instrument to be ready.")
            print("Check that VNA is powered on and connected to PC.")
            print("The status Ready should appear in the lower right")
            print("corner of the VNA application window.")
            input("\nPress Enter to Exit Program\n")
            exit()
        else:
            print("Instrument ready! Continuing...")

        #Get and echo the instrument name, serial number, etc.#
        print(app.name)

        return app

    def configure_measurement(self, app=None, parameters=None):

        if app is None:
            app = self.app

        app.scpi.system.preset()

        if parameters is None:
            parameters = self.parameters

        instrument, use_center_and_span, power_level_dbm, f1_hz, f2_hz, num_points, IF_bandwidth, parameter, format = parameters

        # Configure the stimulus
        if use_center_and_span == 1:
            app.scpi.GetSENSe(1).frequency.center = f1_hz
            app.scpi.GetSENSe(1).frequency.span = f2_hz
        else:
            app.scpi.GetSENSe(1).frequency.start = f1_hz
            app.scpi.GetSENSe(1).frequency.stop = f2_hz

        app.scpi.GetSENSe(1).sweep.points = num_points
        app.scpi.GetSENSe(1).bandwidth.resolution = IF_bandwidth

        if instrument[0] != "R":
            app.scpi.GetSOURce(1).power.level.immediate.amplitude = power_level_dbm

        # Configure the measurement
        app.scpi.GetCALCulate(1).GetPARameter(1).define = parameter
        app.scpi.GetCALCulate(1).GetPARameter(1).select()
        app.scpi.GetCALCulate(1).selected.format = format
        app.scpi.trigger.sequence.source = "bus"

    def measure_F_and_Y(self, app=None):

        if app is None:
            app = self.app

        # Execute the measurement
        app.scpi.trigger.sequence.single()

        app.scpi.GetCALCulate(1).GetPARameter(1).select()
        Y = app.scpi.GetCALCulate(1).selected.data.Fdata

        # Discard complex-valued points
        self.Y = Y[0::2]

        self.F = app.scpi.GetSENSe(1).frequency.data

        return self.F, self.Y

    def take_mesurment(self, app=None, parameters=None):

        if app is None:
            app = self.app

        if parameters is None:
            parameters = self.parameters

        self.configure_measurement(app, parameters)
        self.F, self.Y = self.measure_F_and_Y(app)

        return self.F, self.Y

    def peak_measurment(self, width, number_interations=2):

        def new_parameters(F, Y):
            # takes old parameters and returns new parameters around peak

            peak_freq = F[Y.index(max(Y))]
            new_f1 = peak_freq - width/2
            new_f2 = peak_freq + width/2

            self.parameters[3] = new_f1
            self.parameters[4] = new_f2

        for i in range(number_interations):

            self.configure_measurement()
            self.F, self.Y = self.measure_F_and_Y()
            new_parameters(self.F, self.Y)

        return self.parameters

    def plot_data(self, F=None, Y=None):

        if F is None:
            F = self.F

        if Y is None:
            Y = self.Y

        plt.plot(F, Y, label="data")
        plt.show()


class QFactor:

    def __init__(self, F, Y):

        # initialise curve data
        self.F = F
        self.Y = Y
        
        # initialise 3db peak variables 
        self.Qf_db = None
        self.maximum_freq_db = None
        
        # initialise lorentzian fit variables 
        self.Q_fit = None
        self.peak_frequency_fit = None
        self.fit = None
        self.r_squared_fit = None   
        
    def db_Qf(self):
        # calculate 3db Q

        F = self.F
        Y = self.Y

        Y_at_low_index = min(
            range(len(Y[:Y.index(max(Y))])), key=lambda i: abs(Y[i]-max(Y) + 3))
        Y_at_heigh_index = min(range(len(Y[Y.index(max(Y)):])), key=lambda i: abs(
            Y[i+Y.index(max(Y))]-max(Y) + 3)) + Y.index(max(Y))
        self.Qf_db = F[Y.index(max(Y))]/(F[Y_at_heigh_index] - F[Y_at_low_index])

        return self.Qf_db

    def peak_freq(self):
        # calculate 3db peak frequency

        F = self.F
        Y = self.Y

        self.maximum_freq_db = F[Y.index(max(Y))]

        return self.maximum_freq_db

    def fit_lorentzian(self):
        # calculate Q and peak frequency from lorentzian

        F = self.F
        Y = self.Y

        def lorentzian(x, a, mu, sigma):
            return (a*sigma)/(np.pi*((x-mu)**2 + sigma**2))

        def r_squared(xdata, ydata, variables, function):
            xdata = np.array(xdata)
            residuals = ydata - function(xdata, *variables)
            ss_res = np.sum(residuals**2)
            ss_tot = np.sum((ydata-np.mean(ydata))**2)
            r_squared = 1 - (ss_res / ss_tot)
            return r_squared

        V = [10**(i/10) for i in Y]

        mu_guess = F[Y.index(max(Y))]
        sigma_guess = 2 * \
            min(range(len(Y[:Y.index(max(Y))])),
                key=lambda i: abs(Y[i]-max(Y)+3))
        amplitude_guess = max(V)*sigma_guess

        self.fit = curve_fit(lorentzian, F, V, p0=[
                             amplitude_guess, mu_guess, sigma_guess])
        self.r_squared_fit = r_squared(F, Y, self.fit[0], lorentzian)
        self.peak_frequency_fit = self.fit[0][1]
        self.Q_fit = self.fit[0][1]/(2*self.fit[0][2])

        return self.Q_fit, self.peak_frequency_fit, self.fit, self.r_squared_fit


class MultiPeakAnalysis:

    def __init__(self, **kwargs):

        self.F_list = []
        self.Y_list = []
        self.time_list = []
        self.Q_list_db = []
        self.peak_list_db = []
        self.Q_list_fit = []
        self.peak_list_fit = []

        for key in self.__dict__:
            value = kwargs.get(key, self.__dict__[key])
            setattr(self, key, value)

    def calculate_Q_db(self, overwrite=False):

        if not self.Q_list_db or overwrite is True:
            if overwrite is True:
                self.Q_list_db.clear()

            for i in range(len(self.F_list)):
                Qf = QFactor(self.F_list[i], self.Y_list[i])
                self.Q_list_db.append(Qf.db_Qf())
        else:
            print("Q_db already calculated")

        return self.Q_list_db

    def calculate_Q_fit(self, overwrite=False):

        if not self.Q_list_fit or overwrite is True:
            if overwrite is True:
                self.Q_list_fit.clear()

            for i in range(len(self.F_list)):
                Q, peak, * \
                    fit = QFactor(self.F_list[i],
                                  self.Y_list[i]).fit_lorentzian()
                self.Q_list_fit.append(Q)

        else:
            print("Q_fit already calculated")

        return self.Q_list_fit

    def calculate_peak_db(self, overwrite=False):

        if not self.peak_list_db or overwrite is True:
            if overwrite is True:
                self.peak_list_db.clear()

            for i in range(len(self.F_list)):
                Qf = QFactor(self.F_list[i], self.Y_list[i])
                self.peak_list_db.append(Qf.peak_freq())

        else:
            print("peak_db already calculated")

        return self.peak_list_db

    def calculate_peak_fit(self, overwrite=False):

        if not self.peak_list_fit or overwrite is True:
            if overwrite is True:
                self.peak_list_fit.clear()

            for i in range(len(self.F_list)):
                Q, peak, * \
                    fit = QFactor(self.F_list[i],
                                  self.Y_list[i]).fit_lorentzian()
                self.peak_list_fit.append(peak)

        else:
            print("peak_fit already calculated")

        return self.peak_list_fit

    def _mean_and_std_list(self, list_):

        mean = np.mean(list_)
        sigma = np.std(list_)

        return mean, sigma

    def mean_Q_db(self):

        if not self.Q_list_db:
            self.calculate_Q_db()

        self.mean_Q_db_result, self.sigma_Q_db = self._mean_and_std_list(
            self.Q_list_db)
        return self.mean_Q_db_result, self.sigma_Q_db

    def mean_peak_freq_db(self):

        if not self.peak_list_db:
            self.calculate_peak_db()

        self.mean_peak_freq_db_result, self.sigma_peak_freq_db = self._mean_and_std_list(
            self.peak_list_db)
        return self.mean_peak_freq_db_result, self.sigma_peak_freq_db

    def mean_Q_fit(self):

        if not self.Q_list_fit:
            self.calculate_Q_fit()

        self.mean_Q_fit_result, self.sigma_Q_fit = self._mean_and_std_list(
            self.Q_list_fit)
        return self.mean_Q_fit_result, self.sigma_Q_fit

    def mean_peak_freq_fit(self):

        if not self.peak_list_fit:
            self.calculate_peak_fit()

        self.mean_peak_freq_fit_result, self.sigma_peak_freq_fit = self._mean_and_std_list(
            self.peak_list_fit)
        return self.mean_peak_freq_fit_result, self.sigma_peak_freq_fit

    def plot_Q_db_time(self):

        if not self.Q_list_db:
            self.calculate_Q_db()

        plt.plot(self.time_list, self.Q_list_db)
        plt.xlabel("time / s")
        plt.ylabel("3db Q")
        plt.show()

    def plot_peak_db_time(self):

        if not self.peak_list_db:
            self.calculate_peak_db()

        plt.plot(self.time_list, self.peak_list_db)
        plt.xlabel("time / s")
        plt.ylabel("3db peak frequency / Hz")
        plt.show()

    def plot_Q_fit_time(self):

        if not self.Q_list_fit:
            self.calculate_Q_fit()

        plt.plot(self.time_list, self.Q_list_fit)
        plt.xlabel("time / s")
        plt.ylabel("Q")
        plt.show()

    def plot_peak_fit_time(self):

        if not self.peak_list_fit:
            self.calculate_peak_fit()

        plt.plot(self.time_list, self.peak_list_fit)
        plt.xlabel("time / s")
        plt.ylabel("peak frequency / Hz")
        plt.show()

    def calculate_sheet_resistance(self, other, t_s=0.52*10**-3, t_s_sigma=0.005*10**-3, epsilon_s=4.5, epsilon_s_sigma=0.05):

        # calculate sheet resistance using control measurment as other
        epsilon_0 = 8.8541878128*10**-12

        f_sample, f_sample_sigma = self.mean_peak_freq_fit()
        Q_sample, Q_sample_sigma = self.mean_Q_fit()
        f_no_sample, f_no_sample_sigma = other.mean_peak_freq_fit()
        Q_no_sample, Q_no_sample_sigma = other.mean_Q_fit()

        delta_f = f_no_sample - f_sample
        delta_f_sigma = np.sqrt(f_no_sample_sigma**2 + f_sample_sigma**2)

        delta_Q = 1/Q_sample - 1/Q_no_sample
        delta_Q_sigma = np.sqrt(
            (Q_no_sample_sigma/Q_no_sample**2)**2 + (Q_sample_sigma/Q_sample**2)**2)

        R_s = delta_f/(np.pi*(f_no_sample**2)*epsilon_0 *
                       t_s*(epsilon_s - 1)*delta_Q)
        R_s_sigma = R_s*np.sqrt((delta_f_sigma/delta_f)**2 + (delta_Q_sigma/delta_Q)**2 + (
            2*f_no_sample_sigma/f_no_sample)**2 + (t_s_sigma/t_s)**2 + (epsilon_s_sigma/epsilon_s)**2)
        
        Z_c = delta_f/(np.pi*(f_no_sample**2)*epsilon_0*t_s*(epsilon_s - 1))
        
        return R_s, R_s_sigma, Z_c
    
    def sheet_resistance_time(self, other, plot = False, t_s=0.52*10**-3, t_s_sigma=0.005*10**-3, epsilon_s=4.5, epsilon_s_sigma=0.05):
        
        # calculate sheet resistance using control measurment as other
        epsilon_0 = 8.8541878128*10**-12
        
        R_s_list = [] 
        
        f_sample_list = self.peak_list_fit
        Q_sample_list = self.Q_list_fit
        f_no_sample, f_no_sample_sigma = other.mean_peak_freq_fit()
        Q_no_sample, Q_no_sample_sigma = other.mean_Q_fit()
        
        for f, Q in zip(f_sample_list, Q_sample_list):
            delta_f = f_no_sample - f
            delta_Q = 1/Q - 1/Q_no_sample
            R_s = delta_f/(np.pi*(f**2)*epsilon_0 *t_s*(epsilon_s - 1)*delta_Q)
            R_s_list.append(R_s)
            
        if plot:
            plt.plot(self.time_list, R_s_list)
            plt.xlabel("time / s")
            plt.ylabel("Sheet Reistance / ohm/sq")
            plt.show()
        
        return R_s_list


class MultiMeasurment(Measurment, MultiPeakAnalysis):

    def __init__(self, parameters, peak_width=None, compute_Q_db=False, compute_Q_fit=False, plot_Q=False, time_delay=0, peak_interations=2):

        self.plot_Q = plot_Q
        self.no_interations = 0
        self.time_delay = time_delay
        self.compute_Q_db = compute_Q_db
        self.compute_Q_fit = compute_Q_fit

        # initialise lists of Q_list_db, peak_list_db, Q_list_fit, peak_list_fit, F_list, Y_list, time_list
        MultiPeakAnalysis.__init__(self)

        # initialise parameters
        Measurment.__init__(self, parameters)

        if peak_width is not None:
            self.peak_measurment(peak_width, peak_interations)

        self.configure_measurement()

        self.start_time = time.time()

    def __single_measure(self):

        F, Y = self.measure_F_and_Y()
        self.time_list.append(time.time()-self.start_time)
        self.F_list.append(F)
        self.Y_list.append(Y)

        if self.compute_Q_db is True:
            Q = QFactor(F, Y)
            self.Q_list_db.append(Q.db_Qf())
            self.peak_list_db.append(Q.peak_freq())

            if self.plot_Q is True and self.compute_Q_fit is False:
                plt.plot(self.time_list, self.Q_list_db)
                plt.show()

        if self.compute_Q_fit is True:
            Q, peak, *fit = QFactor(F, Y).fit_lorentzian()
            self.Q_list_fit.append(Q)
            self.peak_list_fit.append(peak)

            if self.plot_Q is True:
                plt.plot(self.time_list, self.Q_list_fit)
                plt.show()

        time.sleep(self.time_delay)

    def multi_measurment(self, no_interations):

        self.no_interations = no_interations

        for i in range(self.no_interations):
            self.__single_measure()

        return self.F_list, self.Y_list, self.Q_list_db, self.peak_list_db, self.Q_list_fit, self.peak_list_fit, self.time_list

    def time_measurment(self, time_length = None):

        try:
            print("loop started! Use keyboard interrupt to stop")
            
            if time_length is not None:
                t_end = time.time() + time_length
            
            while time.time() < t_end if time_length is not None else True:
                self.__single_measure()
                self.no_interations += 1

        except KeyboardInterrupt:
            print("loop interupted!")
            pass

        return self.F_list, self.Y_list, self.Q_list_db, self.peak_list_db, self.Q_list_fit, self.peak_list_fit, self.time_list

    def clear_results(self):
        
        to_clear = [self.F_list, self.Y_list, self.Q_list_db, self.peak_list_db, self.Q_list_fit, self.peak_list_fit, self.time_list]
        
        for list_ in to_clear:
            list_.clear()

class RealTimeMeasurmentPlot(Measurment, MultiPeakAnalysis):

    def __init__(self, parameters, peak_width=None, peak_interations=2, Q_fit=False):

        # calculate 3 db Q if False or fit Q if True. Defult False
        self.fit = Q_fit

        # initialise parameters
        super().__init__(parameters)

        if peak_width is not None:
            self.peak_measurment(peak_width, peak_interations)

        self.configure_measurement()

        # initialise variables
        self.F_list = []
        self.Y_list = []
        self.Q_list_db = []
        self.peak_list_db = []
        self.Q_list_fit = []
        self.peak_list_fit = []
        self.time_list = []

        # initialise plot
        fig, (ax_1, ax_2) = plt.subplots(2)
        self.fig = fig
        self.ax_1 = ax_1
        self.ax_2 = ax_2
        self.ln0_1, = ax_1.plot([], [])
        self.ln0_2, = ax_2.plot([], [])

        # start time
        self.start_time = time.time()

    def init_plot(self):
        self.ln0_1.set_data([], [])
        self.ln0_2.set_data([], [])
        return self.ln0_1, self.ln0_2

    def update_plot(self, i):
        
        F, Y = self.measure_F_and_Y()

        self.time_list.append(time.time()-self.start_time)
        self.F_list.append(F)
        self.Y_list.append(Y)

        self.ax_1.cla()
        self.ax_2.cla()

        if self.fit is False:
            Qf = QFactor(F, Y)
            self.Q_list_db.append(Qf.db_Qf())
            self.peak_list_db.append(Qf.peak_freq())
            self.ax_1.plot(self.time_list, self.Q_list_db)
            self.ax_2.plot(self.time_list, self.peak_list_db)

        else:
            Q, peak, *fit = QFactor(F, Y).fit_lorentzian()
            self.Q_list_fit.append(Q)
            self.peak_list_fit.append(peak)
            self.ax_1.plot(self.time_list, self.Q_list_fit)
            self.ax_2.plot(self.time_list, self.peak_list_fit)

        self.ax_1.set_xlabel('time / s')
        self.ax_1.set_ylabel('Q factor')
        self.ax_2.set_xlabel('time / s')
        self.ax_2.set_ylabel('peak frequency / Hz')

        return self.ln0_1, self.ln0_2

    def real_time_plot(self, time_interval=1000):
        self.anim = FuncAnimation(fig=self.fig, func=self.update_plot,
                                  init_func=self.init_plot, interval=time_interval, blit=False)


class CalculateSheetResistance:
    
    def __init__(self, sample, no_sample):
        
        self.sample = sample
        self.no_sample = no_sample
        
    def calculate_sheet_resistance(self, t_s=0.52*10**-3, t_s_sigma=0.005*10**-3, epsilon_s=4.5, epsilon_s_sigma=0.05):

        # calculate sheet resistance using control measurment as other
        epsilon_0 = 8.8541878128*10**-12

        f_sample, f_sample_sigma = self.sample.mean_peak_freq_fit()
        Q_sample, Q_sample_sigma = self.sample.mean_Q_fit()
        f_no_sample, f_no_sample_sigma = self.no_sample.mean_peak_freq_fit()
        Q_no_sample, Q_no_sample_sigma = self.no_sample.mean_Q_fit()

        delta_f = f_no_sample - f_sample
        delta_f_sigma = np.sqrt(f_no_sample_sigma**2 + f_sample_sigma**2)

        delta_Q = 1/Q_sample - 1/Q_no_sample
        delta_Q_sigma = np.sqrt(
            (Q_no_sample_sigma/Q_no_sample**2)**2 + (Q_sample_sigma/Q_sample**2)**2)

        R_s = delta_f/(np.pi*(f_no_sample**2)*epsilon_0 *
                       t_s*(epsilon_s - 1)*delta_Q)
        R_s_sigma = R_s*np.sqrt((delta_f_sigma/delta_f)**2 + (delta_Q_sigma/delta_Q)**2 + (
            2*f_no_sample_sigma/f_no_sample)**2 + (t_s_sigma/t_s)**2 + (epsilon_s_sigma/epsilon_s)**2)
        
        Z_c = delta_f/(np.pi*(f_no_sample**2)*epsilon_0*t_s*(epsilon_s - 1))
        
        return R_s, R_s_sigma, Z_c



class PickleObject:

    @staticmethod
    def save_object(obj, filename):
        try:
            with open(filename, "wb") as f:
                pickle.dump(obj, f, protocol=pickle.HIGHEST_PROTOCOL)
        except Exception as ex:
            print("Error during pickling object (Possibly unsupported):", ex)

    @staticmethod
    def load_object(filename):
        try:
            with open(filename, "rb") as f:
                return pickle.load(f)
        except Exception as ex:
            print("Error during unpickling object (Possibly unsupported):", ex)

    @staticmethod
    def save_object_no_app(obj, filename):
        del obj.app
        PickleObject.save_object(obj, filename)