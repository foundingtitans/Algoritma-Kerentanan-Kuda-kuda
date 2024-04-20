import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import lognorm
from scipy.stats import norm
from scipy.optimize import minimize
import os
import math

df = pd.read_csv("C:/Users/febyf/Documents/TGA FEBY/SAP HASIL RUNNING/TRIPLE FINK/CSV/Rekap Data/DATA KEAMANAN TUNGGAL")

# Function to perform optimization and plotting for each damage state
def plot_fragility_curve(IM, damaged_roof, total_roof, label, color):
    # Convert data to binomial
    y = damaged_roof / total_roof     #y itu probabilitas awal (pi)

    # Define the log-likelihood function
    def neg_loglik(theta):
        mu = theta[0]
        sigma = theta[1]
        prob = norm.cdf((np.log(IM) - mu) / sigma) #(P)
        ll = y * np.log(prob) + (1 - y) * np.log(1 - prob) #rumus likelihood
        return -np.sum(ll)

    # Perform optimization to estimate the parameters
    theta_start = np.array([1, 1])
    res = minimize(neg_loglik, theta_start, method='Nelder-Mead')
    print(res)
    mu_hat, sigma_hat = res.x #didapatkan nilai mean dan std yang ideal dari ll maksimum

    # Generate IM values for fragility curve
    IM_range = np.linspace(1, 50, 50)  #untuk sumbu x

    # Calculate fragility curve
    p_collapse_range = norm.cdf((np.log(IM_range) - mu_hat) / sigma_hat)   #utk bikin garis prob akhir kurva

    # Plotting the fragility curve
    plt.plot(IM_range, p_collapse_range, label=f'{label} (mu={mu_hat:.2f}, sigma={sigma_hat:.2f})', color=color)
    return IM_range, p_collapse_range, mu_hat, sigma_hat

    ####################################################################################################################################################################

Kecepatan = df['Kecepatan Angin'].unique()

total_roof = df['Kecepatan Angin'].value_counts().sort_values().values

damaged_roof = []
for speed in Kecepatan:
    # Filter the DataFrame for the current speed
    filtered_data = df[df['Kecepatan Angin'] == speed]
    
    # Check if the current speed appears more than once
    count_tidak_aman = (filtered_data['Aman Terhadap Penampang'] == 'Tidak Aman').sum()

    damaged_roof.append(count_tidak_aman)  

# Given data for each damage state
IM_DS1 = np.array(Kecepatan)
damaged_roof = np.array(damaged_roof)
total_roof = np.array(total_roof)
print(damaged_roof)

# Plotting
plt.figure(figsize=(10, 6))
IM, prop, mu, sigma = plot_fragility_curve(IM_DS1, damaged_roof, total_roof, '', 'Green')

data ={
    'Kecepatan Angin': IM,
    'Kurva Kerapuhan' : prop
} 

data = pd.DataFrame(data)
data.to_csv('C:/Users/febyf/Documents/TGA FEBY/MLE/EXCEL/Sebelum ANN/CDF Triple Fink.csv', index = None)

plt.ylim(0,1)
plt.title('Double Fink Vulnerability Curves')
plt.xlabel('Kecepatan Angin')
plt.ylabel('Percentage (%)')
plt.legend()
plt.grid(True)
plt.show()