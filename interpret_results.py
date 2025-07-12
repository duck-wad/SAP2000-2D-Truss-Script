import pandas as pd
import matplotlib.pyplot as plt
import os

def plot_save(plt, out_path, section, name):

    section_formatted = (section.replace(' ', '')).lower()
    name_formatted = name.replace(' ', '')
    full_name = '_'.join([section_formatted, name_formatted])
    full_name = full_name + '.pdf'
    full_path = out_path + os.sep + full_name
    plt.savefig(full_path, format='pdf')

def plot_mass_vs_deflection(df, sheet_name, out_path):

    mass = df['Module mass (kg)']
    deflection = df['Max vertical deflection for SLS (m)']
    
    plt.figure(figsize=(10,6))
    plt.scatter(deflection*1000, mass)
    plt.xlabel('SLS Deflection (mm)')
    plt.ylabel('Mass of module (kg)')
    plt.title(f'Module mass vs SLS deflection for "{sheet_name}" sections')
    plt.grid(True)
    plt.minorticks_on()
    plot_save(plt, out_path, sheet_name, 'mass vs deflection')

def plot_mass_vs_freq(df, sheet_name, out_path):

    mass = df['Module mass (kg)']
    frequency = df['Natural frequency (Hz)']
    
    plt.figure(figsize=(10,6))
    plt.scatter(frequency, mass)
    plt.xlabel('Natural frequency (Hz)')
    plt.ylabel('Mass of module (kg)')
    plt.title(f'Module mass vs natural frequency for "{sheet_name}" sections')
    plt.grid(True)
    plt.minorticks_on()
    plot_save(plt, out_path, sheet_name, 'mass vs natural frequency')

def interpret_results(file_path, sheets, folderpath):

    os.makedirs('./plots', exist_ok=True)
    out_path = folderpath + '/plots'
    
    dfs = []
    for sheet in sheets:
        dfs.append(pd.read_excel(file_path, sheet_name=sheet))
    
    for index, df in enumerate(dfs):
        plot_mass_vs_deflection(df, sheets[index], out_path)
        plot_mass_vs_freq(df, sheets[index], out_path)

def test():
    root_path = os.getcwd()
    file_path = root_path + '/output.xlsx'
    sheet_names = ['Box Box Box', 'Box Box Round']
    interpret_results(file_path, sheet_names, root_path)

test()