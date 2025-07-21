import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np


def plot_save(plt, out_path, section, name):

    section_formatted = (section.replace(" ", "")).lower()
    name_formatted = name.replace(" ", "")
    full_name = "_".join([section_formatted, name_formatted])
    full_name = full_name + ".pdf"
    full_path = out_path + os.sep + full_name
    plt.savefig(full_path, format="pdf")


def get_uls_indices(df):
    uls_result = df["Passed steel design check for ULS"].to_numpy()
    index_true = np.where(uls_result)[0]
    index_false = np.where(~uls_result)[0]
    return index_true, index_false


def determine_optimal_section(mass, harmonic, deflection):

    # normalize the harmonics, masses and deflections so they range from 0 to 1, where 1 is good and 0 is bad
    # for instance, a high harmonic is good so normalize up, whereas low mass is good so normalize down
    max_mass = np.max(mass)
    min_mass = np.min(mass)
    max_harmonic = np.max(harmonic)
    min_harmonic = np.min(harmonic)
    max_deflection = np.max(deflection)
    min_deflection = np.min(deflection)

    normalized_mass = 1.0 - (mass - min_mass) / (max_mass - min_mass)
    normalized_harmonic = (harmonic - min_harmonic) / (max_harmonic - min_harmonic)
    normalized_deflection = 1.0 - (deflection - min_deflection) / (
        max_deflection - min_deflection
    )

    # assign equal weighting to mass and harmonic
    mass_weighting = 2.0 / 4.0
    harmonic_weighting = 1.0 / 4.0
    deflection_weighting = 1.0 / 4.0
    score = (
        mass_weighting * normalized_mass
        + harmonic_weighting * normalized_harmonic
        + deflection_weighting * normalized_deflection
    )

    high_score_index = np.argmax(score)
    return high_score_index


def plot_mass_vs_deflection(
    mass_uls_passed,
    mass_uls_failed,
    deflection_uls_passed,
    deflection_uls_failed,
    optimal_mass,
    optimal_deflection,
    optimal_sections,
    sheet_name,
    out_path,
):

    plt.figure(figsize=(8, 6))
    plt.scatter(
        mass_uls_failed,
        deflection_uls_failed * 1000,
        label="ULS Failed",
        color="red",
        s=10,
    )
    plt.scatter(
        mass_uls_passed,
        deflection_uls_passed * 1000,
        label="ULS Passed",
        color="green",
        s=10,
    )
    plt.scatter(
        optimal_mass,
        optimal_deflection * 1000,
        label="Optimal Section Combination",
        color="cyan",
        s=20,
    )
    plt.ylabel("SLS Deflection [mm]")
    plt.xlabel("Mass of module [kg]")
    plt.title(f'Module mass vs SLS deflection for "{sheet_name}" sections')
    plt.grid(True)
    plt.minorticks_on()
    plt.legend()
    plt.annotate(
        optimal_sections,
        xy=(optimal_mass, optimal_deflection * 1000),
        xytext=(optimal_mass + 400, optimal_deflection * 1000 + 2),
        arrowprops=dict(facecolor="cyan", arrowstyle="->"),
        fontsize=9,
        bbox=dict(
            boxstyle="round,pad=0.3", edgecolor="cyan", facecolor="white", alpha=0.8
        ),
    )

    plot_save(plt, out_path, sheet_name, "mass vs deflection")


def plot_mass_vs_harmonic(
    mass_uls_passed,
    mass_uls_failed,
    harmonic_uls_passed,
    harmonic_uls_failed,
    optimal_mass,
    optimal_harmonic,
    optimal_sections,
    sheet_name,
    out_path,
):

    plt.figure(figsize=(8, 6))
    plt.scatter(
        mass_uls_failed, harmonic_uls_failed, label="ULS Failed", color="red", s=10
    )
    plt.scatter(
        mass_uls_passed, harmonic_uls_passed, label="ULS Passed", color="green", s=10
    )
    plt.scatter(
        optimal_mass,
        optimal_harmonic,
        label="Optimal Section Combination",
        color="cyan",
        s=20,
    )
    plt.ylabel("Occupied resonating harmonic")
    plt.xlabel("Mass of module [kg]")
    plt.title(
        f'Module mass vs occupied resonating harmonic for "{sheet_name}" sections'
    )
    plt.grid(True)
    plt.minorticks_on()
    plt.legend()
    plt.annotate(
        optimal_sections,
        xy=(optimal_mass, optimal_harmonic),
        xytext=(optimal_mass + 400, optimal_harmonic - 0.7),
        arrowprops=dict(facecolor="cyan", arrowstyle="->"),
        fontsize=9,
        bbox=dict(
            boxstyle="round,pad=0.3", edgecolor="cyan", facecolor="white", alpha=0.8
        ),
    )

    plot_save(plt, out_path, sheet_name, "mass vs occupied resonating harmonic")


# interpret_results can be run from main.py, or just from the run() function in this file
def interpret_results(file_path, sheets, folderpath):

    os.makedirs("./plots", exist_ok=True)
    out_path = folderpath + "/plots"

    dfs = []
    for sheet in sheets:
        dfs.append(pd.read_excel(file_path, sheet_name=sheet))

    # track the optimal deflections, harmonics, and masses for each section combination type

    for index, df in enumerate(dfs):

        top_chord = df["Top chord"].to_numpy()
        bottom_chord = df["Bottom chord"].to_numpy()
        web = df["Web members"].to_numpy()
        mass = df["Module mass (kg)"].to_numpy()
        # get the occupied case since it is critical
        harmonic = df["Resonating harmonic occupied"].to_numpy()
        deflection = df["Max vertical deflection for SLS (m)"].to_numpy()

        # get the indices of section combos that pass ULS and split into separate lists
        index_true, index_false = get_uls_indices(df)
        mass_uls_passed = mass[index_true]
        mass_uls_failed = mass[index_false]
        harmonic_uls_passed = harmonic[index_true]
        harmonic_uls_failed = harmonic[index_false]
        deflection_uls_passed = deflection[index_true]
        deflection_uls_failed = deflection[index_false]

        top_chord_uls_passed = top_chord[index_true]
        bottom_chord_uls_passed = bottom_chord[index_true]
        web_uls_passed = web[index_true]

        high_score_index = determine_optimal_section(
            mass_uls_passed, harmonic_uls_passed, deflection_uls_passed
        )
        optimal_mass = mass_uls_passed[high_score_index]
        optimal_harmonic = harmonic_uls_passed[high_score_index]
        optimal_deflection = deflection_uls_passed[high_score_index]
        optimal_top_chord = top_chord_uls_passed[high_score_index]
        optimal_bottom_chord = bottom_chord_uls_passed[high_score_index]
        optimal_web = web_uls_passed[high_score_index]
        optimal_sections = (
            "Top: "
            + optimal_top_chord
            + ", Bottom: "
            + optimal_bottom_chord
            + ", Web: "
            + optimal_web
        )
        # with line breaks for plotting
        optimal_sections_spaced = (
            "Top: "
            + optimal_top_chord
            + "\nBottom: "
            + optimal_bottom_chord
            + "\nWeb: "
            + optimal_web
        )

        plot_mass_vs_deflection(
            mass_uls_passed,
            mass_uls_failed,
            deflection_uls_passed,
            deflection_uls_failed,
            optimal_mass,
            optimal_deflection,
            optimal_sections_spaced,
            sheets[index],
            out_path,
        )
        plot_mass_vs_harmonic(
            mass_uls_passed,
            mass_uls_failed,
            harmonic_uls_passed,
            harmonic_uls_failed,
            optimal_mass,
            optimal_harmonic,
            optimal_sections_spaced,
            sheets[index],
            out_path,
        )


def run():
    root_path = os.getcwd()
    file_path = root_path + "/output.xlsx"
    sheet_names = ["Box Box Box", "Box Box Round"]
    interpret_results(file_path, sheet_names, root_path)


if __name__ == "__main__":
    run()
