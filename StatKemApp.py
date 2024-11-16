import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import re
import io
import contextlib
from tkinter import *
from tkinter import messagebox
from termcolor import colored  # You can remove termcolor if not needed for GUI
from scipy.stats import ttest_ind

# Load q- and t-value tables
q_table = pd.read_excel("q_table.xlsx")
t_table = pd.read_excel("t_table.xlsx")

# Main App Class
class StatAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("StatKemApp")
        self.root.geometry("500x650")
        self.root.resizable(False, False)

        # Widgets for input fields
        Label(root, text="Enter the set (comma-separated):").pack()
        self.set1_entry = Entry(root, width=50)
        self.set1_entry.pack()

        Label(root, text="Enter the second set (comma-separated or 0 if there is no second set):").pack()
        self.set2_entry = Entry(root, width=50)
        self.set2_entry.pack()

        Label(root, text="Enter Confidence Level for Q-Test (90, 95, 99):").pack()
        self.nq_entry = Entry(root, width=10)
        self.nq_entry.pack()

        Label(root, text="Enter Confidence Level for T-Test (90, 95, 99):").pack()
        self.nt_entry = Entry(root, width=10)
        self.nt_entry.pack()

        Button(root, text="Analyze", command=self.check_and_run).pack(pady=10)
        self.output_text = Text(root, wrap=WORD, height=25, width=50)
        self.output_text.pack(pady=10)

    def check_errors(self, set1, set2, nq_inp, nt_inp):
        # Validation function
        res = True
        pattern = r"^(-?\d+(\.\d+)?)(,\s*-?\d+(\.\d+)?)*$"
        if not re.match(pattern, set1) or len(set1.split(',')) < 3:
            messagebox.showerror("Input Error", "Set 1 is in the wrong format")
            res = False
        if set2 != '0' and (not re.match(pattern, set2) or len(set2.split(',')) < 3):
            messagebox.showerror("Input Error", "Set 2 is in the wrong format")
            res = False
        if nq_inp not in ['90', '95', '99']:
            messagebox.showerror("Input Error", "Q-Test confidence level must be 90, 95, or 99")
            res = False
        if set2 != '0' and nt_inp not in ['90', '95', '99']:
            messagebox.showerror("Input Error", "T-Test confidence level must be 90, 95, or 99")
            res = False
        return res

    def std(self, sett):
        return round(np.nanstd(sett, ddof=1 if len(sett) <= 20 else 0), 4)

    def q_test(self, sett, suspect, N):
        sorted_set = sorted(set(sett))
        range_val = max(sett) - min(sett)
        q_calc = (sorted_set[1] - suspect) / range_val if suspect == min(sett) else (suspect - sorted_set[-2]) / range_val
        q_table_value = float(q_table[q_table['Measurements'] == len(sett)][f'Confidence level {N}'])
        return q_calc, q_table_value, q_calc > q_table_value

    def t_test(self, sett1, sett2, N):
        dof = len(sett1) + len(sett2) - 2
        t_calc = ttest_ind(sett1, sett2)
        t_table_value = float(t_table[t_table['Degrees of Freedom'] == dof][f'Confidence level {N}'])
        return t_calc.statistic, t_table_value, abs(t_calc.statistic) > t_table_value, t_calc.pvalue

    def anal_stat(self, set1, set2, Nq, Nt=None):
        set1 = [float(i) for i in set1.split(',')]
        set2 = [float(i) for i in set2.split(',')] if set2 != '0' else None
        Nq = int(Nq)
        Nt = int(Nt) if set2 else None

        with io.StringIO() as buf, contextlib.redirect_stdout(buf):
            print("Results:")
            print("Set 1:", set1)
            if set2:
                print("Set 2:", set2)

            print("Standard Deviation for Set 1:", self.std(set1))
            if set2:
                print("Standard Deviation for Set 2:", self.std(set2))

            # Q-Test for outliers
            print("\nQ-Test Results:")
            for suspect in [min(set1), max(set1)]:
                q_calc, q_val, is_outlier = self.q_test(set1, suspect, Nq)
                print(f"{suspect} {'is an outlier' if is_outlier else 'is not an outlier'} (q_calc: {q_calc:.3f}, q_table: {q_val})")
            if set2:
                for suspect in [min(set2), max(set2)]:
                    q_calc, q_val, is_outlier = self.q_test(set2, suspect, Nq)
                    print(f"{suspect} {'is an outlier' if is_outlier else 'is not an outlier'} (q_calc: {q_calc:.3f}, q_table: {q_val})")

            # T-Test between sets
            if set2:
                print("\nT-Test Result:")
                t_stat, t_val, is_significant, p_value = self.t_test(set1, set2, Nt)
                print(f"T-test result: {'Significant' if is_significant else 'Not significant'} (t_calc: {abs(t_stat):.4f}, t_table: {t_val})")
                print("P-value:", p_value)

            self.output_text.delete(1.0, END)
            self.output_text.insert(END, buf.getvalue())

        # Plot the boxplot
        self.plot_boxplot(set1, set2)

    def plot_boxplot(self, set1, set2):
        colors = ['#558AB5', "#F5AA36"] if set2 != [0.0] else ['#558AB5']
        plt.style.use('fivethirtyeight')
        fig, ax = plt.subplots()
        labels = ['Set 1', 'Set 2'] if set2 else ['Set 1']
        sets = [set1, set2] if set2 else [set1]
        bplot = ax.boxplot(sets, patch_artist=True, labels=labels, medianprops=dict(color="black", linewidth=1.5))
        ax.set_ylabel('Values')
        plt.title("Boxplot for Sets")
        for patch, color in zip(bplot['boxes'], colors):
          patch.set_facecolor(color)
        plt.show()

    def check_and_run(self):
        set1 = self.set1_entry.get()
        set2 = self.set2_entry.get()
        nq = self.nq_entry.get()
        nt = self.nt_entry.get() if set2 != '0' else None

        if self.check_errors(set1, set2, nq, nt):
            self.anal_stat(set1, set2, nq, nt)

# Main execution
if __name__ == "__main__":
    root = Tk()
    app = StatAnalyzerApp(root)
    root.mainloop()
