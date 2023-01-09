import pandas as pd
import numpy as np

# creating data frame
df1 = pd.read_excel(r'C:/Users/p.zakeri/Desktop/second phase zed.xlsx', sheet_name='moaven')

mylist1 = []
count_row = df1.shape[0]
for i in range(0, count_row):
    # iloc[i] seperates row i and deleting index 0 is for removing c11 & ....
    rowlist1 = df1.iloc[i].tolist()
    del rowlist1[0]
    # fuzzification
    rowlist1 = [[0, 0, 0.25] if i == 1 else
                [0, 0.25, 0.5] if i == 2 else
                [0.25, 0.5, 0.75] if i == 3 else
                [0.5, 0.75, 1] if i == 4 else
                [0.75, 1, 1] if i == 5 else
                i for i in rowlist1]
    mylist1.append(rowlist1)


# fuzzy dematel first step creating arrays for l,m,u
X_l = np.zeros((count_row, 32))
X_m = np.zeros((count_row, 32))
X_u = np.zeros((count_row, 32))

# replacing with actual value
for i in range(0, count_row):
    for j in range(0, 32):
        a, b, c = mylist1[i][j]
        X_l[i, j] = a
        X_m[i, j] = b
        X_u[i, j] = c

S_l = np.sum(X_l, axis=0)
S_m = np.sum(X_m, axis=0)
S_u = np.sum(X_u, axis=0)

M_l = S_l / 18
M_m = S_m / 18
M_u = S_u / 18

defuzzied_results = (((M_u - M_l) + (M_m - M_l)) / 3) + M_l



dff = pd.DataFrame(defuzzied_results)
dff.to_excel('moaven.xlsx')
dff2 = pd.DataFrame(M_l)
dff2.to_excel('moaven_l.xlsx')
dff3 = pd.DataFrame(M_m)
dff3.to_excel('moaven_m.xlsx')
dff4 = pd.DataFrame(M_u)
dff4.to_excel('moaven_u.xlsx')