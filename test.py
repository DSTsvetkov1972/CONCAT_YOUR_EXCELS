import pandas as pd
list = [('A',1.0,5.0),('B',2,7.0)]
df = pd.DataFrame(list)
df[1] = df[2].apply(int)
print(df)
