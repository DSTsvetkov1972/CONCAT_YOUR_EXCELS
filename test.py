import pandas as pd

df =pd.DataFrame([{1:'\n',2:777}])
df.to_csv('a.csv')
df = df.map(lambda x: str(x).replace('\n','\\n'))
df.to_csv('b.csv')
