import pandas as pd
df1 = pd.DataFrame({'key': ['foo', 'bar', 'baz', 'foo'],
                    'value': [1, 2, 3, 5],
                    'vlue': [1, 2, 3, 5]})
df2 = pd.DataFrame({'key': ['foo', 'bar', 'baz', 'foo'],
                    'value': [5, 6, 7, 8],
                    'vlue': [1, 2, 3, 5]})
x=pd.merge(df1,df2, on=['key','value'],
          suffixes=('_left', '_right'))
print(x)
