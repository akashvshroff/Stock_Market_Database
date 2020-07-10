import pandas as pd
import numpy as np
df = pd.DataFrame(np.arange(12).reshape(3, 4),
                  columns=['A', 'B', 'C', 'D'])
print(df)
df.drop([0], inplace=True)
print(df)
print(df['A'][0])
