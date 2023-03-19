import seaborn as sns
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import warnings
from sklearn.metrics import mean_squared_error
from sklearn.model_selection import train_test_split
import xgboost as xgb
warnings.filterwarnings("ignore")
diamonds = pd.read_csv('DataForXGBoost.csv')
diamonds.head()

X = diamonds.drop('Дата снятия', axis=1)
y=diamonds[['Дата снятия']]
# Extract text features
cats = X.select_dtypes(exclude=np.number).columns.tolist()
# Convert to Pandas category
for col in cats:
    X[col] = X[col].astype('category')
# print(X.dtypes)
# Split the data
X_train, X_test, y_train, y_test = train_test_split(X, y, random_state=1)
# Create regression matrices
dtrain_reg = xgb.DMatrix(X_train, y_train, enable_categorical=True)
dtest_reg = xgb.DMatrix(X_test, y_test, enable_categorical=True)
params = {"objective": "reg:squarederror", "tree_method": "gpu_hist"}
evals = [(dtrain_reg, "train"),(dtest_reg, "validation") ]
n = 1000
model = xgb.train(
    params=params,
    dtrain=dtrain_reg,
    num_boost_round=n,
    evals=evals,
    verbose_eval=50,
    # Activate early stopping
    early_stopping_rounds=50
)

preds = model.predict(dtest_reg)
rmse = mean_squared_error(y_test, preds, squared=False)

# writer = pd.ExcelWriter('test.xlsx')
preds = [round(x) for x in preds]
with pd.ExcelWriter('test.xlsx') as writer: 
    pd.DataFrame(X_test).to_excel(writer ,sheet_name='X_test')
    pd.DataFrame(preds).to_excel(writer,sheet_name='preds')
    pd.DataFrame(y_test).to_excel(writer,sheet_name='y_test')
    
    
print(f"RMSE of the base model: {rmse:.3f}")
params = {"objective": "reg:squarederror", "tree_method": "gpu_hist"}

results = xgb.cv(
    params, dtrain_reg,
    num_boost_round=n,
    nfold=5,
    verbose_eval=50,
    early_stopping_rounds=20
)
print(results.head())
best_rmse = results['test-rmse-mean'].min()
print(best_rmse)
