# ============================================================================
# ANALYSE DE RÉGRESSION - MONEY LAUNDERING RISK SCORE
# ============================================================================

# 1. IMPORTATION DES BIBLIOTHÈQUES
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.tree import DecisionTreeRegressor
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
import warnings
warnings.filterwarnings("ignore")

# 2. CHARGEMENT DES DONNÉES
print("Chargement des données...")
df = pd.read_csv('/content/drive/MyDrive/DM_ML/BigBlackMoneyDataset.csv')
print(f"Dataset chargé: {df.shape[0]} lignes, {df.shape[1]} colonnes")
print(df.head())

# 3. PRÉPARATION ET INGÉNIERIE DES CARACTÉRISTIQUES
print("\nPréparation des données...")
df['Date of Transaction'] = pd.to_datetime(df['Date of Transaction'])
df['transaction_year'] = df['Date of Transaction'].dt.year
df['transaction_month'] = df['Date of Transaction'].dt.month
df['transaction_day'] = df['Date of Transaction'].dt.day
df['transaction_hour'] = df['Date of Transaction'].dt.hour
df = df.drop('Date of Transaction', axis=1)

# Séparation cible/prédicteurs
y = df['Money Laundering Risk Score']
X = df.drop('Money Laundering Risk Score', axis=1)

# Encodage des variables catégorielles
categorical_features = X.select_dtypes(include=['object', 'bool']).columns
numerical_features = X.select_dtypes(include=['int64', 'float64']).columns

X_categorical = pd.get_dummies(X[categorical_features], drop_first=True)
X_numerical = X[numerical_features]
X = pd.concat([X_numerical, X_categorical], axis=1)

print(f"Shape of X after one-hot encoding: {X.shape}")
print(f"Shape of y: {y.shape}")

# 4. DIVISION TRAIN/TEST
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
print(f"\nDivision des données:")
print(f"X_train: {X_train.shape}, X_test: {X_test.shape}")
print(f"y_train: {y_train.shape}, y_test: {y_test.shape}")

# 5. FONCTION D'ÉVALUATION DES MODÈLES
def evaluate_model(model, X_train, X_test, y_train, y_test, model_name):
    print(f"\n--- {model_name} ---")
    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)
    
    mae = mean_absolute_error(y_test, y_pred)
    mse = mean_squared_error(y_test, y_pred)
    rmse = np.sqrt(mse)
    r2 = r2_score(y_test, y_pred)
    
    print(f"Mean Absolute Error (MAE): {mae:.4f}")
    print(f"Mean Squared Error (MSE): {mse:.4f}")
    print(f"Root Mean Squared Error (RMSE): {rmse:.4f}")
    print(f"R-squared (R²): {r2:.4f}")
    
    # Visualisation
    plt.figure(figsize=(10, 7))
    plt.scatter(y_test, y_pred, alpha=0.5)
    plt.plot([y_test.min(), y_test.max()], [y_test.min(), y_test.max()], 'r--', lw=2)
    plt.xlabel('Valeurs réelles')
    plt.ylabel('Prédictions')
    plt.title(f'{model_name} - Prédictions vs Réel')
    plt.show()
    
    return {'MAE': mae, 'MSE': mse, 'RMSE': rmse, 'R2': r2}

# 6. ENTRAÎNEMENT ET ÉVALUATION DES MODÈLES
results = {}

# Régression Linéaire
lr = LinearRegression()
results['Linear Regression'] = evaluate_model(lr, X_train, X_test, y_train, y_test, 'Modèle de Régression Linéaire')

# Arbre de Décision
dt = DecisionTreeRegressor(random_state=42)
results['Decision Tree'] = evaluate_model(dt, X_train, X_test, y_train, y_test, 'Arbre de Décision')

# Forêt Aléatoire
rf = RandomForestRegressor(n_estimators=100, random_state=42)
results['Random Forest'] = evaluate_model(rf, X_train, X_test, y_train, y_test, 'Forêt Aléatoire')

# Gradient Boosting
gb = GradientBoostingRegressor(random_state=42)
results['Gradient Boosting'] = evaluate_model(gb, X_train, X_test, y_train, y_test, 'Gradient Boosting')

# 7. COMPARAISON DES MODÈLES
print("\n" + "="*60)
print("COMPARAISON DES PERFORMANCES DES MODÈLES")
print("="*60)
comparison_df = pd.DataFrame(results).T
print(comparison_df.round(4))

# Meilleur modèle
best_model = comparison_df['RMSE'].idxmin()
print(f"\n🏆 MEILLEUR MODÈLE: {best_model}")
print(f"RMSE optimal: {comparison_df.loc[best_model, 'RMSE']:.4f}")

print("\nAnalyse terminée avec succès !")

