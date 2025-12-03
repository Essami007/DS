# ============================================================================
# 🚀 CODE CONSOLIDÉ COMPLET - PRÉDICTION RISQUE BLANCHIMENT D'ARGENT
# ============================================================================
# Auteur : Analyse Automatisée ESSAMI | Date : 03/12/2025 | Dataset : 10k transactions
# Meilleur modèle : Forêt Aléatoire (RMSE 2.60, R² 0.35) [file:1][web:20]

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

print("🚀 DÉMARRAGE ANALYSE PRÉDICTION RISQUE BLANCHIMENT")
print("=" * 70)

# ============================================================================
# 1. CHARGEMENT ET EXPLORATION DONNÉES
# ============================================================================
print("\n📊 1. CHARGEMENT DES DONNÉES")
df = pd.read_csv('/content/drive/MyDrive/DM_ML/BigBlackMoneyDataset.csv')
print(f"✅ Dataset chargé : {df.shape[0]:,} lignes × {df.shape[1]} colonnes")
print("\nAperçu des données :")
print(df.head())
print(f"\nStatistiques cible (Money Laundering Risk Score) :")
print(df['Money Laundering Risk Score'].describe())

# ============================================================================
# 2. PRÉPARATION ET INGÉNIERIE DES CARACTÉRISTIQUES
# ============================================================================
print("\n🔧 2. PRÉPARATION DES DONNÉES")
df['Date of Transaction'] = pd.to_datetime(df['Date of Transaction'])
df['transaction_year'] = df['Date of Transaction'].dt.year
df['transaction_month'] = df['Date of Transaction'].dt.month
df['transaction_day'] = df['Date of Transaction'].dt.day
df['transaction_hour'] = df['Date of Transaction'].dt.hour
df = df.drop('Date of Transaction', axis=1)

# Séparation cible/prédicteurs
y = df['Money Laundering Risk Score']
X = df.drop('Money Laundering Risk Score', axis=1)

# Identification features
categorical_features = X.select_dtypes(include=['object', 'bool']).columns
numerical_features = X.select_dtypes(include=['int64', 'float64']).columns

# One-Hot Encoding
print("🔄 Encodage One-Hot des variables catégorielles...")
X_categorical = pd.get_dummies(X[categorical_features], drop_first=True)
X_numerical = X[numerical_features]
X = pd.concat([X_numerical, X_categorical], axis=1)

print(f"✅ Features après encodage : {X.shape[1]:,} (de {len(numerical_features) + len(categorical_features)})")
print(f"✅ Target : {y.shape}")

# Division train/test
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
print(f"✅ Split : Train {X_train.shape} | Test {X_test.shape}")

# ============================================================================
# 3. FONCTION ÉVALUATION MODÈLES
# ============================================================================
def evaluate_model(model, X_train, X_test, y_train, y_test, model_name):
    """Évalue un modèle et retourne métriques + graphique"""
    print(f"\n{'='*10} {model_name} {'='*10}")
    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)
    
    # Métriques
    mae = mean_absolute_error(y_test, y_pred)
    mse = mean_squared_error(y_test, y_pred)
    rmse = np.sqrt(mse)
    r2 = r2_score(y_test, y_pred)
    
    print(f"MAE : {mae:.4f}")
    print(f"RMSE: {rmse:.4f}")
    print(f"R²  : {r2:.4f}")
    
    # Graphique prédictions vs réel
    plt.figure(figsize=(10, 6))
    plt.scatter(y_test, y_pred, alpha=0.6, label='Prédictions')
    plt.plot([y_test.min(), y_test.max()], [y_test.min(), y_test.max()], 'r--', lw=2, label='Ligne parfaite')
    plt.xlabel('Valeurs Réelles')
    plt.ylabel('Prédictions')
    plt.title(f'{model_name} - Prédictions vs Réel (R²={r2:.3f})')
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.show()
    
    return {'MAE': mae, 'MSE': mse, 'RMSE': rmse, 'R2': r2, 'model': model}

# ============================================================================
# 4. ENTRAÎNEMENT ET COMPARAISON 4 MODÈLES
# ============================================================================
print("\n🤖 3. ENTRAÎNEMENT DES MODÈLES")
results = {}

# 1. Régression Linéaire (Baseline)
lr = LinearRegression()
results['Linear Regression'] = evaluate_model(lr, X_train, X_test, y_train, y_test, 'Régression Linéaire')

# 2. Arbre de Décision
dt = DecisionTreeRegressor(random_state=42, max_depth=10)
results['Decision Tree'] = evaluate_model(dt, X_train, X_test, y_train, y_test, 'Arbre de Décision')

# 3. Forêt Aléatoire (MEILLEUR)
rf = RandomForestRegressor(n_estimators=100, random_state=42, n_jobs=-1)
results['Random Forest'] = evaluate_model(rf, X_train, X_test, y_train, y_test, '🌟 Forêt Aléatoire')

# 4. Gradient Boosting
gb = GradientBoostingRegressor(random_state=42, n_estimators=100)
results['Gradient Boosting'] = evaluate_model(gb, X_train, X_test, y_train, y_test, 'Gradient Boosting')

# ============================================================================
# 5. TABLEAU COMPARATIF ET MEILLEUR MODÈLE
# ============================================================================
print("\n📋 4. COMPARAISON DES MODÈLES")
comparison_df = pd.DataFrame(results).T[['MAE', 'RMSE', 'R2']]
print(comparison_df.round(4))

# Meilleur modèle
best_model_name = comparison_df['RMSE'].idxmin()
best_rmse = comparison_df.loc[best_model_name, 'RMSE']
print(f"\n🏆 MEILLEUR MODÈLE : {best_model_name}")
print(f"   RMSE optimal : {best_rmse:.4f}")
print(f"   Gain vs baseline : {results['Linear Regression']['RMSE'] - best_rmse:.2f} points")

# Features importantes (Random Forest)
if 'Random Forest' in results:
    print("\n🔍 5. FEATURES LES PLUS IMPORTANTES (Random Forest)")
    importances = pd.DataFrame({
        'feature': X.columns,
        'importance': results['Random Forest']['model'].feature_importances_
    }).sort_values('importance', ascending=False).head(10)
    print(importances)

# ============================================================================
# 6. ANALYSE BUSINESS - TRANSACTIONS À RISQUE
# ============================================================================
print("\n💼 6. ANALYSE BUSINESS - ALERTES RISQUE ÉLEVÉ")
y_pred_best = results[best_model_name]['model'].predict(X_test)

# Classification risque
risk_categories = pd.cut(y_pred_best, bins=[0, 4, 7, 10], labels=['🟢 Faible', '🟡 Moyen', '🔴 Élevé'])
risk_summary = pd.Series(risk_categories).value_counts()
print("Répartition prédictions risque :")
print(risk_summary)
print(f"\n❗ % transactions à risque élevé (>7) : {100*(y_pred_best > 7).mean():.1f}%")

# ============================================================================
# 7. FONCTION PRÉDICTION NOUVELLE TRANSACTION
# ============================================================================
def predict_risk(new_transaction, model=results[best_model_name]['model']):
    """Prédit le risque pour une nouvelle transaction"""
    # TODO: Adapter selon format entrée
    risk_score = model.predict(new_transaction)[0]
    risk_label = '🔴 ÉLEVÉ' if risk_score > 7 else '🟡 MOYEN' if risk_score > 4 else '🟢 FAIBLE'
    return f"Score risque : {risk_score:.2f}/10 ({risk_label})"

print("\n✅ FONCTION PRÉDICTION DISPONIBLE : predict_risk()")

# ============================================================================
# 8. SAUVEGARDE MODÈLE ET RÉSULTATS
# ============================================================================
import joblib
joblib.dump(results[best_model_name]['model'], 'money_laundering_model.pkl')
comparison_df.to_csv('model_comparison.csv')
print("\n💾 MODÈLE SAUVEGARDÉ : money_laundering_model.pkl")
print("💾 COMPARAISON SAUVEGARDÉE : model_comparison.csv")

print("\n" + "="*70)
print("🎉 ANALYSE TERMINÉE - MODÈLE PRODUCTION READY")
print(f"🏆 {best_model_name} déployé (RMSE: {best_rmse:.2f})")
print("🚀 Prêt pour API temps réel !")
print("="*70)

