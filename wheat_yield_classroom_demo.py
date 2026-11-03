# Wheat Yield Category Classification
# Simple version for explaining the main project steps.

import os
import pandas as pd

from sklearn.model_selection import train_test_split
from sklearn.impute import SimpleImputer
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import accuracy_score, confusion_matrix, classification_report
from sklearn.linear_model import LogisticRegression
from sklearn.tree import DecisionTreeClassifier
from sklearn.svm import SVC
from sklearn.neighbors import KNeighborsClassifier
from sklearn.naive_bayes import GaussianNB

# 1. Load the dataset.
folder = os.path.dirname(__file__)
file_path = os.path.join(
    folder, "..", "Dataset", "dataverse_files",
    "Wheat_Yld_Eastern_IND_Data.xlsx"
)
data = pd.read_excel(file_path)

# 2. Select information known before or at sowing.
numeric_columns = [
    "PrevCropResidue", "SeedRate", "SowingDay", "BasalDAP",
    "BasalNPK", "BasalMOP", "Prcp_Monsoon", "Prcp_Oct"
]

category_columns = [
    "Year", "District", "LandType", "SoilType", "PreviousCrop",
    "CropEstablishment", "Variety", "VarietyClass"
]
feature_columns = numeric_columns + category_columns

# Keep text capitalization consistent.
for column in ["SoilType", "PreviousCrop"]:
    data[column] = data[column].astype(str).str.capitalize()

# 3. Split the data into training and test sets.
# These labels are only used to keep the split balanced.
first_cutoff, second_cutoff = data["GrainYield"].quantile([1/3, 2/3])
stratify_labels = pd.cut(
    data["GrainYield"],
    bins=[-float("inf"), first_cutoff, second_cutoff, float("inf")],
    labels=["Low", "Medium", "High"]
)

train_rows, test_rows = train_test_split(
    data.index,
    test_size=0.2,
    random_state=42,
    stratify=stratify_labels
)
train_data = data.loc[train_rows].copy()
test_data = data.loc[test_rows].copy()

# 4. Create yield categories using cutoffs calculated from training data.
low_cutoff = train_data["GrainYield"].quantile(1/3)
high_cutoff = train_data["GrainYield"].quantile(2/3)

def label_yield(value):
    if value <= low_cutoff:
        return "Low"
    elif value <= high_cutoff:
        return "Medium"
    else:
        return "High"

y_train = train_data["GrainYield"].apply(label_yield)
y_test = test_data["GrainYield"].apply(label_yield)

X_train = train_data[feature_columns].copy()
X_test = test_data[feature_columns].copy()

print("Training records:", len(X_train))
print("Test records:", len(X_test))
print("Yield cutoffs:", round(low_cutoff, 2), "and", round(high_cutoff, 2), "tonnes/ha")

# 5. Fill missing values using only the training data.
number_filler = SimpleImputer(strategy="median")
X_train[numeric_columns] = number_filler.fit_transform(X_train[numeric_columns])
X_test[numeric_columns] = number_filler.transform(X_test[numeric_columns])

for column in category_columns:
    if column in ["SoilType", "PreviousCrop"]:
        # The original data uses the text "Nan" for missing values in these fields.
        X_train[column] = X_train[column].fillna("Nan")
        X_test[column] = X_test[column].fillna("Nan")
    else:
        most_common = X_train[column].mode()[0]
        X_train[column] = X_train[column].fillna(most_common)
        X_test[column] = X_test[column].fillna(most_common)

# 6. Turn text categories into 0/1 columns and scale numeric values.
X_train = pd.get_dummies(X_train, columns=category_columns)
X_test = pd.get_dummies(X_test, columns=category_columns)
X_test = X_test.reindex(columns=X_train.columns, fill_value=0)

scaler = StandardScaler()
X_train[numeric_columns] = scaler.fit_transform(X_train[numeric_columns])
X_test[numeric_columns] = scaler.transform(X_test[numeric_columns])

# 7. Set up the five classifiers.
# These settings were selected in the full 5-fold CV analysis.
classifiers = {
    "Logistic Regression": LogisticRegression(max_iter=1000, random_state=42),
    "Decision Tree": DecisionTreeClassifier(
        max_depth=7, min_samples_leaf=5, random_state=42
    ),
    "Support Vector Machine": SVC(C=2.0, kernel="rbf"),
    "K-Nearest Neighbors": KNeighborsClassifier(
        n_neighbors=13, weights="distance", metric="manhattan"
    ),
    "Naive Bayes": GaussianNB()
}

# 8. Train each classifier, make predictions, and compare accuracy.
results = []

for name, classifier in classifiers.items():
    classifier.fit(X_train, y_train)
    predictions = classifier.predict(X_test)
    accuracy = accuracy_score(y_test, predictions)
    results.append([name, round(accuracy * 100, 2)])

    print("\n----------------------------------------")
    print(name)
    print("Accuracy:", round(accuracy * 100, 2), "%")
    print("Confusion matrix (actual rows, predicted columns):")
    print(confusion_matrix(y_test, predictions, labels=["Low", "Medium", "High"]))
    print("Classification report:")
    print(classification_report(
        y_test, predictions,
        labels=["Low", "Medium", "High"],
        zero_division=0
    ))

# 9. Show a final comparison table.
results_table = pd.DataFrame(results, columns=["Classifier", "Accuracy (%)"])
results_table = results_table.sort_values(
    by="Accuracy (%)", ascending=False
)

print("\nFINAL ACCURACY COMPARISON")
print(results_table.to_string(index=False))
