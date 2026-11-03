"""
Data Preprocessing Module for Wheat Yield Category Classification.
Dataset: CIMMYT Eastern India Wheat Dataset (Bihar & Eastern UP, 2016-17 to 2021-22).
Course: INT234 - Predictive Analytics
Student: Parth Ghuge (Reg No: omitted from public copy, Roll No: omitted from public copy, Group: omitted from public copy, Section: omitted from public copy)
Supervisor: Course Supervisor
"""

import os
import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.impute import SimpleImputer

# Candidate pre-sowing numeric features available prior to or at sowing.
# NOTE: BasalZn is excluded because ~91.1% of records are missing and the variable guide
# does not confirm that unrecorded values imply zero application.
NUMERIC_FEATURES = [
    'PrevCropResidue',
    'SeedRate',
    'SowingDay',
    'BasalDAP',
    'BasalNPK',
    'BasalMOP',
    'Prcp_Monsoon',
    'Prcp_Oct'
]

# Candidate pre-sowing categorical features.
# NOTE: GradeNPK is excluded due to structural collinearity with BasalNPK (missing when BasalNPK == 0).
CATEGORICAL_FEATURES = [
    'Year',
    'District',
    'LandType',
    'SoilType',
    'PreviousCrop',
    'CropEstablishment',
    'Variety',
    'VarietyClass'
]

TARGET_COLUMN = 'GrainYield'


def assign_yield_category(yield_val, c1, c2):
    """
    Categorize continuous GrainYield into data-derived tertile categories:
    - Low: <= c1
    - Medium: > c1 and <= c2
    - High: > c2
    """
    if yield_val <= c1:
        return 'Low'
    elif yield_val <= c2:
        return 'Medium'
    else:
        return 'High'


def load_raw_data(filepath):
    """Load raw Excel data and perform initial structural validation."""
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Dataset file not found at: {filepath}")
    df = pd.read_excel(filepath)
    return df


def clean_feature_representations(df):
    """
    Standardize text representations across categorical columns:
    - Clean case discrepancies in SoilType and PreviousCrop
    """
    df_clean = df.copy()

    # Standardize casing to avoid duplicate categories
    df_clean['SoilType'] = df_clean['SoilType'].astype(str).str.capitalize()
    df_clean['PreviousCrop'] = df_clean['PreviousCrop'].astype(str).str.capitalize()

    return df_clean


def build_preprocessor(numeric_features=NUMERIC_FEATURES, categorical_features=CATEGORICAL_FEATURES):
    """
    Construct a scikit-learn ColumnTransformer for pre-sowing features:
    - Numeric: Median imputation (fit on training data) + StandardScaler
    - Categorical: Most frequent imputation (fit on training data) + OneHotEncoder
    """
    num_pipeline = Pipeline([
        ('imputer', SimpleImputer(strategy='median')),
        ('scaler', StandardScaler())
    ])

    cat_pipeline = Pipeline([
        ('imputer', SimpleImputer(strategy='most_frequent')),
        ('ohe', OneHotEncoder(handle_unknown='ignore', sparse_output=False))
    ])

    preprocessor = ColumnTransformer(
        transformers=[
            ('num', num_pipeline, numeric_features),
            ('cat', cat_pipeline, categorical_features)
        ],
        remainder='drop'
    )
    return preprocessor


def prepare_dataset(filepath, test_size=0.2, random_state=42):
    """
    Execute end-to-end data preparation pipeline:
    1. Load raw data (5,689 records).
    2. Standardize casing in categorical features.
    3. Stratified 80/20 train/test split with fixed random seed.
    4. Compute empirical tertile cutoffs strictly on training set.
    5. Categorize target into 'Low', 'Medium', 'High'.
    6. Fit preprocessor strictly on training data and transform train and test sets.
    """
    raw_df = load_raw_data(filepath)
    clean_df = clean_feature_representations(raw_df)

    # Initial tertiles on full dataset used solely to guide stratified sampling
    q1_strat, q2_strat = clean_df[TARGET_COLUMN].quantile([1/3, 2/3])
    strat_labels = clean_df[TARGET_COLUMN].apply(lambda y: assign_yield_category(y, q1_strat, q2_strat))

    train_idx, test_idx = train_test_split(
        clean_df.index,
        test_size=test_size,
        random_state=random_state,
        stratify=strat_labels
    )

    train_df = clean_df.loc[train_idx].copy().reset_index(drop=True)
    test_df = clean_df.loc[test_idx].copy().reset_index(drop=True)

    # Calculate data-derived cutoffs strictly from the training partition
    c1 = float(train_df[TARGET_COLUMN].quantile(1/3))
    c2 = float(train_df[TARGET_COLUMN].quantile(2/3))

    cutoffs = {'c1_lower': c1, 'c2_upper': c2}

    # Generate discrete class labels
    y_train = train_df[TARGET_COLUMN].apply(lambda y: assign_yield_category(y, c1, c2))
    y_test = test_df[TARGET_COLUMN].apply(lambda y: assign_yield_category(y, c1, c2))

    # Feature subsets
    X_train_raw = train_df[NUMERIC_FEATURES + CATEGORICAL_FEATURES].copy()
    X_test_raw = test_df[NUMERIC_FEATURES + CATEGORICAL_FEATURES].copy()

    # Preprocessing transformation fit strictly on training partition
    preprocessor = build_preprocessor(NUMERIC_FEATURES, CATEGORICAL_FEATURES)
    X_train = preprocessor.fit_transform(X_train_raw)
    X_test = preprocessor.transform(X_test_raw)

    # Extract transformed feature names
    cat_encoder = preprocessor.named_transformers_['cat'].named_steps['ohe']
    cat_names = list(cat_encoder.get_feature_names_out(CATEGORICAL_FEATURES))
    feature_names = NUMERIC_FEATURES + cat_names

    return {
        'X_train': X_train,
        'X_test': X_test,
        'y_train': y_train,
        'y_test': y_test,
        'X_train_raw': X_train_raw,
        'X_test_raw': X_test_raw,
        'train_df': train_df,
        'test_df': test_df,
        'cutoffs': cutoffs,
        'feature_names': feature_names,
        'preprocessor': preprocessor
    }


if __name__ == '__main__':
    data_path = os.path.join(os.path.dirname(__file__), '..', 'Dataset', 'dataverse_files', 'Wheat_Yld_Eastern_IND_Data.xlsx')
    data_bundle = prepare_dataset(data_path)
    print("Data preparation complete.")
    print(f"X_train shape: {data_bundle['X_train'].shape}")
    print(f"X_test shape: {data_bundle['X_test'].shape}")
    print(f"Training-derived cutoffs: Lower = {data_bundle['cutoffs']['c1_lower']:.4f} t/ha, Upper = {data_bundle['cutoffs']['c2_upper']:.4f} t/ha")
    print("\nTrain class distribution:\n", data_bundle['y_train'].value_counts())
    print("\nTest class distribution:\n", data_bundle['y_test'].value_counts())
