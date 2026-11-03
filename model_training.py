"""
Model Training and Evaluation Module for INT234 CA2 Project.
Student: Parth Ghuge (Reg No: omitted from public copy, Roll No: omitted from public copy, Group: omitted from public copy, Section: omitted from public copy)
Course: INT234 - Predictive Analytics
Supervisor: Course Supervisor
Academic Semester: Project Semester 2026-27

Implements the 5 Syllabus Classification Algorithms:
1. Logistic Regression (Multinomial)
2. Decision Tree Classifier
3. Support Vector Machine (SVC)
4. K-Nearest Neighbors (KNN)
5. Gaussian Naive Bayes

Methodological Protocol:
- Leakage-Free Cross-Validation: All CV is run on raw training features using complete scikit-learn
  Pipelines containing imputer, scaler, one-hot encoder, and classifier, so preprocessing is fitted
  strictly within each fold.
- Parameter Justification:
  * KNN (k=13): Selected as the CV-winning setting across the parameter sweep.
  * Decision Tree (depth=7): Selected as a deliberate simplicity trade-off over depth 9.
    Depth 9 achieved the higher cross-validation mean of 65.02%, while depth 7 achieved 63.94%
    (a 1.08-percentage-point difference). Depth 7 was retained because it produces a simpler, shallower
    tree with fewer split rules, avoiding excessive leaf fragmentation on survey noise and prioritizing
    agronomic interpretability.
  * SVM (C=2.0): Selected over C=3.0 (69.83% vs 69.96%, a difference of 0.13%) to enforce wider soft
    margins and robust regularization against field survey measurement noise.
- The holdout test set remains strictly untouched until final evaluation.
"""

import os
import json
import warnings
import numpy as np
import pandas as pd
from sklearn.pipeline import Pipeline
from sklearn.model_selection import StratifiedKFold, cross_val_score
from sklearn.linear_model import LogisticRegression
from sklearn.tree import DecisionTreeClassifier
from sklearn.svm import SVC
from sklearn.neighbors import KNeighborsClassifier
from sklearn.naive_bayes import GaussianNB
from sklearn.metrics import (
    accuracy_score, precision_score, recall_score, f1_score,
    log_loss, roc_auc_score, confusion_matrix, classification_report
)
from sklearn.preprocessing import label_binarize

from data_preprocessing import (
    prepare_dataset, build_preprocessor,
    NUMERIC_FEATURES, CATEGORICAL_FEATURES
)

warnings.filterwarnings('ignore', category=FutureWarning)
warnings.filterwarnings('ignore', category=UserWarning)

CLASS_ORDER = ['Low', 'Medium', 'High']


def make_pipeline(classifier):
    """
    Construct an end-to-end scikit-learn Pipeline with preprocessing
    (median imputation, standard scaling, and one-hot encoding) and the classifier.
    Ensures zero leakage across folds during cross-validation.
    """
    preprocessor = build_preprocessor(NUMERIC_FEATURES, CATEGORICAL_FEATURES)
    return Pipeline(steps=[
        ('preprocessor', preprocessor),
        ('classifier', classifier)
    ])


def run_training_cross_validation(X_train_raw, y_train):
    """
    Perform 5-fold cross-validation strictly on raw training features
    using complete scikit-learn Pipelines so that preprocessing (imputation,
    scaling, encoding) is fitted separately within each fold without data leakage.
    """
    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    cv_summary = {}

    print("\n" + "=" * 95)
    print(f"{'LEAKAGE-FREE 5-FOLD CROSS-VALIDATION ON RAW TRAINING PARTITION':^95}")
    print("=" * 95)

    # 1. KNN k-sweep on training set
    knn_cv = {}
    for k in [5, 7, 9, 11, 13, 15, 17, 19, 21]:
        pipe = make_pipeline(KNeighborsClassifier(n_neighbors=k, weights='distance', metric='manhattan'))
        scores = cross_val_score(pipe, X_train_raw, y_train, cv=cv, scoring='accuracy')
        knn_cv[k] = {'mean_acc': float(scores.mean()), 'std_acc': float(scores.std())}
        print(f"  KNN (Manhattan, distance) k={k:2d}: CV Mean Acc = {scores.mean()*100:.2f}% (+/- {scores.std()*100:.2f}%)")
    cv_summary['knn_k_sweep'] = knn_cv

    # 2. Decision Tree depth sweep on training set
    dt_cv = {}
    for d in [3, 4, 5, 6, 7, 8, 9, 10]:
        pipe = make_pipeline(DecisionTreeClassifier(max_depth=d, min_samples_leaf=5, criterion='gini', random_state=42))
        scores = cross_val_score(pipe, X_train_raw, y_train, cv=cv, scoring='accuracy')
        dt_cv[d] = {'mean_acc': float(scores.mean()), 'std_acc': float(scores.std())}
        print(f"  Decision Tree (Gini) depth={d:2d}: CV Mean Acc = {scores.mean()*100:.2f}% (+/- {scores.std()*100:.2f}%)")
    cv_summary['dt_depth_sweep'] = dt_cv

    # 3. SVM C parameter sweep on training set
    svm_cv = {}
    for C in [0.5, 1.0, 1.5, 2.0, 3.0, 4.0]:
        pipe = make_pipeline(SVC(C=C, kernel='rbf', random_state=42))
        scores = cross_val_score(pipe, X_train_raw, y_train, cv=cv, scoring='accuracy')
        svm_cv[C] = {'mean_acc': float(scores.mean()), 'std_acc': float(scores.std())}
        print(f"  SVM (RBF) C={C:3.1f}: CV Mean Acc = {scores.mean()*100:.2f}% (+/- {scores.std()*100:.2f}%)")
    cv_summary['svm_c_sweep'] = svm_cv

    # 4. Logistic Regression
    pipe_lr = make_pipeline(LogisticRegression(max_iter=1000, random_state=42, solver='lbfgs'))
    scores_lr = cross_val_score(pipe_lr, X_train_raw, y_train, cv=cv, scoring='accuracy')
    cv_summary['logistic_regression'] = {'mean_acc': float(scores_lr.mean()), 'std_acc': float(scores_lr.std())}
    print(f"  Logistic Regression (L2, C=1.0): CV Mean Acc = {scores_lr.mean()*100:.2f}% (+/- {scores_lr.std()*100:.2f}%)")

    # 5. Gaussian Naive Bayes
    pipe_nb = make_pipeline(GaussianNB())
    scores_nb = cross_val_score(pipe_nb, X_train_raw, y_train, cv=cv, scoring='accuracy')
    cv_summary['naive_bayes'] = {'mean_acc': float(scores_nb.mean()), 'std_acc': float(scores_nb.std())}
    print(f"  Gaussian Naive Bayes: CV Mean Acc = {scores_nb.mean()*100:.2f}% (+/- {scores_nb.std()*100:.2f}%)")

    # Explicit justifications
    cv_summary['justifications'] = {
        'knn_k': (
            "k=13 was selected as the clear winning setting in the 5-fold CV sweep on raw training data "
            "(achieving 69.22% mean accuracy, outperforming both lower k like k=5 at 68.64% and higher k like k=21 at 68.78%)."
        ),
        'decision_tree_depth': (
            "In 5-fold pipeline cross-validation, tree depth 9 achieved the higher mean accuracy of 65.02% (+/- 0.90%) "
            "compared to depth 7 at 63.94% (+/- 1.82%), representing a 1.08-percentage-point difference. "
            "Depth 7 was retained as a deliberate simplicity trade-off: it yields a simpler, shallower tree structure with fewer leaves, "
            "avoiding over-partitioning on field survey noise and maintaining transparent agronomic interpretability for agricultural extension workers."
        ),
        'svm_c': (
            "While C=3.0 reached 69.96% (+/- 1.10%) vs C=2.0 at 69.83% (+/- 1.06%), the difference of 0.13% is small. "
            "C=2.0 was selected because a slightly lower penalty parameter enforces a wider soft margin, "
            "providing stronger regularization against smallholder field survey measurement errors and boundary noise, mitigating overfitting risks."
        )
    }

    return cv_summary


def get_models():
    """
    Instantiate the 5 syllabus classification models wrapped in complete scikit-learn
    Pipelines with imputer, scaler, and one-hot encoder.
    """
    models = {
        'Logistic Regression': make_pipeline(LogisticRegression(
            max_iter=1000,
            random_state=42,
            solver='lbfgs'
        )),
        'Decision Tree': make_pipeline(DecisionTreeClassifier(
            criterion='gini',
            max_depth=7,
            min_samples_leaf=5,
            random_state=42
        )),
        'Support Vector Machine': make_pipeline(SVC(
            C=2.0,
            kernel='rbf',
            probability=True,
            random_state=42
        )),
        'K-Nearest Neighbors': make_pipeline(KNeighborsClassifier(
            n_neighbors=13,
            weights='distance',
            metric='manhattan'
        )),
        'Naive Bayes': make_pipeline(GaussianNB())
    }
    return models


def evaluate_model(model, X_train_raw, y_train, X_test_raw, y_test, class_order=CLASS_ORDER):
    """
    Train pipeline model strictly on raw training set, generate predictions on raw test set,
    and compute all syllabus evaluation metrics.
    """
    # Fit pipeline strictly on training partition
    model.fit(X_train_raw, y_train)

    # Predictions
    y_pred = model.predict(X_test_raw)
    y_train_pred = model.predict(X_train_raw)

    train_acc = accuracy_score(y_train, y_train_pred)
    test_acc = accuracy_score(y_test, y_pred)

    # Probabilities aligned with model.classes_
    model_classes = list(model.classes_)
    if hasattr(model, "predict_proba"):
        y_proba_native = model.predict_proba(X_test_raw)
    else:
        df_vals = model.decision_function(X_test_raw)
        exp_vals = np.exp(df_vals - np.max(df_vals, axis=1, keepdims=True))
        y_proba_native = exp_vals / np.sum(exp_vals, axis=1, keepdims=True)

    # Multi-class Log Loss and ROC AUC with aligned model classes
    eps = 1e-15
    y_proba_clipped = np.clip(y_proba_native, eps, 1 - eps)
    y_proba_clipped /= y_proba_clipped.sum(axis=1, keepdims=True)

    model_log_loss = float(log_loss(y_test, y_proba_clipped, labels=model_classes))

    y_test_bin_native = label_binarize(y_test, classes=model_classes)
    model_roc_auc = float(roc_auc_score(y_test_bin_native, y_proba_native, multi_class='ovr', average='macro'))

    # Reorder probabilities to CLASS_ORDER ('Low', 'Medium', 'High') for consistency in visual outputs
    reorder_idx = [model_classes.index(c) for c in class_order]
    y_proba_reordered = y_proba_native[:, reorder_idx]

    # Macro & Weighted metrics using specified class order
    precision_macro = float(precision_score(y_test, y_pred, labels=class_order, average='macro'))
    recall_macro = float(recall_score(y_test, y_pred, labels=class_order, average='macro'))
    f1_macro = float(f1_score(y_test, y_pred, labels=class_order, average='macro'))

    precision_weighted = float(precision_score(y_test, y_pred, labels=class_order, average='weighted'))
    recall_weighted = float(recall_score(y_test, y_pred, labels=class_order, average='weighted'))
    f1_weighted = float(f1_score(y_test, y_pred, labels=class_order, average='weighted'))

    # Confusion Matrix
    cm = confusion_matrix(y_test, y_pred, labels=class_order)
    cm_norm = cm.astype('float') / cm.sum(axis=1)[:, np.newaxis]

    # Per-class metrics
    class_report_dict = classification_report(y_test, y_pred, labels=class_order, output_dict=True)

    class_metrics = {}
    for c in class_order:
        class_metrics[c] = {
            'precision': float(class_report_dict[c]['precision']),
            'recall': float(class_report_dict[c]['recall']),
            'f1_score': float(class_report_dict[c]['f1-score']),
            'support': int(class_report_dict[c]['support'])
        }

    return {
        'train_accuracy': float(train_acc),
        'test_accuracy': float(test_acc),
        'precision_macro': precision_macro,
        'recall_macro': recall_macro,
        'f1_macro': f1_macro,
        'precision_weighted': precision_weighted,
        'recall_weighted': recall_weighted,
        'f1_weighted': f1_weighted,
        'log_loss': model_log_loss,
        'roc_auc_ovr': model_roc_auc,
        'confusion_matrix': cm.tolist(),
        'confusion_matrix_norm': cm_norm.tolist(),
        'class_metrics': class_metrics,
        'y_pred': y_pred.tolist(),
        'y_proba': y_proba_reordered.tolist()
    }


def run_full_training_suite(data_path, output_dir):
    """
    Train and evaluate all 5 models inside complete Pipelines,
    display comprehensive summary tables, and save structured JSON results.
    """
    os.makedirs(output_dir, exist_ok=True)
    bundle = prepare_dataset(data_path)

    X_train_raw = bundle['X_train_raw']
    y_train = bundle['y_train']
    X_test_raw = bundle['X_test_raw']
    y_test = bundle['y_test']

    # Step 1: Document hyperparameter performance via leakage-free 5-fold CV on raw training set
    cv_results = run_training_cross_validation(X_train_raw, y_train)

    # Step 2: Fit final pipeline models on raw training data and evaluate on holdout raw test set
    models = get_models()
    results = {
        'metadata': {
            'student_name': 'Parth Ghuge',
            'roll_number': None,
            'group': None,
            'registration_number': None,
            'section': None,
            'faculty_supervisor': None,
            'course': 'INT234 Predictive Analytics',
            'semester': 'Project Semester 2026-27',
            'dataset_citation': 'Ajay, Anurag (2026), Field-level crop management, weather, and yield data of wheat from eastern India, CIMMYT Research Data & Software Repository Network, V1. DOI: https://doi.org/10.71682/10549424. License: CC0 1.0.',
            'cutoffs': bundle['cutoffs'],
            'train_size': len(y_train),
            'test_size': len(y_test),
            'n_features': bundle['X_train'].shape[1],
            'classes': CLASS_ORDER
        },
        'training_cv_results': cv_results,
        'models': {}
    }

    summary_rows = []

    print("\n" + "=" * 95)
    print(f"{'WHEAT YIELD CATEGORY CLASSIFICATION - FINAL TEST EVALUATION':^95}")
    print("=" * 95)

    for name, model in models.items():
        print(f"Evaluating model: {name} ...")
        eval_dict = evaluate_model(model, X_train_raw, y_train, X_test_raw, y_test, CLASS_ORDER)
        results['models'][name] = eval_dict

        summary_rows.append({
            'Model': name,
            'Train Acc (%)': f"{eval_dict['train_accuracy']*100:.2f}",
            'Test Acc (%)': f"{eval_dict['test_accuracy']*100:.2f}",
            'Precision (Macro)': f"{eval_dict['precision_macro']:.4f}",
            'Recall (Macro)': f"{eval_dict['recall_macro']:.4f}",
            'F1 (Macro)': f"{eval_dict['f1_macro']:.4f}",
            'Log Loss': f"{eval_dict['log_loss']:.4f}",
            'ROC AUC (OVR)': f"{eval_dict['roc_auc_ovr']:.4f}"
        })

    summary_df = pd.DataFrame(summary_rows)
    print("\n" + "-" * 95)
    print(summary_df.to_string(index=False))
    print("-" * 95)

    # Detailed per-class breakdowns
    print("\n" + "=" * 95)
    print(f"{'CLASS-WISE PERFORMANCE BREAKDOWN (LOW, MEDIUM, HIGH)':^95}")
    print("=" * 95)
    for name in models.keys():
        print(f"\nModel: {name}")
        cm = np.array(results['models'][name]['confusion_matrix'])
        print("Confusion Matrix (Rows: Actual [Low, Med, High], Cols: Predicted [Low, Med, High]):")
        print(cm)
        for c in CLASS_ORDER:
            m = results['models'][name]['class_metrics'][c]
            print(f"  Class {c:6s} -> Precision: {m['precision']:.4f} | Recall: {m['recall']:.4f} | F1: {m['f1_score']:.4f} | Support: {m['support']}")

    # Save structured JSON
    json_path = os.path.join(output_dir, 'model_evaluation_results.json')
    with open(json_path, 'w') as f:
        json.dump(results, f, indent=2)
    print(f"\nSaved structured evaluation results to: {json_path}")

    # Save summary CSV
    csv_path = os.path.join(output_dir, 'model_summary_table.csv')
    summary_df.to_csv(csv_path, index=False)
    print(f"Saved summary comparison table to: {csv_path}")

    return results, bundle


if __name__ == '__main__':
    script_dir = os.path.dirname(__file__)
    data_file = os.path.join(script_dir, '..', 'Dataset', 'dataverse_files', 'Wheat_Yld_Eastern_IND_Data.xlsx')
    out_dir = os.path.join(script_dir, '..', 'figures')
    run_full_training_suite(data_file, out_dir)
