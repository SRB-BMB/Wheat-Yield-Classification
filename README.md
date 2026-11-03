# Wheat Yield Category Classification 🌾

## Project Overview
This project uses Machine Learning to predict wheat yield categories (Low, Medium, and High) using pre-sowing information such as soil type, previous crop, sowing day, seed rate, fertilizer application, and rainfall.

The project focuses on wheat fields in eastern India.

## Objectives
- Predict wheat yield categories using pre-sowing information.
- Compare five Machine Learning classification algorithms.
- Evaluate model performance using accuracy, precision, recall, F1-score, confusion matrices, and ROC-AUC.
- Avoid data leakage by excluding information available only after sowing.

## Dataset
The project uses field-level wheat crop management, weather, and yield data from eastern India.

- 5,689 field observations
- 10 districts across Bihar and eastern Uttar Pradesh
- Six crop years (2016–17 to 2021–22)
- Target classes: Low, Medium, and High

## Machine Learning Algorithms
1. Logistic Regression
2. Decision Tree
3. Support Vector Machine (SVM)
4. K-Nearest Neighbors (KNN)
5. Gaussian Naive Bayes

## Results

| Model | Test Accuracy |
|---|---:|
| Support Vector Machine | 69.42% |
| K-Nearest Neighbors | 68.63% |
| Decision Tree | 64.32% |
| Logistic Regression | 62.04% |
| Gaussian Naive Bayes | 50.70% |

SVM achieved the highest test accuracy of 69.42%.

## Technologies Used
- Python
- Pandas
- NumPy
- Scikit-learn
- Excel
- Machine Learning

## Project Structure
```text
Wheat-Yield-Classification/
├── Code/
│   └── wheat_yield_classroom_demo.py
├── Dataset/
│   └── dataverse_files/
│       └── Wheat_Yld_Eastern_IND_Data.xlsx
├── Report/
│   └── Wheat_Yield_Project.pdf
├── README.md
└── requirements.txt
```

## How to Run

1. Clone this repository.
2. Install the required libraries:
   `pip install -r requirements.txt`
3. Ensure the Excel dataset is placed in the correct Dataset folder.
4. Run the Python file:

```bash
python Code/wheat_yield_classroom_demo.py
```

## Conclusion
The project demonstrates how Machine Learning can be used to classify wheat yield categories using information available before sowing. SVM achieved the best overall performance among the five evaluated classifiers.

The findings are observational and should not be interpreted as causal effects.

## Author
Suraj Bharad

B.Tech Computer Science and Engineering  
Lovely Professional University

## Acknowledgment
Thanks to Lovely Professional University, the project supervisor, and CIMMYT for the dataset and academic support.
