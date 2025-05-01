import numpy as np
import pandas as pd
from openpyxl import Workbook

# 1. Tabular Representation of Performance Parameters (Formulas)
def get_performance_formulas():
    formulas_data = {
        "Metric": [
            "Character Recognition Accuracy",
            "Word Error Rate (WER)",
            "Character Error Rate (CER)",
            "Precision",
            "Recall (Sensitivity)",
            "F1-Score",
            "Processing Time per Page (Tp)",
            "Frames Per Second (FPS)",
            "Computational Efficiency Gain (%)",
            "Handwriting Complexity Score (HCS)",
            "Digitization Latency (L)"
        ],
        "Formula": [
            "Accuracy = (Correctly Recognized Characters / Total Characters) * 100",
            "WER = (S + D + I) / N  where S = Substitutions, D = Deletions, I = Insertions, N = Total Words",
            "CER = (S + D + I) / N  where S = Substitutions, D = Deletions, I = Insertions, N = Total Characters",
            "Precision = TP / (TP + FP)",
            "Recall = TP / (TP + FN)",
            "F1-Score = 2 * (Precision * Recall) / (Precision + Recall)",
            "Tp = Total Pages Processed / Total Processing Time",
            "FPS = Tp / 1",
            "CE = (Time Without PSO - Time With PSO) / Time Without PSO * 100",
            "HCS = (Sum of Handwriting Phases) / N",
            "L = Tp + Tpost"
        ],
        "Description": [
            "Measures the OCR’s correctness in extracting characters.",
            "Measures errors in digitized text: S = Substitutions, D = Deletions, I = Insertions, N = Total Words.",
            "Similar to WER but computed at the character level.",
            "Measures how many recognized words were correctly identified.",
            "Measures the ability to recognize actual words correctly.",
            "Balances precision and recall for OCR effectiveness.",
            "Measures the speed of handwritten note digitization.",
            "Determines real-time OCR processing capability.",
            "Measures performance improvement using PSO.",
            "Evaluates OCR performance on different handwriting styles.",
            "Measures the total time taken from input image processing to text output."
        ]
    }

    return pd.DataFrame(formulas_data)

# 2. Calculate Metrics
def calculate_metrics(true_text, predicted_text):
    true_words = true_text.split()
    predicted_words = predicted_text.split()

    # Word Error Rate (WER)
    substitutions = sum([w1 != w2 for w1, w2 in zip(true_words, predicted_words)])
    deletions = len(true_words) - len(predicted_words)
    insertions = len(predicted_words) - len(true_words)
    wer = (substitutions + deletions + insertions) / len(true_words) if len(true_words) > 0 else 0

    # Character Error Rate (CER)
    true_chars = list(''.join(true_words))
    predicted_chars = list(''.join(predicted_words))
    cer = sum([c1 != c2 for c1, c2 in zip(true_chars, predicted_chars)]) / len(true_chars) if len(true_chars) > 0 else 0

    # Precision, Recall, F1-Score
    true_positives = sum([w1 == w2 for w1, w2 in zip(true_words, predicted_words)])
    false_positives = len(predicted_words) - true_positives
    false_negatives = len(true_words) - true_positives
    precision = true_positives / (true_positives + false_positives) if true_positives + false_positives > 0 else 0
    recall = true_positives / (true_positives + false_negatives) if true_positives + false_negatives > 0 else 0
    f1_score = 2 * (precision * recall) / (precision + recall) if precision + recall > 0 else 0

    return wer, cer, precision, recall, f1_score

# 3. Save Metrics to Excel
def save_metrics_to_excel(wer, cer, precision, recall, f1_score):
    # Prepare the data to save into an Excel sheet for results
    metrics_data = {
        "Metric": [
            "Word Error Rate (WER)",
            "Character Error Rate (CER)",
            "Precision",
            "Recall",
            "F1-Score"
        ],
        "Value": [
            f"{wer*100:.2f}%",
            f"{cer*100:.2f}%",
            f"{precision:.2f}",
            f"{recall:.2f}",
            f"{f1_score:.2f}"
        ]
    }

    # Create DataFrame for performance formulas and results
    formulas_df = get_performance_formulas()
    results_df = pd.DataFrame(metrics_data)

    # Create a Workbook and write both DataFrames to separate sheets
    with pd.ExcelWriter("performance_metrics_results.xlsx", engine='openpyxl') as writer:
        formulas_df.to_excel(writer, sheet_name="Formulas", index=False)
        results_df.to_excel(writer, sheet_name="Results", index=False)

    print("\nMetrics and formulas have been saved to 'performance_metrics_results.xlsx'.")

# 4. Fitness Function for Optimization
def calculate_fitness(accuracy, wer, cer, hcs, latency):
    w1, w2, w3, w4, w5 = 1, 1, 1, 1, 1
    max_latency = 1000
    latency_normalized = latency / max_latency
    fitness = (w1 * accuracy) + (w2 * (1 - wer)) + (w3 * (1 - latency_normalized)) + (w4 * hcs) + (w5 * (1 - cer))
    return fitness

# CLI Menu
def performance_metrics_cli():
    print("\nPerformance Evaluation CLI")
    while True:
        print("\nMENU:")
        print("1. Load Text from Files & Calculate Metrics")
        print("2. Calculate Fitness Function")
        print("3. Exit")

        choice = input("\nEnter your choice: ")

        if choice == "1":
            true_path = input("Enter path to TRUE text file: ")
            pred_path = input("Enter path to PREDICTED text file: ")

            try:
                with open(true_path, "r", encoding="utf-8") as f:
                    true_text = f.read()
                with open(pred_path, "r", encoding="utf-8") as f:
                    predicted_text = f.read()

                # Calculate the metrics
                wer, cer, precision, recall, f1_score = calculate_metrics(true_text, predicted_text)

                # Save the metrics and formulas to an Excel file
                save_metrics_to_excel(wer, cer, precision, recall, f1_score)

            except Exception as e:
                print(f"Error reading files: {e}")

        elif choice == "2":
            accuracy = float(input("Enter accuracy (0 to 1): "))
            wer = float(input("Enter WER (0 to 1): "))
            cer = float(input("Enter CER (0 to 1): "))
            hcs = float(input("Enter Handwriting Complexity Score (0 to 1): "))
            latency = float(input("Enter latency in ms: "))

            # Calculate and display the fitness score
            fitness = calculate_fitness(accuracy, wer, cer, hcs, latency)
            print(f"\nFitness Score: {fitness:.2f}")

        elif choice == "3":
            print("Exiting. See you soon!")
            break

        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    performance_metrics_cli()
