from codecarbon import EmissionsTracker
import pandas as pd
import time

# Load training times recorded from R
times = pd.read_csv("D:/yqy/硕士-mpp/第四学期/ds&climate change/project/data/carbon_cost/training_times.csv")

results = []

for _, row in times.iterrows():
    tracker = EmissionsTracker(log_level="error")
    tracker.start()
    time.sleep(row["seconds"])       # Simulate training duration for each model
    emissions_kg = tracker.stop()    # CodeCarbon returns emissions in kg CO₂e
    results.append({
        "model": row["model"],
        "seconds": round(row["seconds"], 3),
        "emissions_g": round(emissions_kg * 1000, 6)    # Convert to gCO₂e
    })
    print(f"{row['model']}: {row['seconds']:.2f}s → {emissions_kg*1000:.4f} gCO₂e")

# Save results to CSV for import into R
pd.DataFrame(results).to_csv(
    "D:/yqy/硕士-mpp/第四学期/ds&climate change/project/data/carbon_cost/carbon_results.csv",
    index=False
)
print("\nDone! carbon_results.csv saved.")