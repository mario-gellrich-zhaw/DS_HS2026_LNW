"""Generate fictional bicycle data for AP03; run: python generate_bicycle_data.py.

Writes bicycle_data.csv to the work package folder
DS_HS2026_LNW_I_Examples/AP03. Seed and
noise are fixed in advance, independently of model evaluation. Each row is
one bicycle; no repeated inspections and no real personal data.
"""
from pathlib import Path
import numpy as np
import pandas as pd

SEED = 42
N_BICYCLES = 1200
BASE = Path(__file__).resolve().parents[2] / 'DS_HS2026_LNW_I_Examples' / 'AP03'


def generate_data():
    rng = np.random.default_rng(SEED)
    age = np.round(rng.uniform(0.5, 10, N_BICYCLES), 1)
    mileage = np.round(age * rng.uniform(400, 1800, N_BICYCLES)).astype(int)
    equipment = rng.integers(1, 6, N_BICYCLES)
    ebike = rng.binomial(1, 0.4, N_BICYCLES)
    brake = np.round(rng.uniform(5, 95, N_BICYCLES), 1)
    chain = np.round(rng.uniform(5, 95, N_BICYCLES), 1)
    months = rng.integers(1, 31, N_BICYCLES)
    # Linear signal plus unobserved factors such as brand and negotiation.
    price = (1300 + 220 * equipment + 1700 * ebike - 95 * age
             - 0.03 * mileage + rng.normal(0, 180, N_BICYCLES))
    # Probabilistic workshop decisions rather than a deterministic cutoff.
    score = (brake - 65) / 9 + (chain - 65) / 11 + (months - 14) / 9
    service = rng.binomial(1, 1 / (1 + np.exp(-score)))
    return pd.DataFrame({
        'bicycle_id': [f'B{i:04d}' for i in range(1, N_BICYCLES + 1)],
        'age_years': age, 'mileage_km': mileage, 'equipment_level': equipment,
        'is_ebike': ebike, 'brake_wear_pct': brake, 'chain_wear_pct': chain,
        'months_since_service': months, 'price_chf': np.round(price, 2),
        'service_needed': service,
    })

def main():
    data = generate_data()
    data.to_csv(BASE / 'bicycle_data.csv', index=False, encoding='utf-8')
    print(f'Generated {len(data)} bicycles in {BASE}')
    print(f'Price range: CHF {data.price_chf.min():.2f} to {data.price_chf.max():.2f}')
    print(f'Share needing service: {data.service_needed.mean():.3f}')

if __name__ == '__main__':
    main()