# monte-carlo-production-line
Monte Carlo simulation of a production line to evaluate reliability, throughput, buffer capacity, and preventive maintenance strategies.
The complete Excel workbook includes source data, distribution fitting, Monte Carlo modelling, and validation analysis.

[Download the full Excel workbook](https://docs.google.com/spreadsheets/d/1WJWuxzPR-P0S_l86INXvGOO6A7FvBj8N/edit?usp=drive_link&ouid=108802250627485773863&rtpof=true&sd=true)

To run the Python simulation, use the lightweight `Project.xlsx` file included in this repository.
The presentation reports results from 125 configurations with 1,000 Monte Carlo runs per configuration.

For buffers B1 = B2 = 30:
- Preventive maintenance every 100 operating hours achieved demand in 99.9% of simulated runs, compared with 69.2% under reactive maintenance.
- Mean finished output increased from approximately 27,422 to 29,088 units per 720-hour simulation.
- Output standard deviation decreased from 991 to 141 units.

The recommended configuration combines preventive maintenance every 100 operating hours with buffers of approximately 30 units.

[View the project presentation](Monte_Carlo_Presentation.pdf)
## How to Run

Requires Python 3.10 or later.

1. Download and extract the repository using **Code → Download ZIP**.
2. Keep `production_line_simulation.py` and `Project.xlsx` in the same folder.
3. Open a terminal in that folder and install the dependency:

```bash
python -m pip install openpyxl
```

4. Run the simulation:

```bash
python production_line_simulation.py --excel Project.xlsx --runs 100 --b1 30 --b2 30 --pm 100
```

This runs 100 simulations with both buffers set to 30 units and preventive maintenance every 100 operating hours.

Results are saved as CSV files in `simulation_results_v2`.

To run the complete experiment with 1,000 replications for each of the 125 configurations:

```bash
python production_line_simulation.py --excel Project.xlsx --experiment --runs 1000
```