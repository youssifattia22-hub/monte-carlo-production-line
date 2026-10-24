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
