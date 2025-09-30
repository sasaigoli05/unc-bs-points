# Sangeet Points Calculator

Compute **total points** for each member from a single Excel workbook and write the results to a final sheet called **`Points Standing`**.

> **Scoring Logic**
> - **Meeting attendance:** Each attended meeting is worth `1 / D` points  
>   where `D = (# of subgroups the member is in) + (1 if GBM is included)`.  
>   “Attended” = `Present`, `Tardy`, `Tardy (Excused)`, `Excessive Tardy`.  
>   Absences (excused/unexcused) = 0 points.
> - **Tabling:** `+1` point per tabling count.
> - **Gigs:** `+2` points per gig/performance count.
> - Final output contains **only two columns**: `Member` and `Total_Points`.

## Setup

# Install dependencies
pip install -r requirements.txt
