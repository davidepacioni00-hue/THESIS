import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter


# ============================================================
# 1. PATHS
# ============================================================

SIMULATION_PATH = r'../Simulations/simulation_(2)_BE-NUC'

EXCEL_FILE = os.path.join(
    SIMULATION_PATH,
    'UNLEASH_BE_NUC_FINAL_RESULTS.xlsx'
)

GRAPH_FOLDER = os.path.join(
    SIMULATION_PATH,
    'Graphs'
)

os.makedirs(
    GRAPH_FOLDER,
    exist_ok=True
)


# ============================================================
# 2. GENERAL GRAPH SETTINGS
# ============================================================

plt.rcParams.update({
    'font.size': 10,
    'axes.titlesize': 14,
    'axes.labelsize': 11,
    'legend.fontsize': 9,
    'figure.dpi': 120
})


def save_graph(fig, filename):
    """
    Save a figure in the Graphs folder.
    """

    path = os.path.join(
        GRAPH_FOLDER,
        filename
    )

    fig.savefig(
        path,
        dpi=300,
        bbox_inches='tight'
    )

    plt.close(fig)

    print(
        f"CREATED: {filename}"
    )


def format_annual_axis(ax):
    """
    Format annual time axis using month names.
    """

    ax.xaxis.set_major_locator(
        mdates.MonthLocator()
    )

    ax.xaxis.set_major_formatter(
        mdates.DateFormatter('%b')
    )

    ax.grid(
        True,
        alpha=0.25
    )


def format_week_axis(ax):
    """
    Format weekly time axis.
    """

    ax.xaxis.set_major_locator(
        mdates.DayLocator()
    )

    ax.xaxis.set_major_formatter(
        mdates.DateFormatter('%d %b')
    )

    ax.grid(
        True,
        alpha=0.25
    )


# ============================================================
# 3. LOAD FINAL EXCEL RESULTS
# ============================================================

print()
print("=" * 70)
print("LOADING FINAL UNLEASH RESULTS")
print("=" * 70)


vpp = pd.read_excel(
    EXCEL_FILE,
    sheet_name='VPP_Boundary',
    header=3
)

dispatch = pd.read_excel(
    EXCEL_FILE,
    sheet_name='Dispatch',
    header=3
)

system = pd.read_excel(
    EXCEL_FILE,
    sheet_name='System',
    header=3
)

reserves = pd.read_excel(
    EXCEL_FILE,
    sheet_name='Reserves',
    header=3
)

units = pd.read_excel(
    EXCEL_FILE,
    sheet_name='Units_Summary',
    header=3
)


# ============================================================
# 4. DATE / TIME
# ============================================================

for df in [
    vpp,
    dispatch,
    system,
    reserves
]:

    df['Date / Time'] = pd.to_datetime(
        df['Date / Time']
    )

    df.sort_values(
        'Date / Time',
        inplace=True
    )

    df.reset_index(
        drop=True,
        inplace=True
    )


print("All Excel sheets loaded successfully.")
print("Hourly values:", len(vpp))


# ============================================================
# 5. COLUMN NAMES
# ============================================================

VPP_CONSUMPTION = (
    'VPP Electricity Consumption [MW]'
)

VPP_POWERX = (
    'VPP Useful PowerX [MW]'
)

BOUNDARY_DEMAND = (
    'Boundary Sector Demand [MW]'
)

STORAGE_NET_INPUT = (
    'SectorX Storage Net Input [MW]'
)

STORAGE_SOC = (
    'SectorX Storage SOC [%]'
)

STORED_ENERGY = (
    'SectorX Stored Energy [MWh]'
)

SELF_DISCHARGE = (
    'SectorX Self-Discharge [MW]'
)

BASE_DEMAND = (
    'Base Electricity Demand [MW]'
)

TOTAL_DEMAND = (
    'Total Electricity Demand incl. VPP [MW]'
)

RESIDUAL_LOAD = (
    'Residual Load [MW]'
)

SYSTEM_COST = (
    'System Cost [EUR]'
)

CO2 = (
    'CO2 Emissions [tCO2]'
)

LOAD_SHEDDING = (
    'Load Shedding [MW]'
)

ENS = (
    'Energy Not Served [MWh]'
)


# ============================================================
# 6. FIND REPRESENTATIVE WEEK
# ============================================================

annual_mean_vpp = (
    vpp[VPP_CONSUMPTION].mean()
)

candidate_weeks = []


for start in range(
    0,
    len(vpp) - 167,
    168
):

    week = vpp.iloc[
        start:start + 168
    ]

    week_mean = (
        week[VPP_CONSUMPTION].mean()
    )

    difference = abs(
        week_mean
        - annual_mean_vpp
    )

    candidate_weeks.append(
        (
            difference,
            start
        )
    )


best_difference, best_start = min(
    candidate_weeks,
    key=lambda x: x[0]
)


week_start = best_start
week_end = best_start + 168


week_vpp = vpp.iloc[
    week_start:week_end
].copy()

week_dispatch = dispatch.iloc[
    week_start:week_end
].copy()

week_reserves = reserves.iloc[
    week_start:week_end
].copy()


week_start_date = (
    week_vpp[
        'Date / Time'
    ].iloc[0]
)

week_end_date = (
    week_vpp[
        'Date / Time'
    ].iloc[-1]
)


print()
print("Representative week:")
print(
    week_start_date.strftime(
        '%d/%m/%Y %H:%M'
    ),
    "->",
    week_end_date.strftime(
        '%d/%m/%Y %H:%M'
    )
)

print(
    "Annual average VPP consumption:",
    f"{annual_mean_vpp:.2f} MW"
)

print(
    "Representative week average:",
    f"{week_vpp[VPP_CONSUMPTION].mean():.2f} MW"
)


# ============================================================
# 7. DAILY AND MONTHLY DATA
# ============================================================

vpp_indexed = (
    vpp
    .set_index('Date / Time')
)

dispatch_indexed = (
    dispatch
    .set_index('Date / Time')
)

system_indexed = (
    system
    .set_index('Date / Time')
)

reserves_indexed = (
    reserves
    .set_index('Date / Time')
)


vpp_daily = (
    vpp_indexed
    .select_dtypes(include=[np.number])
    .resample('D')
    .mean()
)

dispatch_daily = (
    dispatch_indexed
    .select_dtypes(include=[np.number])
    .resample('D')
    .mean()
)

# ============================================================
# GRAPH 01
# VPP ANNUAL OPERATION
# ============================================================

fig, ax1 = plt.subplots(
    figsize=(15, 7)
)


ax1.plot(
    vpp_daily.index,
    vpp_daily[VPP_CONSUMPTION],
    linewidth=1.3,
    label='VPP Electricity Consumption'
)

ax1.plot(
    vpp_daily.index,
    vpp_daily[VPP_POWERX],
    linewidth=1.3,
    label='VPP Useful PowerX'
)

ax1.plot(
    vpp_daily.index,
    vpp_daily[BOUNDARY_DEMAND],
    linewidth=1.3,
    label='Boundary Sector Demand'
)


ax1.set_ylabel(
    'Daily Average Power [MW]'
)

ax1.set_xlabel(
    '2023'
)


ax2 = ax1.twinx()


ax2.plot(
    vpp_daily.index,
    vpp_daily[STORAGE_SOC] * 100,
    linewidth=1.8,
    linestyle='--',
    label='SectorX Storage SOC'
)


ax2.set_ylabel(
    'Storage SOC [%]'
)

ax2.set_ylim(
    0,
    100
)


format_annual_axis(
    ax1
)


lines1, labels1 = (
    ax1.get_legend_handles_labels()
)

lines2, labels2 = (
    ax2.get_legend_handles_labels()
)


ax1.legend(
    lines1 + lines2,
    labels1 + labels2,
    loc='upper center',
    bbox_to_anchor=(0.5, -0.12),
    ncol=4,
    frameon=False
)


ax1.set_title(
    'UNLEASH – Annual VPP and Boundary Sector Operation\n'
    'Daily Average Values – 2023',
    fontweight='bold'
)


fig.tight_layout()

fig.subplots_adjust(
    bottom=0.18
)


save_graph(
    fig,
    '01_VPP_Annual_Operation.png'
)


# ============================================================
# GRAPH 02
# VPP REPRESENTATIVE WEEK
# ============================================================

fig, ax1 = plt.subplots(
    figsize=(15, 7)
)


ax1.plot(
    week_vpp['Date / Time'],
    week_vpp[VPP_CONSUMPTION],
    linewidth=1.6,
    label='VPP Electricity Consumption'
)

ax1.plot(
    week_vpp['Date / Time'],
    week_vpp[VPP_POWERX],
    linewidth=1.6,
    label='VPP Useful PowerX'
)

ax1.plot(
    week_vpp['Date / Time'],
    week_vpp[BOUNDARY_DEMAND],
    linewidth=1.6,
    label='Boundary Sector Demand'
)


ax1.set_ylabel(
    'Power [MW]'
)

ax1.set_xlabel(
    'Date / Time'
)


ax2 = ax1.twinx()


ax2.plot(
    week_vpp['Date / Time'],
    week_vpp[STORAGE_SOC] * 100,
    linewidth=2,
    linestyle='--',
    label='SectorX Storage SOC'
)


ax2.set_ylabel(
    'Storage SOC [%]'
)

ax2.set_ylim(
    0,
    100
)


format_week_axis(
    ax1
)


lines1, labels1 = (
    ax1.get_legend_handles_labels()
)

lines2, labels2 = (
    ax2.get_legend_handles_labels()
)


ax1.legend(
    lines1 + lines2,
    labels1 + labels2,
    loc='upper center',
    bbox_to_anchor=(0.5, -0.13),
    ncol=4,
    frameon=False
)


ax1.set_title(
    (
        'UNLEASH – VPP and Boundary Sector Operation\n'
        f'Representative Week: '
        f'{week_start_date:%d %B %Y} – '
        f'{week_end_date:%d %B %Y}'
    ),
    fontweight='bold'
)


fig.tight_layout()

fig.subplots_adjust(
    bottom=0.20
)


save_graph(
    fig,
    '02_VPP_Representative_Week.png'
)


# ============================================================
# GRAPH 03
# MONTHLY VPP ENERGY BALANCE
# ============================================================

monthly_vpp = (
    vpp_indexed[
        [
            VPP_CONSUMPTION,
            VPP_POWERX,
            BOUNDARY_DEMAND,
            SELF_DISCHARGE
        ]
    ]
    .resample('MS')
    .sum()
    / 1000
)


monthly_vpp.index = (
    monthly_vpp.index
    .strftime('%b')
)


monthly_vpp.columns = [
    'Electricity Consumption',
    'Useful PowerX',
    'Boundary Sector Demand',
    'Self-Discharge'
]


fig, ax = plt.subplots(
    figsize=(14, 7)
)


monthly_vpp.plot(
    kind='bar',
    ax=ax
)


ax.set_ylabel(
    'Monthly Energy [GWh]'
)

ax.set_xlabel(
    'Month'
)

ax.set_title(
    'UNLEASH – Monthly VPP / Boundary Sector Energy Balance – 2023',
    fontweight='bold'
)

ax.grid(
    axis='y',
    alpha=0.25
)

ax.legend(
    frameon=False,
    ncol=2
)

plt.xticks(
    rotation=0
)


fig.tight_layout()


save_graph(
    fig,
    '03_VPP_Monthly_Energy_Balance.png'
)


# ============================================================
# GRAPH 04
# ANNUAL ELECTRICITY DEMAND
# ============================================================

fig, ax1 = plt.subplots(
    figsize=(15, 7)
)


ax1.plot(
    dispatch_daily.index,
    dispatch_daily[BASE_DEMAND],
    linewidth=1.3,
    label='Base Electricity Demand'
)

ax1.plot(
    dispatch_daily.index,
    dispatch_daily[TOTAL_DEMAND],
    linewidth=1.3,
    label='Total Demand incl. VPP'
)


ax1.set_ylabel(
    'Daily Average Demand [MW]'
)

ax1.set_xlabel(
    '2023'
)


ax2 = ax1.twinx()


ax2.plot(
    dispatch_daily.index,
    vpp_daily[VPP_CONSUMPTION],
    linewidth=1.2,
    linestyle='--',
    label='VPP Electricity Consumption'
)


ax2.set_ylabel(
    'VPP Consumption [MW]'
)


format_annual_axis(
    ax1
)


lines1, labels1 = (
    ax1.get_legend_handles_labels()
)

lines2, labels2 = (
    ax2.get_legend_handles_labels()
)


ax1.legend(
    lines1 + lines2,
    labels1 + labels2,
    loc='upper center',
    bbox_to_anchor=(0.5, -0.12),
    ncol=3,
    frameon=False
)


ax1.set_title(
    'Belgian Electricity Demand with UNLEASH VPP – 2023\n'
    'Daily Average Values',
    fontweight='bold'
)


fig.tight_layout()

fig.subplots_adjust(
    bottom=0.18
)


save_graph(
    fig,
    '04_Electricity_Demand_Annual.png'
)


# ============================================================
# GRAPH 05
# ELECTRICITY DEMAND – REPRESENTATIVE WEEK
# ============================================================

fig, ax1 = plt.subplots(
    figsize=(15, 7)
)


ax1.plot(
    week_dispatch['Date / Time'],
    week_dispatch[BASE_DEMAND],
    linewidth=1.5,
    label='Base Electricity Demand'
)

ax1.plot(
    week_dispatch['Date / Time'],
    week_dispatch[TOTAL_DEMAND],
    linewidth=1.5,
    label='Total Demand incl. VPP'
)


ax1.set_ylabel(
    'Electricity Demand [MW]'
)

ax1.set_xlabel(
    'Date / Time'
)


ax2 = ax1.twinx()


ax2.plot(
    week_vpp['Date / Time'],
    week_vpp[VPP_CONSUMPTION],
    linewidth=1.5,
    linestyle='--',
    label='VPP Electricity Consumption'
)


ax2.set_ylabel(
    'VPP Consumption [MW]'
)


format_week_axis(
    ax1
)


lines1, labels1 = (
    ax1.get_legend_handles_labels()
)

lines2, labels2 = (
    ax2.get_legend_handles_labels()
)


ax1.legend(
    lines1 + lines2,
    labels1 + labels2,
    loc='upper center',
    bbox_to_anchor=(0.5, -0.13),
    ncol=3,
    frameon=False
)


ax1.set_title(
    (
        'Electricity Demand and VPP Operation\n'
        f'Representative Week: '
        f'{week_start_date:%d %B} – '
        f'{week_end_date:%d %B %Y}'
    ),
    fontweight='bold'
)


fig.tight_layout()

fig.subplots_adjust(
    bottom=0.20
)


save_graph(
    fig,
    '05_Electricity_Demand_Representative_Week.png'
)


# ============================================================
# 8. MAP DISPATCH UNITS TO FUELS
# ============================================================

generation_columns = [
    col
    for col in dispatch.columns
    if str(col).startswith(
        'Generation | '
    )
]


unit_to_fuel = {}


if (
    'Unit' in units.columns
    and 'Fuel' in units.columns
):

    for _, row in units.iterrows():

        unit_to_fuel[
            str(row['Unit'])
        ] = str(row['Fuel'])


generation_by_fuel = pd.DataFrame(
    index=dispatch_indexed.index
)


for column in generation_columns:

    unit_name = (
        column
        .replace(
            'Generation | ',
            ''
        )
        .replace(
            ' [MW]',
            ''
        )
    )


    fuel = unit_to_fuel.get(
        unit_name,
        'Other'
    )


    if fuel not in generation_by_fuel.columns:

        generation_by_fuel[
            fuel
        ] = 0.0


    generation_by_fuel[
        fuel
    ] = (
        generation_by_fuel[fuel]
        +
        pd.to_numeric(
            dispatch_indexed[column],
            errors='coerce'
        ).fillna(0)
    )


# Remove fuels with zero generation

generation_by_fuel = (
    generation_by_fuel.loc[
        :,
        generation_by_fuel.sum() > 0
    ]
)


# ============================================================
# GRAPH 06
# MONTHLY GENERATION BY FUEL
# ============================================================

monthly_generation = (
    generation_by_fuel
    .resample('MS')
    .sum()
    / 1_000_000
)


monthly_generation.index = (
    monthly_generation.index
    .strftime('%b')
)


fig, ax = plt.subplots(
    figsize=(15, 8)
)


monthly_generation.plot(
    kind='bar',
    stacked=True,
    ax=ax
)


ax.set_ylabel(
    'Electricity Generation [TWh]'
)

ax.set_xlabel(
    'Month'
)

ax.set_title(
    'Belgian Monthly Electricity Generation by Fuel – 2023',
    fontweight='bold'
)

ax.grid(
    axis='y',
    alpha=0.25
)

ax.legend(
    title='Fuel',
    bbox_to_anchor=(1.02, 1),
    loc='upper left',
    frameon=False
)

plt.xticks(
    rotation=0
)


fig.tight_layout()


save_graph(
    fig,
    '06_Monthly_Generation_by_Fuel.png'
)


# ============================================================
# GRAPH 07
# ANNUAL GENERATION BY FUEL
# ============================================================

if (
    'Fuel' in units.columns
    and
    'Annual Generation [MWh]'
    in units.columns
):

    annual_fuel = (

        units.groupby(
            'Fuel'
        )[
            'Annual Generation [MWh]'
        ]

        .sum()

        .sort_values()

        / 1_000_000
    )


    annual_fuel = (
        annual_fuel[
            annual_fuel > 0
        ]
    )


    fig, ax = plt.subplots(
        figsize=(11, 7)
    )


    annual_fuel.plot(
        kind='barh',
        ax=ax
    )


    ax.set_xlabel(
        'Annual Electricity Generation [TWh]'
    )

    ax.set_ylabel(
        'Fuel'
    )

    ax.set_title(
        'Belgian Annual Electricity Generation by Fuel – 2023',
        fontweight='bold'
    )

    ax.grid(
        axis='x',
        alpha=0.25
    )


    fig.tight_layout()


    save_graph(
        fig,
        '07_Annual_Generation_by_Fuel.png'
    )


# ============================================================
# GRAPH 08
# ANNUAL GENERATION BY UNIT
# ============================================================

if (
    'Unit' in units.columns
    and
    'Annual Generation [MWh]'
    in units.columns
):

    unit_generation = (
        units[
            [
                'Unit',
                'Annual Generation [MWh]'
            ]
        ]
        .copy()
    )


    unit_generation[
        'Annual Generation [TWh]'
    ] = (

        unit_generation[
            'Annual Generation [MWh]'
        ]
        / 1_000_000
    )


    unit_generation = (
        unit_generation[
            unit_generation[
                'Annual Generation [TWh]'
            ] > 0
        ]
        .sort_values(
            'Annual Generation [TWh]'
        )
    )


    fig_height = max(
        8,
        len(unit_generation) * 0.38
    )


    fig, ax = plt.subplots(
        figsize=(13, fig_height)
    )


    ax.barh(
        unit_generation['Unit'],
        unit_generation[
            'Annual Generation [TWh]'
        ]
    )


    ax.set_xlabel(
        'Annual Electricity Generation [TWh]'
    )

    ax.set_ylabel(
        'Generating Unit'
    )

    ax.set_title(
        'Annual Electricity Generation by Unit – 2023',
        fontweight='bold'
    )

    ax.grid(
        axis='x',
        alpha=0.25
    )


    fig.tight_layout()


    save_graph(
        fig,
        '08_Annual_Generation_by_Unit.png'
    )


# ============================================================
# GRAPH 09
# MONTHLY CO2 EMISSIONS
# ============================================================

if CO2 in system_indexed.columns:

    monthly_co2 = (
        system_indexed[CO2]
        .resample('MS')
        .sum()
        / 1_000_000
    )


    month_labels = (
        monthly_co2.index
        .strftime('%b')
    )


    fig, ax = plt.subplots(
        figsize=(12, 6)
    )


    ax.bar(
        month_labels,
        monthly_co2.values
    )


    ax.set_ylabel(
        'CO2 Emissions [MtCO2]'
    )

    ax.set_xlabel(
        'Month'
    )

    ax.set_title(
        'Monthly Power-System CO2 Emissions – 2023',
        fontweight='bold'
    )

    ax.grid(
        axis='y',
        alpha=0.25
    )


    fig.tight_layout()


    save_graph(
        fig,
        '09_Monthly_CO2_Emissions.png'
    )


# ============================================================
# GRAPH 10
# MONTHLY SYSTEM COST
# ============================================================

if SYSTEM_COST in system_indexed.columns:

    monthly_cost = (
        system_indexed[
            SYSTEM_COST
        ]
        .resample('MS')
        .sum()
        / 1_000_000_000
    )


    month_labels = (
        monthly_cost.index
        .strftime('%b')
    )


    fig, ax = plt.subplots(
        figsize=(12, 6)
    )


    ax.bar(
        month_labels,
        monthly_cost.values
    )


    ax.set_ylabel(
        'System Cost [billion EUR]'
    )

    ax.set_xlabel(
        'Month'
    )

    ax.set_title(
        'Monthly System Cost – 2023',
        fontweight='bold'
    )

    ax.grid(
        axis='y',
        alpha=0.25
    )


    fig.tight_layout()


    save_graph(
        fig,
        '10_Monthly_System_Cost.png'
    )


# ============================================================
# GRAPH 11
# MONTHLY RELIABILITY
# ============================================================

if ENS in system_indexed.columns:

    monthly_ens = (
        system_indexed[ENS]
        .resample('MS')
        .sum()
        / 1000
    )


    scarcity_hours = (
        (system_indexed[ENS] > 0)
        .astype(int)
        .resample('MS')
        .sum()
    )


    month_labels = (
        monthly_ens.index
        .strftime('%b')
    )


    fig, ax1 = plt.subplots(
        figsize=(13, 7)
    )


    ax1.bar(
        month_labels,
        monthly_ens.values,
        label='Energy Not Served'
    )


    ax1.set_ylabel(
        'Energy Not Served [GWh]'
    )

    ax1.set_xlabel(
        'Month'
    )

    ax1.grid(
        axis='y',
        alpha=0.25
    )


    ax2 = ax1.twinx()


    ax2.plot(
        month_labels,
        scarcity_hours.values,
        marker='o',
        linewidth=2,
        label='Hours with Energy Not Served'
    )


    ax2.set_ylabel(
        'Hours with Energy Not Served [h]'
    )


    lines1, labels1 = (
        ax1.get_legend_handles_labels()
    )

    lines2, labels2 = (
        ax2.get_legend_handles_labels()
    )


    ax1.legend(
        lines1 + lines2,
        labels1 + labels2,
        loc='upper center',
        bbox_to_anchor=(0.5, -0.10),
        ncol=2,
        frameon=False
    )


    ax1.set_title(
        'Monthly Reliability Indicators – 2023',
        fontweight='bold'
    )


    fig.tight_layout()

    fig.subplots_adjust(
        bottom=0.16
    )


    save_graph(
        fig,
        '11_Monthly_Reliability.png'
    )


# ============================================================
# GRAPH 12
# ANNUAL RESERVE REQUIREMENTS / AVAILABILITY
# ============================================================

reserve_columns = [

    '2U Reserve Requirement [MW]',
    '2D Reserve Requirement [MW]',
    '3U Reserve Requirement [MW]',

    'Available 2U Reserve [MW]',
    'Available 2D Reserve [MW]',
    'Available 3U Reserve [MW]'
]


available_reserve_columns = [
    col
    for col in reserve_columns
    if col in reserves_indexed.columns
]


if available_reserve_columns:

    monthly_reserves = (
        reserves_indexed[
            available_reserve_columns
        ]
        .resample('MS')
        .mean()
    )


    fig, ax = plt.subplots(
        figsize=(15, 7)
    )


    for column in monthly_reserves.columns:

        ax.plot(
            monthly_reserves.index,
            monthly_reserves[column],
            marker='o',
            linewidth=1.5,
            label=column.replace(
                ' [MW]',
                ''
            )
        )


    ax.set_ylabel(
        'Monthly Average Reserve [MW]'
    )

    ax.set_xlabel(
        '2023'
    )

    ax.set_title(
        'Monthly Average Reserve Requirements and Availability – 2023',
        fontweight='bold'
    )


    format_annual_axis(
        ax
    )


    ax.legend(
        bbox_to_anchor=(1.02, 1),
        loc='upper left',
        frameon=False
    )


    fig.tight_layout()


    save_graph(
        fig,
        '12_Annual_Reserve_Requirements.png'
    )


# ============================================================
# GRAPH 13
# RESERVES – REPRESENTATIVE WEEK
# ============================================================

week_reserve_columns = [
    col
    for col in reserve_columns
    if col in week_reserves.columns
]


if week_reserve_columns:

    fig, ax = plt.subplots(
        figsize=(15, 7)
    )


    for column in week_reserve_columns:

        ax.plot(
            week_reserves[
                'Date / Time'
            ],
            week_reserves[column],
            linewidth=1.3,
            label=column.replace(
                ' [MW]',
                ''
            )
        )


    ax.set_ylabel(
        'Reserve [MW]'
    )

    ax.set_xlabel(
        'Date / Time'
    )


    format_week_axis(
        ax
    )


    ax.legend(
        bbox_to_anchor=(1.02, 1),
        loc='upper left',
        frameon=False
    )


    ax.set_title(
        (
            'Reserve Requirements and Availability\n'
            f'Representative Week: '
            f'{week_start_date:%d %B} – '
            f'{week_end_date:%d %B %Y}'
        ),
        fontweight='bold'
    )


    fig.tight_layout()


    save_graph(
        fig,
        '13_Reserve_Representative_Week.png'
    )


# ============================================================
# GRAPH 14
# CAPACITY FACTOR BY UNIT
# ============================================================

if (
    'Unit' in units.columns
    and
    'Capacity Factor [%]' in units.columns
):

    capacity_factor = (
        units[
            [
                'Unit',
                'Capacity Factor [%]'
            ]
        ]
        .copy()
    )


    capacity_factor[
        'Capacity Factor [%]'
    ] = (
        pd.to_numeric(
            capacity_factor[
                'Capacity Factor [%]'
            ],
            errors='coerce'
        )
        * 100
    )


    capacity_factor = (
        capacity_factor[
            capacity_factor[
                'Capacity Factor [%]'
            ] > 0
        ]
        .sort_values(
            'Capacity Factor [%]'
        )
    )


    fig_height = max(
        8,
        len(capacity_factor) * 0.38
    )


    fig, ax = plt.subplots(
        figsize=(13, fig_height)
    )


    ax.barh(
        capacity_factor['Unit'],
        capacity_factor[
            'Capacity Factor [%]'
        ]
    )


    ax.set_xlabel(
        'Capacity Factor [%]'
    )

    ax.set_ylabel(
        'Generating Unit'
    )

    ax.set_xlim(
        0,
        max(
            100,
            capacity_factor[
                'Capacity Factor [%]'
            ].max() * 1.05
        )
    )


    ax.set_title(
        'Generating Unit Capacity Factors – 2023',
        fontweight='bold'
    )

    ax.grid(
        axis='x',
        alpha=0.25
    )


    fig.tight_layout()


    save_graph(
        fig,
        '14_Unit_Capacity_Factors.png'
    )


# ============================================================
# GRAPH 15
# ANNUAL CO2 EMISSIONS BY UNIT
# ============================================================

if (
    'Unit' in units.columns
    and
    'Annual CO2 Emissions [tCO2]'
    in units.columns
):

    unit_emissions = (
        units[
            [
                'Unit',
                'Annual CO2 Emissions [tCO2]'
            ]
        ]
        .copy()
    )


    unit_emissions[
        'Annual CO2 Emissions [MtCO2]'
    ] = (

        pd.to_numeric(
            unit_emissions[
                'Annual CO2 Emissions [tCO2]'
            ],
            errors='coerce'
        )

        / 1_000_000
    )


    unit_emissions = (
        unit_emissions[
            unit_emissions[
                'Annual CO2 Emissions [MtCO2]'
            ] > 0
        ]
        .sort_values(
            'Annual CO2 Emissions [MtCO2]'
        )
    )


    if len(unit_emissions) > 0:

        fig_height = max(
            7,
            len(unit_emissions) * 0.42
        )


        fig, ax = plt.subplots(
            figsize=(13, fig_height)
        )


        ax.barh(
            unit_emissions['Unit'],
            unit_emissions[
                'Annual CO2 Emissions [MtCO2]'
            ]
        )


        ax.set_xlabel(
            'Annual CO2 Emissions [MtCO2]'
        )

        ax.set_ylabel(
            'Generating Unit'
        )

        ax.set_title(
            'Annual CO2 Emissions by Generating Unit – 2023',
            fontweight='bold'
        )

        ax.grid(
            axis='x',
            alpha=0.25
        )


        fig.tight_layout()


        save_graph(
            fig,
            '15_Annual_CO2_Emissions_by_Unit.png'
        )


# ============================================================
# FINAL MESSAGE
# ============================================================

print()
print("=" * 70)
print("ALL UNLEASH GRAPHS CREATED SUCCESSFULLY")
print("=" * 70)

print()
print("Graph folder:")
print(GRAPH_FOLDER)

print()
print("Representative week:")
print(
    f"{week_start_date:%d/%m/%Y}"
    " -> "
    f"{week_end_date:%d/%m/%Y}"
)

print()
print("Done.")