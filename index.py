import pandas as pd # type: ignore
import numpy as np # type: ignore
from openpyxl import load_workbook # type: ignore

# Define inventory policy parameters
s = 120  # Reorder point
S = 180  # Order-up-to level
Q = 60   # Order quantity
package_size = 10  # Package size
lead_time = 2  # Lead time

# Define Excel file name
file_path = "P:\Inventory_management\Inventory_Management\Inventory1.xlsx"
# Load data from Excel file
df = pd.read_excel(file_path)

# Get daily demand input from the user
daily_demand = float(input("Enter the daily demand: "))

def simulate(df, s, S, Q, daily_demand, order_rec):
    # Convert 'Inv Pos' column to numeric type
    df['Inv Pos'] = pd.to_numeric(df['Inv Pos'])

    # Simulation formulas
    df['Begg.Inv'] = df['End Inv'].shift(1).fillna(0)
    df['Order qty'] = np.where((df['Begg.Inv'] - daily_demand) < s, s - (df['Begg.Inv'] - daily_demand), 0)
    df['End Inv'] = df['Begg.Inv'] + order_rec - daily_demand
    df['Inv Pos'] = df['Begg.Inv'] + df['Order qty'].shift(1).fillna(0)
    df['Shortage'] = np.where(df['End Inv'] < 0, np.abs(df['End Inv']), 0)
    df['Total Cost'] = df['Order qty'] + df['End Inv'] + df['Shortage']
    cost = df['Total Cost'].sum()

    return df, cost

def local_search(df, s, S, Q, daily_demand):
    # Initial total cost before optimization
    initial_cost = df['Total Cost'].sum()

    # Step 1: Fixed Q
    while True:
        # Increase reorder point
        s_prime = s + 1
        S_prime = s_prime + Q
        df_s_prime_S_prime, cost_s_prime_S_prime = simulate(df.copy(), s_prime, S_prime, Q, daily_demand, 0)

        if cost_s_prime_S_prime < initial_cost:
            s = s_prime
            S = S_prime
            df = df_s_prime_S_prime
            initial_cost = cost_s_prime_S_prime
        else:
            # Decrease reorder point
            s_prime = s - 1
            S_prime = s_prime + Q
            df_s_prime_S_prime, cost_s_prime_S_prime = simulate(df.copy(), s_prime, S_prime, Q, daily_demand, 0)

            if cost_s_prime_S_prime < initial_cost:
                s = s_prime
                S = S_prime
                df = df_s_prime_S_prime
                initial_cost = cost_s_prime_S_prime
            else:
                break

    # Step 2: Variable Q
    r = min(S - s, s - 0)
    S_prime = S + r
    Q_prime = S_prime - s
    df_S_prime, cost_S_prime = simulate(df.copy(), s, S_prime, Q_prime, daily_demand, 0)

    if cost_S_prime < initial_cost:
        S = S_prime
        df = df_S_prime
        initial_cost = cost_S_prime

    s_prime = s - r
    Q_prime = S - s_prime
    df_s_prime, cost_s_prime = simulate(df.copy(), s_prime, S, Q_prime, daily_demand, 0)

    if cost_s_prime < initial_cost:
        s = s_prime
        df = df_s_prime
        initial_cost = cost_s_prime

    # Determine effectiveness of the policy
    if initial_cost != 0:
        effectiveness = (initial_cost - df['Total Cost'].sum()) / initial_cost * 100
    else:
        effectiveness = 0

    return s, S, Q, effectiveness, df

# Run local search algorithm
s, S, Q, effectiveness, df = local_search(df, s, S, Q, daily_demand)

# Print optimized parameters and effectiveness
print("Optimized parameters:")
print(f"Reorder point (s): {s}")
print(f"Order-up-to level (S): {S}")
print(f"Order quantity (Q): {Q}")
print(f"Effectiveness of the policy: {effectiveness:.2f}%")

# Get order receipts from the user
order_rec = float(input("Enter the order quantity received: "))

# Update inventory data in DataFrame

# Define a dictionary to map the current days to the next corresponding days
next_days_mapping = {'Monday': 'Tuesday', 'Tuesday': 'Wednesday', 'Wednesday': 'Thursday', 'Thursday': 'Friday',
                     'Friday': 'Saturday', 'Saturday': 'Sunday', 'Sunday': 'Monday'}

# Update the 'Week Days' column with the next corresponding day
# Add the new row with the updated data
next_day = df.iloc[-1]['Week Days']
next_day = next_days_mapping.get(next_day, 'Monday')  # Get the next day based on the mapping
inv_pos=df.iloc[-1]['End Inv']+df.iloc[-1]['Order qty']
end_inv=df.iloc[-1]['End Inv']+order_rec-daily_demand
shortage=S-df.iloc[-1]['End Inv']+order_rec
if inv_pos<120:
    place_order='Yes'
    Order_qty=S-inv_pos
else:
    place_order='No'
    Order_qty=0
new_data = {'Demand':daily_demand,'Week Days': next_day, 'Begg.Inv': df.iloc[-1]['End Inv'],
            'Inv Pos': inv_pos,'Place Order?':place_order, 'Order qty': Order_qty, 'Qty rec': order_rec, 'End Inv': end_inv,
            'Shortage': shortage, 'Total Cost': 0}
new_df = pd.DataFrame(new_data, index=[0])
df = pd.concat([df, new_df], ignore_index=True)

# Save the updated data to the Excel file using openpyxl
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, index=False, startrow=0, startcol=0, header=True)