import pandas as pd
import matplotlib.pyplot as plt

Vorname_list = [
    "Adelina",
    "Johannes",
    "Lucia",
    "Anne",
    "Frank",
    "Fabian",
    "Robin",
    "Raphael",
    "Silke",
    "Kerstin",
    "Arno",
    "Marcus",
    "Alfons",
    "Nils",
    "Tim",
    "Achim",
    "Igor",
    "Lukas"
]

def main():
    global Nr_Zeichner

    df_origin = pd.read_excel("Personalplanung Engineering - 2025.xlsx", sheet_name="Kapazität intern 2025", header=1)
    df_origin = pd.DataFrame(df_origin)

    df = df_origin.drop(df_origin.index[-7:], inplace=False)
    df = df.drop(df_origin.index[:4])
    df = df.drop(columns=['Unnamed: 0', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 371'])

    df.rename(columns={'heute': 'Name', 'Unnamed: 2': 'Vorname'}, inplace=True)

    df.columns = df.columns.astype(str)
    df.columns = [col.replace(' 00:00:00', '') if i >= 2 and i < 367 else col for i, col in enumerate(df.columns)]
    df = df.fillna('X')

    df = df[df['Vorname'].isin(Vorname_list)]

    df.iloc[:, 2:] = df.iloc[:, 2:].astype(str)
    df.reset_index(drop=True, inplace=True)

    Nr_Zeichner = len(df)

    workdays_monthly(df, 8)
    employee_working_days(df, "Fabian")
    check_given_time(df, "2025-07-01", "2025-12-31")
    print("plot finished")


def workdays_monthly(df, month):
    month_columns = [col for col in df.columns[2:] if col.startswith(f"2025-{month:02d}")]
    
    df['Working Days'] = df.loc[0:Nr_Zeichner+1, month_columns].apply(lambda row: row.isin(['h', '1']).sum(), axis=1)
    
    plt.figure(figsize=(10, 6))
    bars = plt.bar(df.loc[0:17, 'Vorname'], df.loc[0:17, 'Working Days'], color='skyblue')
    plt.xlabel('Employee')
    plt.ylabel('Number of Working Days')
    plt.title(f'Number of Working Days in {month:02d}/2025')
    plt.xticks(rotation=45)
    
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, yval + 0.1, int(yval), ha='center', va='bottom')
    
    plt.tight_layout()
    # plt.savefig('workdays_monthly.png')
    plt.show(block=False)



def employee_working_days(df, employee_name):
    employee_row = df[df['Vorname'] == employee_name]
    
    working_days_per_month = {}
    
    for month in range(1, 13):
        month_columns = [col for col in df.columns[2:] if col.startswith(f"2025-{month:02d}")]
        
        working_days = employee_row[month_columns].apply(lambda row: row.isin(['h', '1']).sum(), axis=1).values[0]
        
        working_days_per_month[f"2025-{month:02d}"] = working_days
    
    plt.figure(figsize=(10, 6))
    plt.bar(working_days_per_month.keys(), working_days_per_month.values(), color='skyblue')
    plt.xlabel('Month')
    plt.ylabel('Number of Working Days')
    plt.title(f'Number of Working Days for {employee_name} in 2025')
    plt.xticks(rotation=45)
    
    for month, days in working_days_per_month.items():
        plt.text(month, days + 0.1, int(days), ha='center', va='bottom')
    
    plt.tight_layout()
    # plt.savefig('employee_working_days.png')   
    plt.show(block=False)


def check_given_time(start_time, end_time):
    time_columns = [col for col in df.columns[2:] if start_time <= col <= end_time]

    df['Working Days'] = df.loc[0:Nr_Zeichner+1, time_columns].apply(lambda row: row.isin(['h', '1']).sum(), axis=1)

    plt.figure(figsize=(10, 6))
    bars = plt.bar(df.loc[0:Nr_Zeichner+1, 'Vorname'], df.loc[0:Nr_Zeichner+1, 'Working Days'], color='skyblue')
    plt.xlabel('Employee')
    plt.ylabel('Number of Working Days')
    plt.title(f'Number of Working Days in {start_time} -- {end_time}')
    plt.xticks(rotation=45)
    

    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, yval + 0.1, int(yval), ha='center', va='bottom')
    
    plt.tight_layout()
    plt.show(block=False)


if __name__ == "__main__":
    main()