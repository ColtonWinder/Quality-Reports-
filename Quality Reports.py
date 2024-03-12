#!pip install dash
#!pip install dash-html-components
#!pip install dash-core-components
#!pip install pandas
#!pip install pandas openpyxl

import pandas as pd 
import dash
from dash import dash_table
from dash import dcc
from dash import html
import plotly.express as px
from dash.dependencies import Input, Output, State

df = pd.read_excel(r'V:\Quality\1 DMR Report\Inspection Log.xlsx')

df1 = pd.read_excel(r'V:\Quality\1 DMR Report\MO DMR Report Test.xlsx', sheet_name='MO')

df2 = pd.read_excel(r'V:\Quality\1 DMR Report\MO DMR Report Test.xlsx', sheet_name='EVANS')

df3 = pd.read_excel(r'V:\Quality\1 DMR Report\MO DMR Report Test.xlsx', sheet_name='NMB')

df5 = pd.read_excel(r'V:\Quality\1 DMR Report\MO DMR Report Test.xlsx', sheet_name='Defect Codes')

df['Date'] = pd.to_datetime(df['Date'])
df1['Date'] = pd.to_datetime(df1['Date'])
df2['Date'] = pd.to_datetime(df2['Date'])
df3['Date'] = pd.to_datetime(df3['Date'])

def extract_last_word(reason_code):

    processed_code = reason_code.replace("(PRE)", "").replace("(FINAL)", "").strip()

    words = processed_code.split()
    last_word = words[-1] if words else processed_code

    return last_word

df1['Area Type'] = df1['Reason Code'].apply(extract_last_word)
df1['Area Type'] = df1['Area Type'].replace('Assy', 'General Assy')
df1['Area Type'] = df1['Area Type'].replace('COAT', 'Conformal Coat')
df1['Area Type'] = df1['Area Type'].replace('CUT', 'Wirecut')
df1['Area Type'] = df1['Area Type'].replace('Density', 'High Density')
df1['Area Type'] = df1['Area Type'].replace('Development', 'PD')
df1['Area Type'] = df1['Area Type'].replace('FINAL', 'Mold')
df1['Area Type'] = df1['Area Type'].replace('Mold', 'Everflex Mold')
df1['Area Type'] = df1['Area Type'].replace('MOLD', 'Mold')
df1['Area Type'] = df1['Area Type'].replace('cut', 'Wirecut')

df2['Area Type'] = df2['Reason Code'].apply(extract_last_word)
df2['Area Type'] = df2['Area Type'].replace('Assy', 'General Assy')
df2['Area Type'] = df2['Area Type'].replace('COAT', 'Conformal Coat')
df2['Area Type'] = df2['Area Type'].replace('CUT', 'Wirecut')
df2['Area Type'] = df2['Area Type'].replace('Density', 'High Density')
df2['Area Type'] = df2['Area Type'].replace('Development', 'PD')
df2['Area Type'] = df2['Area Type'].replace('FINAL', 'Mold')
df2['Area Type'] = df2['Area Type'].replace('Mold', 'Everflex Mold')
df2['Area Type'] = df2['Area Type'].replace('MOLD', 'Mold')
df2['Area Type'] = df2['Area Type'].replace('cut', 'Wirecut')

df3['Area Type'] = df3['Reason Code'].apply(extract_last_word)
df3['Area Type'] = df3['Area Type'].replace('Assy', 'General Assy')
df3['Area Type'] = df3['Area Type'].replace('COAT', 'Conformal Coat')
df3['Area Type'] = df3['Area Type'].replace('CUT', 'Wirecut')
df3['Area Type'] = df3['Area Type'].replace('Density', 'High Density')
df3['Area Type'] = df3['Area Type'].replace('Development', 'PD')
df3['Area Type'] = df3['Area Type'].replace('FINAL', 'Mold')
df3['Area Type'] = df3['Area Type'].replace('Mold', 'Everflex Mold')
df3['Area Type'] = df3['Area Type'].replace('MOLD', 'Mold')
df3['Area Type'] = df3['Area Type'].replace('cut', 'Wirecut')


failure_code_descriptions = {
    'A01':  'Backshell damaged',
    'A02': 'Mounting hardware damaged',
    'A03': 'Installed wrong',
    'A04': 'Not per print',
    'A05': 'Missing',
    'A06': 'Tape loose, overlap',
    'A07': 'Improper bundle routing',
    'A08': 'FOD',
    'A09': 'No connector covers',
    'A10': 'Hardware loose',
    'A11': 'Plating damage',
    'B01': 'Braid bunched',
    'B02': 'Insufficient shield coverage',
    'B03': 'Not secured / Frayed',
    'B04': 'Braid damaged',
    'C01': 'Bent / Broken Pin',
    'C02': 'Connector Damaged',
    'C03': 'Connector Orientation',
    'C04': 'Recessed Pin',
    'C05': 'Improper crimp',
    'C06': 'Missing contact/filler pin(s)',
    'C07': 'Crimp not in proper location',
    'C08': 'Contact damaged',
    'C09': 'Wrong contact/mixed',
    'CC01': 'Cracking/Damage',
    'CC02': 'Insufficient Coverage/Voids',
    'D01': 'Documentation missing',
    'D02': 'DMR not signed at disposition',
    'D03': 'Missing Stamp',
    'D04': 'DMR not closed',
    'DIM01': 'Dimension Wrong',
    'E01': 'Electrical Leakage',
    'E02': 'Mis-wired',
    'E03': 'Electrical Open',
    'E04': 'Electrical Short',
    'E05': 'High Resistance',
    'E06': 'Intermittent',
    'ID01': 'Illegible',
    'ID02': 'Incorrect',
    'ID03': 'Location/Orientation',
    'ID04': 'ID missing',
    'ID05': 'ID flaking/peeling',
    'ID06': 'ID not shrunk',
    'INS02': 'Insulation damaged',
    'INS03': 'Insulation gap',
    'LA01': 'Lacing tie loose',
    'LA02': 'Lacing knot is incorrect',
    'LA03': 'Lacing routing is incorrect',
    'LA04': 'Lacing tie location is incorrect',
    'LA05': 'Wire missed in ties',
    'LA06': 'Frayed/length',
    'MD01': 'Broken Weave',
    'MD02': 'Compound in mating area',
    'MD03': 'Flash',
    'MD04': 'Ribbon not secure',
    'MD05': 'Voids in Compound',
    'MD06': 'Wire showing through compound',
    'MD07': 'Wires out of track',
    'MD08': 'Wires exposed on ribbon edge',
    'MD09': 'Compound soft',
    'MD10': 'Contact exposed',
    'MD11': 'Mold de-bond',
    'MD12': 'Reworked out of mold',
    'MD13': 'Wicking',
    'MD14': 'Mold out of dimension',
    'MD15': 'PSA peeling/lifting',
    'MD16': 'PSA damaged',
    'MT01': 'Damaged',
    'MT02': 'ID Wrong',
    'MT03': 'Kit Quantity missing',
    'MT04': 'Not per print',
    'S01': 'Cold Solder',
    'S02': 'Excessive Solder',
    'S03': 'Flux',
    'S04': 'Fractured solder',
    'S05': 'Insufficient Solder',
    'S06': 'Pin holes in solder',
    'S07': 'Wires Bird caged',
    'S08': 'Excess Solder Balls',
    'S09': 'Wrong strip length',
    'S10': 'Disturbed solder',
    'S11': 'Coppered solder',
    'S12': 'Insulation in solder',
    'SS01': 'Band intact',
    'SS02': 'Wire Protruding',
    'SS03': 'Wires out of lay',
    'SS04': 'Wires uneven',
    'SS05': 'Insufficient fillet',
    'SS06': 'Not sealed',
    'SB01': 'Damaged/split',
    'SB02': 'Deformed',
    'SB03': 'Location',
    'SB04': 'Missing',
    'SB05': 'Not Shrunk',
    'VI01': 'Damaged',
    'VI02': 'Not per print',
    'VI03': 'Not usable',
    'VI04': 'Over shipment',
    'VI05': 'Shelf Life',
    'W01': 'Wire Damaged',
    'W02': 'Wire Broken',
    'W03': 'Red Plague',
    'W04': 'Insufficient stress relief',
    'W05': 'Solder wicking',
    'W06': 'Wire out of cup',
    'W07': 'Wire out of lay',
    'WC01': 'Incorrect wire length',
    'WC02': 'Incorrect strip length',
    'WC03': 'Wire conductor damaged',
    'nan': 'Empty In Database'

}

df['Failure Codes'] = df['Failure Codes'].str.strip().str.replace('"', '').str.upper()
df1['Failure Codes'] = df1['Failure Codes'].str.strip().str.replace('"', '').str.upper()
df2['Failure Codes'] = df2['Failure Codes'].str.strip().str.replace('"', '').str.upper()
df3['Failure Codes'] = df3['Failure Codes'].str.strip().str.replace('"', '').str.upper()

missing_codes = df[~df['Failure Codes'].isin(failure_code_descriptions.keys())]['Failure Codes'].unique()
if len(missing_codes) > 0:
    print("Missing descriptions for codes:", missing_codes)  
    
missing_codes = df1[~df1['Failure Codes'].isin(failure_code_descriptions.keys())]['Failure Codes'].unique()
if len(missing_codes) > 0:
    print("Missing descriptions for codes:", missing_codes)

missing_codes = df2[~df2['Failure Codes'].isin(failure_code_descriptions.keys())]['Failure Codes'].unique()
if len(missing_codes) > 0:
    print("Missing descriptions for codes:", missing_codes) 
    
missing_codes = df3[~df3['Failure Codes'].isin(failure_code_descriptions.keys())]['Failure Codes'].unique()
if len(missing_codes) > 0:
    print("Missing descriptions for codes:", missing_codes) 

df['Description'] = df['Failure Codes'].map(failure_code_descriptions).fillna('Unknown')
df1['Description'] = df1['Failure Codes'].map(failure_code_descriptions).fillna('Unknown')
df2['Description'] = df2['Failure Codes'].map(failure_code_descriptions).fillna('Unknown')
df3['Description'] = df3['Failure Codes'].map(failure_code_descriptions).fillna('Unknown')

month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

all_reason_codes = pd.concat([df1['Area Type'], df2['Area Type'], df3['Area Type']]).dropna().unique()
all_reason_codes.sort()

all_operations = pd.concat([df1['Operation #'], df2['Operation #'], df3['Operation #']]).dropna().unique()
all_operations.sort()

unique_failure_code = df['Failure Codes'].dropna().unique()
unique_failure_code.sort()

all_customers = pd.concat([df1['Customer'], df2['Customer'], df3['Customer']]).dropna().unique()
all_customers.sort()

all_dates = pd.concat([df1['Date'], df2['Date'], df3['Date']])

unique_years = all_dates.dt.year.unique()
unique_years.sort()

app = dash.Dash(__name__)

df5['Start'] = pd.to_datetime(df5['Start'])
df5['End'] = pd.to_datetime(df5['End'])

app.layout = html.Div([
    html.H1("DMR Report Dashboard"),
    
            html.Div([
                dcc.Dropdown(
                    id='dataframe-selection-dropdown',
                    options=[
                        {'label': 'MO', 'value': 'df1'},
                        {'label': 'EVANS', 'value': 'df2'},
                        {'label': 'NMB', 'value': 'df3'},
                    ],
                    value='df1', 
                    style={'width': '50%', 'marginBottom': '10px'}
                ),
            ]),
    
    
            html.Div([
                dcc.Dropdown(
                    id='chart-dropdown',
                    options=[
                        {'label': 'Inspection Log', 'value': 'insp_log'},
                        {'label': 'DMR Report', 'value': 'dmr_report'},
                    ],
                    value='insp_log',
                    style={'width': '50%'}
                ),
                
                dcc.Dropdown(
                    id='month-dropdown',
                    options=[
                        {'label': month_names[i], 'value': i + 1} for i in range(12)
                    ],
                    value=1,  
                    style={'width': '50%'}
                ),
                
                dcc.Dropdown(
                    id='year-dropdown',
                    options=[{'label': year, 'value': year} for year in unique_years],
                    value=max(unique_years), 
                    style={'width': '50%'}
                ),
                
                dcc.Dropdown(
                    id='insp_log-dropdown',
                    options=[
                        {'label': 'Part Numbers', 'value': 'part_numbers'},
                        {'label': 'Failure Codes', 'value': 'pi_rw_code'},
                    ],
                    value='part_numbers',
                    style={'width': '50%'}
                ),
                
                dcc.Dropdown(
                    id='pi/rw-dropdown',
                    options=[
                        {'label': 'PI', 'value': 'PI'},
                        {'label': 'RW', 'value': 'RW'},
                    ],
                    value='RW',
                    style={'width': '50%'}
                ),
                
                dcc.Dropdown(
                    id='plant_dmr-dropdown',
                    options=[
                        {'label': 'Plant DMR', 'value': 'Plant DMR'},
                        {'label': 'Plant + Minor Rework', 'value': 'Plant + Minor Rework'},
                        {'label': 'Work Centers', 'value': 'Work Centers'},
                        {'label': 'Operation', 'value': 'Operation'},
                        {'label': 'Operator', 'value': 'Operator'},
                        {'label': 'Customer', 'value': 'Customer'}
                    ],
                    value='Plant DMR',
                    style={'width': '50%'}
                ),
                
                dcc.Dropdown(
                    id='plant_area-dropdown',
                    options=[{'label': reason_code, 'value': reason_code} for reason_code in all_reason_codes],
                    value=all_reason_codes[0] if len(all_reason_codes) > 0 else None,
                    style={'width': '50%'}
                ),
                
                dcc.Dropdown(
                    id='operation_type-dropdown',
                    options=[
                        {'label': 'Individual OP', 'value': 'Individual OP'},
                        {'label': 'Top 5 OP', 'value': 'Top 5 OP'},
                        {'label': 'All OP', 'value': 'All OP'}
                    ],
                    value='Individual OP',
                    style={'width': '50%'}
                ),                
                
                dcc.Dropdown(
                    id='operation-dropdown',
                    options=[{'label': operation, 'value': operation} for operation in all_operations],
                    value=all_operations[0] if len(all_operations) > 0 else None,
                    style={'width': '50%'}
                ),
                
                dcc.Dropdown(
                    id='failure_code-dropdown',
                    options=[
                        {'label': 'Individual Codes', 'value': 'Individual Codes'},
                        {'label': 'All Codes', 'value': 'All Codes'},
                    ],
                    value='All Codes',
                    style={'width': '50%'}
                ),   
                
                dcc.Dropdown(
                    id='indiv_codes-dropdown',
                    options=[{'label': failure_code, 'value': failure_code} for failure_code in unique_failure_code],
                    value=unique_failure_code[0] if len(unique_failure_code) > 0 else None,
                    style={'width': '50%'}
                ),
                
                dcc.Dropdown(
                    id='customer-dropdown',
                    options=[
                        {'label': 'All Customers', 'value': 'All Customers'},
                        {'label': 'Part #', 'value': 'Part #'},
                        {'label': 'Failure Codes', 'value': 'Failure Codes'},
                    ],
                    value='All Customers',
                    style={'width': '50%'}
                ),   
                
                dcc.Dropdown(
                    id='indiv_customers-dropdown',
                    options=[{'label': customer, 'value': customer} for customer in all_customers],
                    value=all_customers[0] if len(all_customers) > 0 else None,
                    style={'width': '50%'}
                ),
                
            html.Div(id='output-charts', style={'padding': '20px'})
    ], style={'fontFamily': 'Arial, sans-serif', 'padding': '10px'}),
], style={'padding': '20px', 'backgroundColor': '#f7f7f7'})

def get_date_range_from_excel(month, year):
    month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    month_name = month_names[month - 1] 

    date_range = df5[df5['Month'] == month_name]
    
    
    if date_range.empty:
        raise ValueError(f"No data found for month: {month_name}")

    start_date = pd.to_datetime(date_range['Start'].iloc[0], errors='coerce')
    end_date = pd.to_datetime(date_range['End'].iloc[0], errors='coerce')

    start_date = start_date.replace(year=year)
    end_date = end_date.replace(year=year)
    
    return start_date, end_date

def generate_top_5_chart(df, selected_month, selected_year, selected_pi_rw):
    df_filtered = df[(df['PI or RW'] == selected_pi_rw) & (df['Date'].dt.year == selected_year)]

    if selected_month is not None:
        df_filtered = df_filtered[df_filtered['Date'].dt.month == selected_month]
        title_suffix = f'{month_names[selected_month - 1]} {selected_year}'
    else:
        title_suffix = f'for the year {selected_year}'

    top_parts = df_filtered.groupby('Part #').agg({'QTY': 'sum'}).reset_index().sort_values(by='QTY', ascending=False).head(5)
    top_parts['Part #'] = top_parts['Part #'].astype(str)
    total_qty = df_filtered['QTY'].sum()
    total_qty_df = pd.DataFrame({'Total Quantity': [total_qty]})

    fig = px.bar(top_parts, x='Part #', y='QTY', orientation='v', text='QTY',
        labels={'Part #': 'Part #', 'QTY': 'Quantity'},
        title=f'{selected_pi_rw} Top 5 Part #s {title_suffix}')

    table = dash_table.DataTable(
        columns=[{"name": i, "id": i} for i in top_parts.columns],
        data=top_parts.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={'backgroundColor': 'white', 'fontWeight': 'bold'},
        export_format="xlsx",
        export_headers="display",
        export_columns="visible",
    )
    table_total_qty = dash_table.DataTable(
        columns=[{"name": "Total Quantity", "id": "Total Quantity"}],
        data=total_qty_df.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={'backgroundColor': 'white', 'fontWeight': 'bold'}
    )
    return fig, top_parts, table, table_total_qty

def generate_all_codes_chart(df, selected_month, selected_year, selected_pi_rw):
    
    if selected_month is None:
        filtered_data = df[(df['Date'].dt.year == selected_year) & (df['PI or RW'] == selected_pi_rw)]
        title_suffix = f' for {selected_year}'
    else:
        filtered_data = df[(df['Date'].dt.month == selected_month) & (df['Date'].dt.year == selected_year) & (df['PI or RW'] == selected_pi_rw)]
        title_suffix = f' for {month_names[selected_month - 1]} {selected_year}'

    grouped_data = filtered_data.groupby('Failure Codes')['QTY'].sum().reset_index().sort_values(by='QTY', ascending=False)

    total_qty = filtered_data['QTY'].sum()
    total_qty_df = pd.DataFrame({'Total Quantity': [total_qty]})            
            
    fig = px.bar(grouped_data, x='Failure Codes', y='QTY', orientation='v', text='QTY',
            labels={'Failure Codes': 'Failure Codes', 'QTY': 'Quantity'},
            title=f'{selected_pi_rw} Failure Codes Quantities{title_suffix}')
                        
    if 'Description' not in grouped_data.columns:
        grouped_data = grouped_data.merge(df[['Failure Codes', 'Description']].drop_duplicates(), on='Failure Codes', how='left')

    table = dash_table.DataTable(
        columns=[{"name": i, "id": i} for i in ['Failure Codes', 'QTY', 'Description']],
        data=grouped_data.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        },
        export_format="xlsx",
        export_headers="display",
        export_columns="visible",
    )

    table_total_qty = dash_table.DataTable(
        columns=[{"name": "Total Quantity", "id": "Total Quantity"}],
        data=total_qty_df.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        }
    ) 
    
    return fig, grouped_data, table, table_total_qty

def generate_indiv_codes_chart(df, selected_month, selected_year, selected_code):
    
    if selected_month is None:
        filtered_df = df[(df['Date'].dt.year == selected_year) &
                            (df['Failure Codes'] == selected_code)]
        title_suffix = f' for {selected_year}'
    else:
        filtered_df = df[(df['Date'].dt.year == selected_year) &
                           (df['Date'].dt.month == selected_month) &
                           (df['Failure Codes'] == selected_code)]
        title_suffix = f' for {month_names[selected_month - 1]} {selected_year}'

    grouped_data = filtered_df.groupby(['Part #', 'Failure Codes']).agg({'QTY': 'sum'}).reset_index().sort_values(by='QTY', ascending=False)

    grouped_data['Part #'] = grouped_data['Part #'].astype(str)
    
    total_qty = filtered_df['QTY'].sum()
    total_qty_df = pd.DataFrame({'Total Quantity': [total_qty]})            
            
    fig = px.bar(grouped_data, x='Part #', y='QTY', orientation='v', text='QTY',
            labels={'Part #': 'Part #', 'QTY': 'Quantity'},
            title=f'{selected_code} Quantities{title_suffix}')
                        
    if 'Description' not in grouped_data.columns:
        grouped_data = grouped_data.merge(df[['Failure Codes', 'Description']].drop_duplicates(), on='Failure Codes', how='left')

    table = dash_table.DataTable(

        columns=[{"name": i, "id": i} for i in ['Part #', 'QTY', 'Description']],
        data=grouped_data.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        },
        export_format="xlsx",
        export_headers="display",
        export_columns="visible",
    )

    table_total_qty = dash_table.DataTable(
        columns=[{"name": "Total Quantity", "id": "Total Quantity"}],
        data=total_qty_df.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        }
    ) 
    
    return fig, grouped_data, table, table_total_qty

def generate_plant_dmr_chart(df4, start_date, end_date, selected_year):

    if start_date is not None and end_date is not None:
        filtered_df4 = df4[(df4['Date'] >= start_date) & (df4['Date'] <= end_date)]
        title_suffix = f'from {start_date.strftime("%m/%d/%Y")} to {end_date.strftime("%m/%d/%Y")}'
    else:
        filtered_df4 = df4[df4['Date'].dt.year == selected_year]
        title_suffix = f'for the year {selected_year}'

    grouped_data = filtered_df4.groupby('Failure Codes')['QTY'].sum().reset_index().sort_values(by='QTY', ascending=False)
    grouped_data = grouped_data.merge(df4[['Failure Codes', 'Description']].drop_duplicates(), on='Failure Codes', how='left')

    fig = px.bar(grouped_data, x='Failure Codes', y='QTY', text='QTY', 
                 title=f'Plant DMR - Failure Codes Quantities {title_suffix}',
                 labels={'Failure Codes': 'Failure Codes', 'QTY': 'Quantity'})
    
    table = dash_table.DataTable(
        columns=[{"name": i, "id": i} for i in grouped_data.columns],
        data=grouped_data.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        },
        export_format="xlsx",
        export_headers="display",
        export_columns="visible",
    )

    total_qty = grouped_data['QTY'].sum()
    total_qty_df = pd.DataFrame({'Total Quantity': [total_qty]})
    table_total_qty = dash_table.DataTable(
        columns=[{"name": "Total Quantity", "id": "Total Quantity"}],
        data=total_qty_df.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        }
    )
    
    return fig, grouped_data, table, table_total_qty

def generate_plant_minor_rework_chart(df, df4, start_date, end_date, selected_year):

    filtered_df = df[(df['PI or RW'] == 'RW')]
    combined_df = pd.concat([filtered_df, df4])

    if start_date is not None and end_date is not None:
        filtered_combined_df = combined_df[(combined_df['Date'] >= start_date) & (combined_df['Date'] <= end_date)]
        title_suffix = f'from {start_date.strftime("%m/%d/%Y")} to {end_date.strftime("%m/%d/%Y")}'
    else:
        filtered_combined_df = combined_df[combined_df['Date'].dt.year == selected_year]
        title_suffix = f'for the year {selected_year}'
    
    grouped_data = filtered_combined_df.groupby('Failure Codes')['QTY'].sum().reset_index().sort_values(by='QTY', ascending=False)

    grouped_data = grouped_data.merge(df4[['Failure Codes', 'Description']].drop_duplicates(), on='Failure Codes', how='left')

    table = dash_table.DataTable(
        columns=[{"name": i, "id": i} for i in ['Failure Codes', 'QTY', 'Description']],
        data=grouped_data.to_dict('records'),
        style_table={'marginTop': 20},  
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        },
    export_format="xlsx",
    export_headers="display",
    export_columns="visible",
    )           

    fig = px.bar(grouped_data, x='Failure Codes', y='QTY', orientation='v', text='QTY',
                 labels={'Failure Codes': 'Failure Codes', 'QTY': 'Quantity'},
                 title=f'Plant + Minor Rework - Failure Codes Quantities{title_suffix}')

    total_qty = grouped_data['QTY'].sum()
    total_qty_df = pd.DataFrame({'Total Quantity': [total_qty]})  

    table_total_qty = dash_table.DataTable(
        columns=[{"name": "Total Quantity", "id": "Total Quantity"}],
        data=total_qty_df.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        }
    )     

    return fig, grouped_data, table, table_total_qty

def generate_work_centers_chart(df4, start_date, end_date, selected_year, selected_area):

    if start_date is not None and end_date is not None:
        filtered_df4 = df4[(df4['Date'] >= start_date) & (df4['Date'] <= end_date) &
                           (df4['Area Type'] == selected_area)]
        title_suffix = f' in {selected_area} from {start_date.strftime("%m/%d/%Y")} to {end_date.strftime("%m/%d/%Y")}'
    else:
        filtered_df4 = df4[(df4['Date'].dt.year == selected_year) &
                           (df4['Area Type'] == selected_area)]
        title_suffix = f' in {selected_area} for {selected_year}'

    grouped_data = filtered_df4.groupby('Failure Codes')['QTY'].sum().reset_index().sort_values(by='QTY', ascending=False)
    
    grouped_data = grouped_data.merge(df4[['Failure Codes', 'Description']].drop_duplicates(), on='Failure Codes', how='left')

    table = dash_table.DataTable(
        columns=[{"name": i, "id": i} for i in ['Failure Codes', 'QTY', 'Description']],
        data=grouped_data.to_dict('records'),
        style_table={'marginTop': 20},  
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        },
    export_format="xlsx",
    export_headers="display",
    export_columns="visible",
    )

    fig = px.bar(grouped_data, x='Failure Codes', y='QTY', text='QTY',
                 labels={'Failure Codes': 'Failure Codes', 'QTY': 'Quantity'},
                 title=f'Total Quantities of Failure Codes {title_suffix}')

    total_qty = grouped_data['QTY'].sum()
    total_qty_df = pd.DataFrame({'Total Quantity': [total_qty]})  

    table_total_qty = dash_table.DataTable(
        columns=[{"name": "Total Quantity", "id": "Total Quantity"}],
        data=total_qty_df.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        }
    )        
    
    return fig, grouped_data, table, table_total_qty

def generate_operation_chart(df4, start_date, end_date, selected_year, selected_operation):

    operation_filter = df4['Operation #'] == selected_operation
    if start_date is not None and end_date is not None:
        filtered_df4 = df4[(df4['Date'] >= start_date) & (df4['Date'] <= end_date)]
        title_suffix = f'from {start_date.strftime("%m/%d/%Y")} to {end_date.strftime("%m/%d/%Y")}'
    else:
        filtered_df4 = df4[df4['Date'].dt.year == selected_year]
        title_suffix = f'for the year {selected_year}'

    filtered_df4 = df4[operation_filter]

    grouped_data = filtered_df4.groupby(['Part #', 'Operation #']).agg({'QTY': 'sum'}).reset_index().sort_values(by='QTY', ascending=False)

    fig = px.bar(grouped_data, x='Part #', y='QTY', text='QTY', color='Operation #',
                 labels={'Part #': 'Part #', 'QTY': 'Total Quantity', 'Operation #': 'Operation #'},
                 title=f'Total Quantities for Each Part # in Operation {selected_operation}')
    
    table = dash_table.DataTable(
        columns=[{"name": i, "id": i} for i in grouped_data.columns],
        data=grouped_data.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        },
        export_format="xlsx",
        export_headers="display",
        export_columns="visible",
    )
    
    total_qty = grouped_data['QTY'].sum()
    total_qty_df = pd.DataFrame({'Total Quantity': [total_qty]})  
    
    table_total_qty = dash_table.DataTable(
        columns=[{"name": "Total Quantity", "id": "Total Quantity"}],
        data=total_qty_df.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        }
    )   
    
    return fig, grouped_data, table, table_total_qty

def generate_operation_type_chart(df4, start_date, end_date, selected_year):
    
    if start_date is not None and end_date is not None:
        filtered_df4 = df4[(df4['Date'] >= start_date) & (df4['Date'] <= end_date)]
        title_suffix = f'from {start_date.strftime("%m/%d/%Y")} to {end_date.strftime("%m/%d/%Y")}'
    else:
        filtered_df4 = df4[df4['Date'].dt.year == selected_year]
        title_suffix = f'for the year {selected_year}'

    grouped_data = filtered_df4.groupby(['Operation #']).agg({'QTY': 'sum'}).reset_index().sort_values(by='QTY', ascending=False).head(5)
    grouped_data['Operation #'] = grouped_data['Operation #'].astype(str)   

    fig = px.bar(grouped_data, x='Operation #', y='QTY', text='QTY',
                 labels={'Operation #': 'Operation #', 'QTY': 'Total Quantity'},
                 title=f'Top 5 Quantities for Operations {title_suffix}')

    table = dash_table.DataTable(
        columns=[{"name": i, "id": i} for i in grouped_data.columns],
        data=grouped_data.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={'backgroundColor': 'white', 'fontWeight': 'bold'},
        export_format="xlsx",
        export_headers="display",
        export_columns="visible",
    )

    total_qty = grouped_data['QTY'].sum()
    total_qty_df = pd.DataFrame({'Total Quantity': [total_qty]})  

    table_total_qty = dash_table.DataTable(
        columns=[{"name": "Total Quantity", "id": "Total Quantity"}],
        data=total_qty_df.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={'backgroundColor': 'white', 'fontWeight': 'bold'}
    )                      

    return fig, grouped_data, table, table_total_qty

def generate_all_op_chart(df4, start_date, end_date, selected_year):
        
    if start_date is not None and end_date is not None:
        filtered_df4 = df4[(df4['Date'] >= start_date) & (df4['Date'] <= end_date)]
        title_suffix = f'from {start_date.strftime("%m/%d/%Y")} to {end_date.strftime("%m/%d/%Y")}'
    else:
        filtered_df4 = df4[df4['Date'].dt.year == selected_year]
        title_suffix = f'for the year {selected_year}'

    grouped_data = filtered_df4.groupby(['Operation #']).agg({'QTY': 'sum'}).reset_index().sort_values(by='QTY', ascending=False)

    grouped_data['Operation #'] = grouped_data['Operation #'].astype(str)   

    fig = px.bar(grouped_data, x='Operation #', y='QTY', text='QTY',
        labels={'Operation #': 'Operation #', 'QTY': 'Total Quantity'},
        title=f'Top 5 Quantities for Operations in {title_suffix}')

    table = dash_table.DataTable(
        columns=[{"name": i, "id": i} for i in grouped_data.columns],
        data=grouped_data.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        },
        export_format="xlsx",
        export_headers="display",
        export_columns="visible",
    )

    total_qty = grouped_data['QTY'].sum()
    total_qty_df = pd.DataFrame({'Total Quantity': [total_qty]})  

    table_total_qty = dash_table.DataTable(
        columns=[{"name": "Total Quantity", "id": "Total Quantity"}],
        data=total_qty_df.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        }
    )                      


    return fig, grouped_data, table, table_total_qty

def generate_operator_chart(df4, start_date, end_date, selected_year):
    
    if start_date is not None and end_date is not None:
        filtered_df4 = df4[(df4['Date'] >= start_date) & (df4['Date'] <= end_date)]
        title_suffix = f'from {start_date.strftime("%m/%d/%Y")} to {end_date.strftime("%m/%d/%Y")}'
    else:
        filtered_df4 = df4[df4['Date'].dt.year == selected_year]
        title_suffix = f'for the year {selected_year}'

    grouped_data = filtered_df4.groupby('NAME')['QTY'].sum().reset_index().sort_values(by='QTY', ascending=False)

    table = dash_table.DataTable(
        columns=[{"name": i, "id": i} for i in ['NAME', 'QTY']],
        data=grouped_data.to_dict('records'),
        style_table={'marginTop': 20},  
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        },
    export_format="xlsx",
    export_headers="display",
    export_columns="visible",
    )       

    fig = px.bar(grouped_data, x='NAME', y='QTY', orientation='v', text='QTY',
                 labels={'NAME': 'NAME', 'QTY': 'Quantity'},
                 title=f'Operator DMR Quantities{title_suffix}')

    total_qty = grouped_data['QTY'].sum()
    total_qty_df = pd.DataFrame({'Total Quantity': [total_qty]})  

    table_total_qty = dash_table.DataTable(
        columns=[{"name": "Total Quantity", "id": "Total Quantity"}],
        data=total_qty_df.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        }
    ) 
    
    return fig, grouped_data, table, table_total_qty

def generate_all_customer_chart(df4, start_date, end_date, selected_year):
        
    if start_date is not None and end_date is not None:
        df4_filtered = df4[(df4['Date'] >= start_date) & (df4['Date'] <= end_date)]
        title_suffix = f'from {start_date.strftime("%m/%d/%Y")} to {end_date.strftime("%m/%d/%Y")}'
    else:
        df4_filtered = df4[df4['Date'].dt.year == selected_year]
        title_suffix = f'for the year {selected_year}'

    grouped_data = df4_filtered.groupby('Customer').agg({'QTY': 'sum'}).reset_index().sort_values(by='QTY', ascending=False)
            
    grouped_data['Customer'] = grouped_data['Customer'].astype(str)

    total_qty = df4_filtered['QTY'].sum()
    total_qty_df = pd.DataFrame({'Total Quantity': [total_qty]})

    fig = px.bar(grouped_data, x='Customer', y='QTY', orientation='v', text='QTY',
        labels={'Customer': 'Customer', 'QTY': 'Quantity'},
        title=f'Customer Data For {title_suffix}')

    table = dash_table.DataTable(
        columns=[{"name": i, "id": i} for i in grouped_data.columns],
        data=grouped_data.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        },
        export_format="xlsx",
        export_headers="display",
        export_columns="visible",
    )
            
    table_total_qty = dash_table.DataTable(
        columns=[{"name": "Total Quantity", "id": "Total Quantity"}],
        data=total_qty_df.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        }
    )
    
    return fig, grouped_data, table, table_total_qty        

def generate_customer_part_chart(df4, start_date, end_date, selected_year, selected_indiv_customer):

    customer_filter = df4['Customer'] == selected_indiv_customer
    if start_date is not None and end_date is not None:
        filtered_df4 = df4[(df4['Date'] >= start_date) & (df4['Date'] <= end_date)]
        title_suffix = f'from {start_date.strftime("%m/%d/%Y")} to {end_date.strftime("%m/%d/%Y")}'
    else:
        filtered_df4 = df4[df4['Date'].dt.year == selected_year]
        title_suffix = f'for the year {selected_year}'

    filtered_df4 = df4[customer_filter]

    grouped_data = filtered_df4.groupby(['Part #', 'Customer']).agg({'QTY': 'sum'}).reset_index().sort_values(by='QTY', ascending=False)

    fig = px.bar(grouped_data, x='Part #', y='QTY', text='QTY', color='Customer',
                 labels={'Part #': 'Part #', 'QTY': 'Total Quantity', 'Customer': 'Customer'},
                 title=f'Total Quantities for Each Part # in Customer {selected_indiv_customer}')
    
    table = dash_table.DataTable(
        columns=[{"name": i, "id": i} for i in grouped_data.columns],
        data=grouped_data.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        },
        export_format="xlsx",
        export_headers="display",
        export_columns="visible",
    )
    
    total_qty = grouped_data['QTY'].sum()
    total_qty_df = pd.DataFrame({'Total Quantity': [total_qty]})  
    
    table_total_qty = dash_table.DataTable(
        columns=[{"name": "Total Quantity", "id": "Total Quantity"}],
        data=total_qty_df.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        }
    )   
    
    return fig, grouped_data, table, table_total_qty

def generate_customer_part_chart(df4, selected_month, selected_year, selected_indiv_customer):

    customer_filter = df4['Customer'] == selected_indiv_customer
    if selected_month is not None:
        customer_filter &= df4['Date'].dt.month == selected_month
    if selected_year is not None:
        customer_filter &= df4['Date'].dt.year == selected_year

    filtered_df4 = df4[customer_filter]

    grouped_data = filtered_df4.groupby(['Part #', 'Customer']).agg({'QTY': 'sum'}).reset_index().sort_values(by='QTY', ascending=False)

    fig = px.bar(grouped_data, x='Part #', y='QTY', text='QTY', color='Customer',
                 labels={'Part #': 'Part #', 'QTY': 'Total Quantity', 'Customer': 'Customer'},
                 title=f'Total Quantities for Each Part # in Customer {selected_indiv_customer}')
    
    table = dash_table.DataTable(
        columns=[{"name": i, "id": i} for i in grouped_data.columns],
        data=grouped_data.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        },
        export_format="xlsx",
        export_headers="display",
        export_columns="visible",
    )
    
    total_qty = grouped_data['QTY'].sum()
    total_qty_df = pd.DataFrame({'Total Quantity': [total_qty]})  
    
    table_total_qty = dash_table.DataTable(
        columns=[{"name": "Total Quantity", "id": "Total Quantity"}],
        data=total_qty_df.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        }
    )   
    
    return fig, grouped_data, table, table_total_qty

def generate_customer_failcode_chart(df4, start_date, end_date, selected_year, selected_indiv_customer):

    customer_filter = (df4['Customer'] == selected_indiv_customer)

    if start_date is not None and end_date is not None:
        date_filter = (df4['Date'] >= start_date) & (df4['Date'] <= end_date)
        title_suffix = f'from {start_date.strftime("%m/%d/%Y")} to {end_date.strftime("%m/%d/%Y")}'
    else:
        date_filter = df4['Date'].dt.year == selected_year
        title_suffix = f'for the year {selected_year}'

    filtered_df4 = df4[customer_filter & date_filter]

    grouped_data = filtered_df4.groupby(['Failure Codes', 'Customer']).agg({'QTY': 'sum'}).reset_index().sort_values(by='QTY', ascending=False)

    fig = px.bar(grouped_data, x='Failure Codes', y='QTY', text='QTY', color='Customer',
                 labels={'Failure Codes': 'Failure Codes', 'QTY': 'Total Quantity', 'Customer': 'Customer'},
                 title=f'Total Quantities for Each Failure Code in Customer {selected_indiv_customer}')
    
    table = dash_table.DataTable(
        columns=[{"name": i, "id": i} for i in grouped_data.columns],
        data=grouped_data.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        },
        export_format="xlsx",
        export_headers="display",
        export_columns="visible",
    )
    
    total_qty = grouped_data['QTY'].sum()
    total_qty_df = pd.DataFrame({'Total Quantity': [total_qty]})  
    
    table_total_qty = dash_table.DataTable(
        columns=[{"name": "Total Quantity", "id": "Total Quantity"}],
        data=total_qty_df.to_dict('records'),
        style_table={'marginTop': 20},
        style_cell={'textAlign': 'center'},
        style_header={
            'backgroundColor': 'white',
            'fontWeight': 'bold'
        }
    )   
    
    return fig, grouped_data, table, table_total_qty

@app.callback(
    [Output('output-charts', 'children'),
     Output('insp_log-dropdown', 'style'), 
     Output('pi/rw-dropdown', 'style'),
     Output('plant_dmr-dropdown', 'style'),
     Output('plant_area-dropdown', 'style'),
     Output('operation-dropdown', 'style'),
     Output('operation_type-dropdown', 'style'),
     Output('failure_code-dropdown', 'style'),
     Output('indiv_codes-dropdown', 'style'),
     Output('customer-dropdown', 'style'),
     Output('indiv_customers-dropdown', 'style')],
    [Input('dataframe-selection-dropdown', 'value'),
     Input('chart-dropdown', 'value'),
     Input('month-dropdown', 'value'),
     Input('year-dropdown', 'value'),
     Input('insp_log-dropdown', 'value'),
     Input('pi/rw-dropdown', 'value'),
     Input('plant_dmr-dropdown', 'value'),
     Input('plant_area-dropdown', 'value'),
     Input('operation-dropdown', 'value'),
     Input('operation_type-dropdown', 'value'),
     Input('failure_code-dropdown', 'value'),
     Input('indiv_codes-dropdown', 'value'),
     Input('customer-dropdown', 'value'),
     Input('indiv_customers-dropdown', 'value')]
)

def update_charts(selected_df, selected_chart, selected_month, selected_year, selected_insp_log, selected_pi_rw, selected_plant_dmr, selected_area, selected_operation, selected_operation_type, selected_failure_code, selected_code, selected_customer, selected_indiv_customer):
    
    pi_rw_dropdown_style = {'width': '50%', 'display': 'none'}  
    plant_dmr_dropdown_style = {'width': '50%', 'display': 'none'}  
    plant_area_dropdown_style = {'width': '50%', 'display': 'none'}
    operation_dropdown_style = {'width': '50%', 'display': 'none'}
    operation_type_dropdown_style = {'width': '50%', 'display': 'none'}
    failure_code_dropdown_style = {'width': '50%', 'display': 'none'}
    indiv_codes_dropdown_style = {'width': '50%', 'display': 'none'}
    insp_log_dropdown_style = {'width': '50%', 'display': 'none'}
    customer_dropdown_style = {'width': '50%', 'display': 'none'}
    indiv_customers_dropdown_style = {'width': '50%', 'display': 'none'}
    
    month_name = month_names[selected_month - 1] if selected_month else "Yearly"
    
    if not selected_month:  
        start_date = pd.to_datetime(f"{selected_year}-01-01")
        end_date = pd.to_datetime(f"{selected_year}-12-31")
    else:
        start_date, end_date = get_date_range_from_excel(selected_month, selected_year)
    
    df4 = df1 if selected_df == 'df1' else df2 if selected_df == 'df2' else df3

    df.reset_index(drop=True, inplace=True)
    df1.reset_index(drop=True, inplace=True)
    
    if selected_chart == 'insp_log':
        
        insp_log_dropdown_style = {'width': '50%'}
                                       
        if selected_insp_log == 'part_numbers':
            
            pi_rw_dropdown_style = {'width': '50%'}
            
            fig, top_parts, table, table_total_qty = generate_top_5_chart(df, selected_month, selected_year, selected_pi_rw)
            
            fig.update_layout(autosize=False, width=2400, height=600)
            
        elif selected_insp_log == 'pi_rw_code':
            
            failure_code_dropdown_style = {'width': '50%'}
            
            if selected_failure_code == 'All Codes':
                
                pi_rw_dropdown_style = {'width': '50%'}
                
                fig, grouped_data, table, table_total_qty = generate_all_codes_chart(df, selected_month, selected_year, selected_pi_rw)
                
                fig.update_layout(autosize=False, width=2400, height=600)
                
            elif selected_failure_code == 'Individual Codes':
                
                indiv_codes_dropdown_style = {'width': '50%'}
                
                fig, grouped_data, table, table_total_qty = generate_indiv_codes_chart(df, selected_month, selected_year, selected_code)
                
                fig.update_layout(autosize=False, width=2400, height=600)
                
    elif selected_chart == 'dmr_report':
        
        plant_dmr_dropdown_style = {'width': '50%'}
        
        if selected_plant_dmr == 'Plant DMR':
            
            fig, grouped_data, table, table_total_qty = generate_plant_dmr_chart(df4, start_date, end_date, selected_year)
            
            fig.update_layout(autosize=False, width=2400, height=600)
            
        elif selected_plant_dmr == 'Plant + Minor Rework':
            
            fig, grouped_data, table, table_total_qty = generate_plant_minor_rework_chart(df, df4, start_date, end_date, selected_year)
            
            fig.update_layout(autosize=False, width=2400, height=600)
            
        elif selected_plant_dmr == 'Work Centers':
            
            fig, grouped_data, table, table_total_qty = generate_work_centers_chart(df4, start_date, end_date, selected_year, selected_area)
            
            plant_area_dropdown_style = {'width': '50%'}
            
            fig.update_layout(autosize=False, width=2400, height=600)
            
        elif selected_plant_dmr == 'Operation':
            
            operation_type_dropdown_style = {'width': '50%'}
            
            if selected_operation_type == 'Individual OP':
                
                fig, grouped_data, table, table_total_qty = generate_operation_chart(df4, start_date, end_date, selected_year, selected_operation)
                
                operation_dropdown_style = {'width': '50%'}
                
                fig.update_layout(autosize=False, width=2400, height=600)
                
            elif selected_operation_type == 'Top 5 OP':
                
                fig, grouped_data, table, table_total_qty = generate_operation_type_chart(df4, start_date, end_date, selected_year)
                
                fig.update_layout(autosize=False, width=2400, height=600)
                
            elif selected_operation_type == 'All OP':
                
                fig, grouped_data, table, table_total_qty = generate_all_op_chart(df4, start_date, end_date, selected_year)
                
                fig.update_layout(autosize=False, width=2400, height=600)
                
        elif selected_plant_dmr == 'Operator':
            
            fig, grouped_data, table, table_total_qty = generate_operator_chart(df4, start_date, end_date, selected_year)
            
            fig.update_layout(autosize=False, width=2400, height=600)
            
        elif selected_plant_dmr == 'Customer':
            
            customer_dropdown_style = {'width': '50%'}
            
            if selected_customer == 'All Customers':
                
                df4.dropna(subset=['Customer'], inplace=True)
                
                fig, grouped_data, table, table_total_qty = generate_all_customer_chart(df4, start_date, end_date, selected_year)
                
                fig.update_layout(autosize=False, width=2400, height=600)
                
            elif selected_customer == 'Part #':
                
                indiv_customers_dropdown_style = {'width': '50%'}
                    
                df4.dropna(subset=['Customer'], inplace=True)
                    
                fig, grouped_data, table, table_total_qty = generate_customer_part_chart(df4, selected_month, selected_year, selected_indiv_customer)
                    
                fig.update_layout(autosize=False, width=2400, height=600)    
                
            elif selected_customer == 'Failure Codes':
                
                df4.dropna(subset=['Customer'], inplace=True)

                indiv_customers_dropdown_style = {'width': '50%'}
                
                fig, grouped_data, table, table_total_qty = generate_customer_failcode_chart(df4, start_date, end_date, selected_year, selected_indiv_customer)
                
                fig.update_layout(autosize=False, width=2400, height=600)
                                       
    return [dcc.Graph(figure=fig), table, table_total_qty], insp_log_dropdown_style, pi_rw_dropdown_style, plant_dmr_dropdown_style, plant_area_dropdown_style, operation_dropdown_style, operation_type_dropdown_style, failure_code_dropdown_style, indiv_codes_dropdown_style, customer_dropdown_style, indiv_customers_dropdown_style       

if __name__ == '__main__':
    app.run_server(debug=True)


