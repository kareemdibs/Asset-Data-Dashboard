import pandas as pd
from dash import Dash, html, dash_table, dcc, Output, Input
import plotly.express as px
from datetime import datetime
import plotly.graph_objects as go
from openpyxl import Workbook

df = pd.read_excel("Aug_Asset_Data_V1.xlsx")

selected_columns = ['operating_date', 'time_hr', 'asset_typ', 'asset_nm', 'da_schd', 'da_lmp_en', 'rt_bll_mtr', 'rt_lmp_en']
df1 = df[selected_columns]

df1 = df1.dropna(how='any')
#df1 = df1[df1 != 0].dropna()

df1.reset_index(drop=True, inplace=True)
df1.insert(0, 'unique_id', range(1, len(df1) + 1))

empty_df = pd.DataFrame()

app = Dash(__name__)

df1['time_hr'] = pd.to_datetime(df1['time_hr']).dt.strftime('%Y-%m-%d %H:%M:%S')
df1['da_schd'] = df1['da_schd'].abs()
df1['rt_bll_mtr'] = df1['rt_bll_mtr'].abs()

asset_type_options = [{'label': asset_type, 'value': asset_type} for asset_type in sorted(df1['asset_typ'].unique())]
asset_name_options = [{'label': asset_name, 'value': asset_name} for asset_name in sorted(df1['asset_nm'].unique())]
date_options = [{'label': pd.to_datetime(date).strftime('%Y-%m-%d'), 'value': date} for date in sorted(pd.to_datetime(df1['operating_date']).unique())]

app.layout = html.Div([
    html.Div(children='PCI August Asset Data Dashboard'),
    html.Hr(),
    html.Div([
        html.Div("Asset Type", style={'width': '33%', 'display': 'inline-block', 'margin-right': '2%'}),
        html.Div("Asset Name", style={'width': '28%', 'display': 'inline-block', 'margin-right': '2%'}),
        html.Div("Date", style={'width': '46%', 'display': 'inline-block'}),
    ], style={'display': 'flex', 'margin-bottom': '5px'}),
    html.Div([
        dcc.Dropdown(style={'display': 'none'}),
        dcc.Dropdown(options=asset_type_options, id='type-dropdown', placeholder="Select an asset type.", style={'width': '58%'}),
        dcc.Dropdown(options=asset_name_options, id='asset-dropdown', placeholder="Select an asset.", style={'width': '50%'}),
        dcc.Dropdown(options=date_options, id='date-dropdown', placeholder="Select a date.", style={'width': '50%'}),
        html.Button('Update', id='btn-generate', style={'width': '25%'}),
    ], style={'display': 'flex', 'margin-bottom': '20px'}),
    dash_table.DataTable(
        id='pci_table',
        columns=[{'name': col, 'id': col} for col in df1.columns if col != 'operating_date' and col != 'unique_id'],
        data=df1.to_dict('records'),
        page_size=27,
        style_table={'margin-bottom': '20px'}
        #row_selectable='multi'  # Allow multiple rows to be selected
    ),
    dcc.Graph(id='pci_graph'),
    dcc.Graph(id='pci_graph2'),
    html.Footer(children='2023 Kareem Dibs')
])

@app.callback(
    Output('asset-dropdown', 'options'),
    [Input('type-dropdown', 'value')]
)
def update_asset_names(selected_asset_type):
    if selected_asset_type:
        filtered_asset_names = df1[df1['asset_typ'] == selected_asset_type]['asset_nm'].unique()
        asset_name_options = [{'label': asset_name, 'value': asset_name} for asset_name in sorted(filtered_asset_names)]
        return asset_name_options
    else:
        # If no asset type is selected, show all asset names
        asset_name_options = [{'label': asset_name, 'value': asset_name} for asset_name in sorted(df1['asset_nm'].unique())]
        return asset_name_options

@app.callback(
    [Output('pci_table', 'data'),
     Output('pci_graph', 'figure'),
     Output('pci_graph2', 'figure')
    ],
    [Input('type-dropdown', 'value'),
     Input('asset-dropdown', 'value'),
     Input('date-dropdown', 'value'),
     Input('btn-generate', 'n_clicks')
    ]
)
def update_table(selected_asset_type, selected_asset_name, selected_date, n_clicks):
    if n_clicks is not None:
        if (selected_asset_type and selected_asset_name and selected_date):
            if (selected_asset_type):
                df_filtered = df1[df1['asset_typ'] == selected_asset_type]
                if (selected_asset_name):
                    df_filtered = df_filtered[df_filtered['asset_nm'] == selected_asset_name]
                if (selected_date):
                    df_filtered = df_filtered[df_filtered['operating_date'] == selected_date]

                total1 = round(df_filtered['da_schd'].sum(), 3)
                total2 = round(df_filtered['rt_bll_mtr'].sum(), 3)
                avg1 = round(df_filtered['da_lmp_en'].mean(), 3)
                avg2 = round(df_filtered['rt_lmp_en'].mean(), 3)

                total_row = {
                    'time_hr': 'Total',
                    'asset_typ': selected_asset_type,
                    'asset_nm': selected_asset_name,
                    'da_schd': total1,
                    'da_lmp_en': avg1,
                    'rt_bll_mtr': total2,
                    'rt_lmp_en': avg2,
                    'unique_id': 'Total'
                }

                df_on_peak = df_filtered.iloc[7:23]

                avg_on_peak1 = round(df_on_peak['da_lmp_en'].mean(skipna=True), 3)
                avg_on_peak2 = round(df_on_peak['rt_lmp_en'].mean(skipna=True), 3)

                df_off_peak = df_filtered.iloc[[0, 1, 2, 3, 4, 5, 6, 23]]

                avg_off_peak1 = round(df_off_peak['da_lmp_en'].mean(skipna=True), 3)
                avg_off_peak2 = round(df_off_peak['rt_lmp_en'].mean(skipna=True), 3)

                price_on_peak = {
                    'time_hr': 'Price (ON PEAK)',
                    'asset_typ': selected_asset_type,
                    'asset_nm': selected_asset_name,
                    'da_lmp_en': avg_on_peak1,
                    'rt_lmp_en': avg_on_peak2
                }

                price_off_peak = {
                    'time_hr': 'Price (OFF PEAK)',
                    'asset_typ': selected_asset_type,
                    'asset_nm': selected_asset_name,
                    'da_lmp_en': avg_off_peak1,
                    'rt_lmp_en': avg_off_peak2
                }

                df_filtered.loc[len(df_filtered)] = total_row
                df_filtered.loc[len(df_filtered)] = price_on_peak
                df_filtered.loc[len(df_filtered)] = price_off_peak

                filtered_df = df_filtered[~df_filtered['time_hr'].isin(['Total', 'Price (ON PEAK)', 'Price (OFF PEAK)'])]

                scatter_chart_data1 = filtered_df[['time_hr', 'da_schd', 'rt_bll_mtr']]
                scatter_chart_data2 = filtered_df[['time_hr', 'da_lmp_en', 'rt_lmp_en']]

                scatter_chart_data1['time_hr'] = (pd.to_datetime(scatter_chart_data1['time_hr']).dt.strftime('%H').astype(int) + 1) % 25
                scatter_chart_data2['time_hr'] = (pd.to_datetime(scatter_chart_data2['time_hr']).dt.strftime('%H').astype(int) + 1) % 25

                date = pd.to_datetime(selected_date).strftime('%Y-%m-%d')

                fig = go.Figure()

                fig.add_trace(go.Scatter(x=scatter_chart_data1['time_hr'], y=scatter_chart_data1['da_schd'], mode='lines+markers', name='da_schd'))
                fig.add_trace(go.Scatter(x=scatter_chart_data1['time_hr'], y=scatter_chart_data1['rt_bll_mtr'], mode='lines+markers', name='rt_bll_mtr'))

                fig.update_layout(
                    title=f'Hourly Data for {selected_asset_name} on {date} (da_schd vs rt_bll_mtr)',
                    xaxis=dict(title='Hour of Day', tickvals=scatter_chart_data1['time_hr']),
                    yaxis=dict(title='Value'),
                )

                fig1 = go.Figure()

                fig1.add_trace(go.Scatter(x=scatter_chart_data2['time_hr'], y=scatter_chart_data2['da_lmp_en'], mode='lines+markers', name='da_lmp_en'))
                fig1.add_trace(go.Scatter(x=scatter_chart_data2['time_hr'], y=scatter_chart_data2['rt_lmp_en'], mode='lines+markers', name='rt_lmp_en'))

                fig1.update_layout(
                    title=f'Hourly Data for {selected_asset_name} on {date} (da_lmp_en vs rt_lmp_en)',
                    xaxis=dict(title='Hour of Day', tickvals=scatter_chart_data2['time_hr']),
                    yaxis=dict(title='Value'),
                )

                return df_filtered.to_dict('records'), fig, fig1
    
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(x=[], y=[], mode='lines+markers'))
    fig2.update_layout(
        title="Hourly Data for Asset (da_schd vs rt_bll_mtr)",
        xaxis=dict(title='Hour of Day', tickvals=[]),
        yaxis=dict(title='Value')
    )

    fig3 = go.Figure()
    fig3.add_trace(go.Scatter(x=[], y=[], mode='lines+markers'))
    fig3.update_layout(
        title="Hourly Data for Asset (da_lmp_en vs rt_lmp_en)",
        xaxis=dict(title='Hour of Day', tickvals=[]),
        yaxis=dict(title='Value')
    )

    return empty_df.to_dict('records'), fig2, fig3

if __name__ == '__main__':
    app.run(debug=True)