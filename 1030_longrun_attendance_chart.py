import plotly.graph_objects as go
import plotly.io as pio
import os
import platform
import webbrowser

# Set the theme for better styling
pio.templates.default = "plotly_white"

# Attendance data
years = [2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029]

historical_attendance = [79.46, 88.3, 83.44, 75.38, 79.96, 76.45, 72.82, 73.36, 73.36, 66.82, 68.53, 76.0, 
                        None, None, None, None]

projected_attendance = [None, None, None, None, None, None, None, None, None, None, None, None,
                       82.92, 91.21, 100.33, 110.37]

# Create the figure
fig = go.Figure()

# Add historical data bars
fig.add_trace(go.Bar(
    x=[str(year) for year in years if historical_attendance[years.index(year)] is not None],
    y=[att for att in historical_attendance if att is not None],
    name='Historical Data (2014-2025)',
    marker=dict(
        color='#3498db',
        line=dict(color='#2980b9', width=1)
    ),
    hovertemplate='<b>%{x}</b><br>Historical: %{y:.1f} average attendance<extra></extra>'
))

# Add projected data bars
fig.add_trace(go.Bar(
    x=[str(year) for year in years if projected_attendance[years.index(year)] is not None],
    y=[att for att in projected_attendance if att is not None],
    name='Projected Growth (10% annually)',
    marker=dict(
        color='#e74c3c',
        line=dict(color='#c0392b', width=1)
    ),
    hovertemplate='<b>%{x}</b><br>Projected: %{y:.1f} average attendance<extra></extra>'
))

# Update layout
fig.update_layout(
    title={
        'text': '<b>10:30AM Congregation Attendance</b><br><span style="font-size:24px;">Historical Data (2014-2025) and Projected Growth (2026-2029)</span>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 32, 'color': '#2c3e50'}
    },
    xaxis=dict(
        title=dict(
            text='<b>Year</b>',
            font=dict(size=20, color='#2c3e50')
        ),
        tickfont=dict(size=16, color='#2c3e50'),
        gridcolor='#ecf0f1',
        showgrid=False
    ),
    yaxis=dict(
        title=dict(
            text='<b>Average Weekly Attendance</b>',
            font=dict(size=20, color='#2c3e50')
        ),
        tickfont=dict(size=16, color='#2c3e50'),
        gridcolor='#ecf0f1',
        showgrid=True,
        range=[0, 120]
    ),
    plot_bgcolor='white',
    paper_bgcolor='#f8f9fa',
    font=dict(family='Arial, sans-serif'),
    # Legend moved to bottom
    legend=dict(
        orientation="h",
        yanchor="top",
        y=-0.12,
        xanchor="center",
        x=0.5,
        bgcolor='rgba(255,255,255,0.9)',
        bordercolor='#bdc3c7',
        borderwidth=1,
        font=dict(size=20, color='#2c3e50')
    ),
    width=1200,
    height=800,
    margin=dict(l=80, r=80, t=140, b=220)
)

# # Add summary statistics annotation (moved higher up)
# summary_stats = [
#     "📈 <b>Summary Statistics</b>",
#     "",
#     "🏆 Peak Attendance (2015): <b>88.3</b>",
#     "📉 Lowest Attendance (2023): <b>66.8</b>",
#     "📊 2024 Average: <b>68.5</b>",
#     "✅ 2025 Average: <b>76.0</b>",
#     "🎯 Strategic Plan 2025 Goal: <b>75.4</b>",
#     "🚀 Projected 2026 (10% growth): <b>82.9</b>",
#     "🚀 Projected 2029 (10% growth): <b>110.4</b>"
# ]

# fig.add_annotation(
#     x=0.02,
#     y=1.02,
#     xref='paper',
#     yref='paper',
#     text="<br>".join(summary_stats),
#     showarrow=False,
#     bgcolor='rgba(236, 240, 241, 0.9)',
#     bordercolor='#bdc3c7',
#     borderwidth=1,
#     font=dict(size=11, color='#2c3e50'),
#     align='left',
#     xanchor='left',
#     yanchor='top'
# )

# Create outputs directory if it doesn't exist
output_dir = "outputs"
os.makedirs(output_dir, exist_ok=True)

# Define output filenames
html_filename = os.path.join(output_dir, "1030_longrun_attendance.html")
png_filename = os.path.join(output_dir, "1030_longrun_attendance.png")

# Save as HTML for interactive version
fig.write_html(html_filename)
print(f"✅ Interactive HTML saved: {html_filename}")

# Save as PNG with high resolution
fig.write_image(png_filename, width=1200, height=800, scale=2)
print(f"✅ PNG chart saved: {png_filename}")

# Open the HTML file automatically
try:
    html_path = os.path.abspath(html_filename)
    webbrowser.open('file://' + html_path)
    print(f"🌐 Opening HTML file in browser: {html_filename}")
except Exception as e:
    print(f"⚠️  Could not auto-open HTML: {e}")

print(f"📊 Historical data: 2014-2025 ({len([x for x in historical_attendance if x is not None])} years)")
print(f"📈 Projected data: 2026-2029 ({len([x for x in projected_attendance if x is not None])} years)")
print("🎯 2025 Goal achieved: 76.0/75.4 (100.8% of 2025 goal achieved)")