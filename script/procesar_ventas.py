import pandas as pd
from openpyxl.styles import Alignment, PatternFill, Font


file_path = './data/datos_ventas.xlsx'
df = pd.read_excel(file_path)

df['Total_Venta'] = df.apply(
    lambda row: row['Cantidad'] * row['Precio_Unitario'] if pd.isna(row['Total_Venta']) else row['Total_Venta'],
    axis=1
)

df['Fecha'] = pd.to_datetime(df['Fecha'])

df_2023 = df[df['Fecha'].dt.year == 2023]
df_2023['Mes'] = df_2023['Fecha'].dt.month

ventas_por_vendedor = df_2023.groupby('Vendedor')['Total_Venta'].sum().reset_index()
ventas_mensuales = df_2023.groupby('Mes')['Total_Venta'].sum().reset_index()

ventas_por_vendedor['Total_Venta'] = ventas_por_vendedor['Total_Venta'].apply(lambda x: f"${x:,.0f}")
ventas_mensuales['Total_Venta'] = ventas_mensuales['Total_Venta'].apply(lambda x: f"${x:,.0f}")


def apply_style(sheet):
    header_fill = PatternFill(start_color="0053A1", end_color="0053A1", fill_type="solid")  # Bancolombia Azul
    header_font = Font(color="FFFFFF", bold=True)
    align_center = Alignment(horizontal="center", vertical="center")

    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center

    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except Exception:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column].width = adjusted_width


with pd.ExcelWriter('./data/resumen_ventas.xlsx', engine='openpyxl') as writer:
    ventas_por_vendedor.to_excel(writer, sheet_name='Resumen_Ventas', index=False)
    ventas_mensuales.to_excel(writer, sheet_name='Ventas_Mensuales', index=False)

    workbook = writer.book
    sheet_ventas_por_vendedor = workbook['Resumen_Ventas']
    sheet_ventas_mensuales = workbook['Ventas_Mensuales']

    # Algunos estilos a gusto
    apply_style(sheet_ventas_por_vendedor)
    apply_style(sheet_ventas_mensuales)

print("Archivo generado: resumen_ventas.xlsx")
