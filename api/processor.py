import os
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np
import xlsxwriter
import uuid

# Configurar matplotlib para entorno sin GUI
plt.switch_backend('Agg')

def process_excel_file(input_path, output_dir):
    # Generar nombres únicos para los archivos
    unique_id = uuid.uuid4().hex[:8]
    
    heatmap_path = os.path.join(output_dir, f"heatmap_{unique_id}.png")
    client_heatmap_path = os.path.join(output_dir, f"client_heatmap_{unique_id}.png")
    bar_chart_path = os.path.join(output_dir, f"barchart_{unique_id}.png")
    output_excel_path = os.path.join(output_dir, f"results_{unique_id}.xlsx")
    
    try:
        # ===============================
        # 2. Lectura y extracción de la información global
        # ===============================
        try:
            # Intentar leer 'Hoja1', si no existe, usar la primera hoja
            df_all = pd.read_excel(input_path, sheet_name='Hoja1', header=None)
        except:
            df_all = pd.read_excel(input_path, sheet_name=0, header=None)  # Usar la primera hoja
        
        # Verificar si el archivo tiene el formato mínimo esperado
        if df_all.shape[1] < 5:
            raise ValueError("El archivo no tiene suficientes columnas. Verifique el formato.")
        
        # Extraer información global
        info_global = {}
        for i in range(5):
            try:
                etiqueta = df_all.iloc[i, 3].strip()  # Columna D (índice 3)
                valor = df_all.iloc[i, 4].strip()     # Columna E (índice 4)
                info_global[etiqueta] = valor
            except Exception as e:
                raise ValueError(f"Error al leer la información global en la fila {i+1}: {str(e)}")

        # ===============================
        # 3. Extraer la tabla principal ("renglones")
        # ===============================
        main_headers = ["Renglón", "Opción", "Código", "Descripción", "Cantidad solicitada", "Unidad de medida"]

        # Buscar el índice de la fila "Total:" en la columna C (índice 2)
        end_index = None
        for idx in range(8, len(df_all)):
            if pd.isna(df_all.iloc[idx, 0]) and pd.isna(df_all.iloc[idx, 1]):
                cell_c = df_all.iloc[idx, 2]
                if isinstance(cell_c, str) and 'Total:' in cell_c:
                    end_index = idx
                    break
        if end_index is None:
            raise ValueError("No se encontró la fila con 'Total:' en la columna C.")

        # Extraer renglones desde la fila 9 (índice 8) hasta antes de "Total:" (columnas A-F)
        df_renglones = df_all.iloc[8:end_index, :6].copy()
        df_renglones.columns = main_headers

        # Eliminar la fila que repite los encabezados
        df_renglones = df_renglones[df_renglones["Renglón"] != "Renglón"].copy()
        df_renglones.reset_index(drop=True, inplace=True)

        # Convertir la columna "Renglón" a número y ordenar
        df_renglones["Renglón"] = df_renglones["Renglón"].apply(lambda x: int(str(x).strip()))
        df_renglones.sort_values("Renglón", inplace=True)
        df_renglones.reset_index(drop=True, inplace=True)

        # ===============================
        # 4. Extraer los nombres de las empresas
        # ===============================
        num_columnas_empresa = 6
        empresas = [
            df_all.iloc[7, col].strip()
            for col in range(6, df_all.shape[1], num_columnas_empresa)
            if pd.notna(df_all.iloc[7, col])
        ]

        # ===============================
        # 5. Extraer la parte de empresas (datos horizontales)
        # ===============================
        df_empresas = df_all.iloc[8:end_index, 6:].copy()

        # ===============================
        # 6. Separar los bloques de empresas
        # ===============================
        datos_empresas = {}
        for i, nombre in enumerate(empresas):
            inicio = i * num_columnas_empresa
            fin = inicio + num_columnas_empresa
            # Extraer el bloque y eliminar la primera fila (encabezado del bloque)
            df_empresa = df_empresas.iloc[:, inicio:fin].copy().iloc[1:, :].reset_index(drop=True)
            df_empresa.columns = [
                "Código Moneda", "Precio unitario", "Cantidad ofertada",
                "Total por renglón", "Especificacion técnica", "Total por renglón en ARS"
            ]
            datos_empresas[nombre] = df_empresa

        # ===============================
        # 7. Función para convertir valores a float (limpieza de símbolos)
        # ===============================
        def convertir_a_float(valor):
            if pd.isna(valor) or valor is None or str(valor).strip() == '':
                return 0.0
            # Eliminar '$', espacios, puntos (separador de miles) y reemplazar coma por punto
            valor = str(valor).replace('$', '').replace(' ', '').replace('.', '').replace(',', '.').strip()
            try:
                return float(valor)
            except ValueError:
                return 0.0

        # ===============================
        # 8. Calcular totales por empresa
        # ===============================
        totales_empresas = {
            nombre: datos_empresas[nombre]["Total por renglón en ARS"].apply(convertir_a_float).sum()
            for nombre in empresas
        }

        df_totales = pd.DataFrame(list(totales_empresas.items()), columns=["Empresa", "Total ARS"])
        df_totales["Total ARS"] = pd.to_numeric(df_totales["Total ARS"], errors="coerce").round(2)

        # ===============================
        # 9. Agrupar precios por "Renglón" (sin distinción de opción) y calcular ranking
        # ===============================
        prices_dict = {}
        for n in empresas:
            s = df_renglones.index.to_series().apply(
                lambda idx: convertir_a_float(datos_empresas[n].iloc[idx, 1]) if idx < len(datos_empresas[n]) else 0
            )
            s.index = df_renglones["Renglón"]
            prices_dict[n] = s.groupby(s.index).apply(lambda x: x[x > 0].min() if any(x > 0) else pd.NA)

        df_precio = pd.DataFrame({n: prices_dict[n] for n in empresas})
        df_precio.index = df_precio.index.astype(int)
        df_precio.sort_index(inplace=True)

        ranking_grouped = {}
        for r in df_precio.index:
            precios = df_precio.loc[r].to_dict()
            valid_prices = {n: p for n, p in precios.items() if pd.notna(p) and p > 0}
            ranking = {n: "NC" for n in precios.keys()}
            if valid_prices:
                sorted_prices = sorted(valid_prices.items(), key=lambda x: x[1])
                for pos, (n, p) in enumerate(sorted_prices, start=1):
                    ranking[n] = pos
            ranking_grouped[r] = ranking

        df_ranking_grouped = pd.DataFrame(ranking_grouped).T
        df_ranking_grouped.index = df_ranking_grouped.index.astype(int)
        df_ranking_grouped.sort_index(inplace=True)

        # ===============================
        # 10. Generar resumen por renglón (incluyendo diferencias con 2° y 3° mejores y diferencias %)
        # ===============================
        resumen_list = []
        cliente = "DIGITAL STRATEGY SAS"  # Se puede parametrizar

        for r in df_precio.index:
            prices = df_precio.loc[r].to_dict()
            valid_prices = {n: p for n, p in prices.items() if pd.notna(p) and p > 0}
            if valid_prices:
                sorted_prices = sorted(valid_prices.items(), key=lambda x: x[1])
                best_provider, best_price = sorted_prices[0]
                if len(sorted_prices) > 1:
                    second_provider, second_price = sorted_prices[1]
                else:
                    second_provider, second_price = None, pd.NA
                if len(sorted_prices) > 2:
                    third_provider, third_price = sorted_prices[2]
                else:
                    third_provider, third_price = None, pd.NA
            else:
                best_provider, best_price = None, pd.NA
                second_provider, second_price = None, pd.NA
                third_provider, third_price = None, pd.NA

            precio_cliente = prices.get(cliente, pd.NA)
            ranking_cliente = df_ranking_grouped.loc[r].get(cliente, "NC")

            def diffs(p_cliente, p_ref):
                if pd.notna(p_cliente) and pd.notna(p_ref) and p_ref != 0:
                    d = round(p_cliente - p_ref, 2)
                    p = round((p_cliente - p_ref) / p_ref * 100, 2)
                    return d, p
                return pd.NA, pd.NA

            diff_best, pct_diff_best = diffs(precio_cliente, best_price)
            diff_second, pct_diff_second = diffs(precio_cliente, second_price)
            diff_third, pct_diff_third = diffs(precio_cliente, third_price)

            resumen_list.append({
                "Renglón": r,
                "Mejor precio": round(best_price, 2) if pd.notna(best_price) else pd.NA,
                "Empresa mejor precio": best_provider,
                "Precio cliente": round(precio_cliente, 2) if pd.notna(precio_cliente) else pd.NA,
                "Ranking cliente": ranking_cliente,
                "Diferencia (cliente - mejor)": diff_best,
                "Diferencia (cliente - segundo)": diff_second,
                "Diferencia (cliente - tercer)": diff_third,
                "% Diferencia (cliente - mejor)": pct_diff_best,
                "% Diferencia (cliente - segundo)": pct_diff_second,
                "% Diferencia (cliente - tercer)": pct_diff_third
            })

        df_resumen_grouped = pd.DataFrame(resumen_list)
        df_resumen_grouped.sort_values("Renglón", inplace=True)
        df_resumen_grouped["Renglón"] = df_resumen_grouped["Renglón"].astype(int)

        # ===============================
        # 11. Heatmap global de rankings (estética personalizada)
        # ===============================
        plt.figure(figsize=(24, 16), facecolor='white')
        df_ranking_numeric = df_ranking_grouped.replace("NC", -1).astype(float)
        rank_data_for_heatmap = df_ranking_numeric.replace(-1, np.nan)
        ax = sns.heatmap(
            rank_data_for_heatmap,
            annot=df_ranking_grouped,
            fmt="",
            cmap="rocket_r",
            vmin=1,
            vmax=len(empresas),
            linewidths=0.5,
            linecolor='white',
            cbar_kws={'label': 'Ranking (1 = más barato; NC = No compite)'}
        )
        plt.title("Ranking de Precio Unitario por Renglón", fontsize=24, pad=20, color='0.3', fontweight='bold')
        plt.xlabel("Empresas Participantes", fontsize=16)
        plt.ylabel("Renglón", fontsize=16)
        plt.tight_layout()
        plt.savefig(heatmap_path, dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()

        # ===============================
        # 12. Heatmap del ranking del cliente (columna única)
        # ===============================
        plt.figure(figsize=(6, 16), facecolor='white')
        client_ranking = df_ranking_grouped[[cliente]].copy()
        client_ranking.index = client_ranking.index.astype(int)
        client_ranking.sort_index(inplace=True)
        client_ranking_numeric = client_ranking.replace("NC", -1).astype(float)
        ax = sns.heatmap(
            client_ranking_numeric.replace(-1, np.nan),
            annot=client_ranking,
            fmt="",
            cmap="rocket_r",
            vmin=1,
            vmax=len(empresas),
            linewidths=0.5,
            linecolor='white',
            cbar_kws={'label': f'Ranking {cliente} (1 = más barato; NC = No compite)'}
        )
        plt.title(f"Ranking de Precio Unitario de {cliente}", fontsize=20, pad=15, color='0.3', fontweight='bold')
        plt.xlabel("")
        plt.ylabel("Renglón", fontsize=14)
        plt.tight_layout()
        plt.savefig(client_heatmap_path, dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()

        # ===============================
        # 13. Gráfico de barras: Distribución del ranking del cliente
        # ===============================
        client_rank_series = pd.to_numeric(df_ranking_grouped[cliente].replace("NC", 0), errors="coerce")
        count_rank1 = (client_rank_series == 1).sum()
        count_rank2 = (client_rank_series == 2).sum()
        count_rank3 = (client_rank_series == 3).sum()
        count_rank4plus = (client_rank_series >= 4).sum()
        count_nc = (df_ranking_grouped[cliente] == "NC").sum()

        df_ranking_summary = pd.DataFrame({
            "Ranking": ["1", "2", "3", "4 o más", "NC"],
            "Cantidad de renglones": [count_rank1, count_rank2, count_rank3, count_rank4plus, count_nc]
        })

        plt.figure(figsize=(10, 6), facecolor='white')
        ax2 = sns.barplot(
            x="Ranking",
            y="Cantidad de renglones",
            data=df_ranking_summary,
            palette="rocket",
            edgecolor='black',
            linewidth=2
        )
        plt.title(f"Distribución de ranking de '{cliente}'", fontsize=22, pad=20, color='0.3', fontweight='bold')
        plt.xlabel("Ranking alcanzado", fontsize=16)
        plt.ylabel("Posicionamiento por renglón", fontsize=16)
        for i, row in df_ranking_summary.iterrows():
            ax2.text(i, row["Cantidad de renglones"] + 0.3, str(row["Cantidad de renglones"]),
                     color='black', ha="center", fontsize=12)
        plt.tight_layout()
        plt.savefig(bar_chart_path, dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()

        # ===============================
        # 14. Crear la hoja "Ofertas por renglon"
        # ===============================
        ofertas = []
        for r in df_precio.index:
            row_prices = df_precio.loc[r].dropna()
            # Filtrar solo las ofertas > 0
            valid_offers = row_prices[row_prices > 0]
            if not valid_offers.empty:
                valid_offers = valid_offers.sort_values()
                for empresa, monto in valid_offers.items():
                    ofertas.append({"Renglón": r, "Empresa": empresa, "Monto": round(monto, 2)})
        df_ofertas = pd.DataFrame(ofertas)
        df_ofertas.sort_values(by=["Renglón", "Monto"], inplace=True)

        # ===============================
        # 15. Exportar todas las hojas a un mismo Excel
        # ===============================
        with pd.ExcelWriter(output_excel_path, engine="xlsxwriter") as writer:
            df_resumen_grouped.to_excel(writer, sheet_name="Resumen", index=False)
            df_totales.to_excel(writer, sheet_name="Totales", index=False)
            df_ranking_summary.to_excel(writer, sheet_name="Ranking_cliente", index=False)
            df_ofertas.to_excel(writer, sheet_name="Ofertas por renglon", index=False)

        return [
            output_excel_path,
            heatmap_path,
            client_heatmap_path,
            bar_chart_path
        ]

    except Exception as e:
        # Limpiar archivos temporales en caso de error
        if os.path.exists(heatmap_path):
            os.remove(heatmap_path)
        if os.path.exists(client_heatmap_path):
            os.remove(client_heatmap_path)
        if os.path.exists(bar_chart_path):
            os.remove(bar_chart_path)
        if os.path.exists(output_excel_path):
            os.remove(output_excel_path)
        raise e