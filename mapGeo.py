import folium
from folium.plugins import MarkerCluster, Fullscreen
import openpyxl
import requests  # Para carregar limites dos estados via GeoJSON

# Função para criar um marcador personalizado no mapa com ícone
def create_marker(coord, color='green', icon='info-sign'):
    geo_correta = coord.get('GeoCorreta')
    geo_correta_line = f"<b>GeoCorreta:</b> {geo_correta}<br>" if geo_correta else ""

    # Criar link para o Google Maps
    google_maps_link = f"https://www.google.com/maps/search/?api=1&query={coord['Latitude']},{coord['Longitude']}"

    # Adicionando mais informações e estilização ao popup
    popup_info = f"""
        <div style="font-family: Arial, sans-serif; line-height: 1.5; padding: 10px;">
            <h4 style="margin: 0; color: #2a9d8f;">{coord['Nome da Escola']}</h4>
            <p style="margin-bottom: 5px;"><i>{coord['Município']}</i>, {coord['UF']}</p>
            <p><b>Endereço:</b> {coord['Endereço']}</p>
            <p><b>Kit Wi-Fi (estimado):</b> {coord.get('Kit Wi-Fi (estimado)', 'N/A')}</p>
            <p><b>AP adicional (estimado):</b> {coord.get('AP adicional (estimado)', 'N/A')}</p>
            <p><b>Nobreak:</b> {coord.get('Nobreak', 'N/A')}</p>
            {geo_correta_line}
            <p><a href="{google_maps_link}" target="_blank">Ver no Google Maps</a></p>
        </div>
    """
    # Usando ícone customizado
    return folium.Marker(
        location=[coord["Latitude"], coord["Longitude"]],
        popup=folium.Popup(popup_info, max_width=300),
        icon=folium.Icon(color=color, icon=icon)
    )

# Carregar dados do arquivo Excel
wb = openpyxl.load_workbook("rampa.xlsx")
sheet = wb.active

coordenadas = []
estados = set()

for row in sheet.iter_rows(min_row=2):
    try:
        latitude = float(row[6].value)
        longitude = float(row[7].value)
    except (ValueError, TypeError):
        continue

    estado = row[1].value
    coordenadas.append({
        "LOTE": row[0].value,
        "UF": estado,
        "Município": row[2].value,
        "Código INEP": row[3].value,
        "Nome da Escola": row[4].value,
        "Endereço": row[5].value,
        "Latitude": latitude,
        "Longitude": longitude,
        "Kit Wi-Fi (estimado)": row[8].value,
        "AP adicional (estimado)": row[9].value,
        "Nobreak": row[10].value
    })
    estados.add(estado)

# Ordenar os estados em ordem alfabética
estados = sorted(estados)

# Criação do mapa com visualização inicial
m = folium.Map(location=[-15.788, -47.879], zoom_start=4)

# Adicionar clusters por estado
clusters = {estado: MarkerCluster(name=estado).add_to(m) for estado in estados}

# Dicionário para mapear valores de "Kit Wi-Fi (estimado)" a cores
wifi_color_map = {
    1: 'blue',
    2: 'green',
    3: 'purple',
    4: 'orange',
    5: 'darkred',
    6: 'lightred',
    7: 'beige',
    8: 'darkblue',
    9: 'darkgreen',
    10: 'cadetblue',
    11: 'darkpurple',
    12: 'white',
    13: 'pink',
    14: 'lightblue',
    15: 'lightgreen'
}

# Adicionar marcadores com ícones personalizados
for coord in coordenadas:
    # Obter a cor com base no valor de "Kit Wi-Fi (estimado)"
    kit_wifi = coord.get('Kit Wi-Fi (estimado)')
    color = wifi_color_map.get(kit_wifi, 'gray')  # 'gray' como cor padrão se não estiver no mapa

    # Criar marcador com a cor correspondente
    marker = create_marker(coord, color=color, icon='cloud')
    marker.add_to(clusters[coord['UF']])

# Adicionar os limites dos estados brasileiros via GeoJSON
geo_json_url = "https://raw.githubusercontent.com/codeforamerica/click_that_hood/master/public/data/brazil-states.geojson"
folium.GeoJson(geo_json_url, name="Limites dos Estados").add_to(m)

# Adicionar botão de fullscreen
Fullscreen().add_to(m)

# Adicionar controle de camadas com filtros
folium.LayerControl(collapsed=False).add_to(m)

# Mensagem de confirmação ao salvar o mapa
try:
    m.save("mapa_escolas_com_limites.html")
    print("Mapa salvo com sucesso!")
except Exception as e:
    print(f"Erro ao salvar o mapa: {e}")