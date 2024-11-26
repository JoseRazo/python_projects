import os
from yt_dlp import YoutubeDL

def descargar_lista_reproduccion(url, carpeta_destino):
    if not os.path.exists(carpeta_destino):
        os.makedirs(carpeta_destino)

    ydl_opts = {
        'format': 'bestaudio/best',
        'outtmpl': os.path.join(carpeta_destino, '%(title)s.%(ext)s'),
        'postprocessors': [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': '128',  # Reduce el bitrate a 128 kbps para archivos más pequeños
        }],
    }

    with YoutubeDL(ydl_opts) as ydl:
        ydl.download([url])

def leer_urls_y_carpeta_desde_archivo(archivo):
    # Leer las URLs desde el archivo y eliminar saltos de línea
    with open(archivo, 'r') as f:
        lines = f.read().splitlines()

    # La primera línea es la carpeta de destino, el resto son URLs
    carpeta_destino = lines[0]
    urls = lines[1:]
    
    return carpeta_destino, urls

if __name__ == "__main__":
    archivo_urls = "urls.txt"  # Archivo con la carpeta y las URLs
    
    # Leer la carpeta de destino y las URLs del archivo
    carpeta_destino, urls = leer_urls_y_carpeta_desde_archivo(archivo_urls)
    
    # Descargar cada lista de reproducción
    for url in urls:
        print(f"Descargando desde la URL: {url}")
        descargar_lista_reproduccion(url, carpeta_destino)
