import webview
import base64

class Api:
    def save_excel(self, b64_data, default_filename):
        # 1. Ambil instance window yang sedang aktif
        window = webview.windows[0]
        
        # 2. Panggil jendela "Save As" bawaan sistem operasi
        file_path = window.create_file_dialog(
            dialog_type=webview.SAVE_DIALOG, 
            directory='', 
            save_filename=default_filename,
            file_types=('Excel Files (*.xlsx)', 'All files (*.*)')
        )
        
        # 3. Jika user memilih lokasi penyimpanan (tidak menekan Cancel)
        if file_path:
            try:
                # file_path berbentuk tuple, ambil elemen pertamanya
                with open(file_path[0], 'wb') as f:
                    f.write(base64.b64decode(b64_data))
                return "Berhasil disimpan"
            except Exception as e:
                print(f"Error saving file: {e}")
                return "Gagal menyimpan file"
        return "Dibatalkan"

if __name__ == '__main__':
    # Buat instance dari class Api penghubung
    api = Api()
    
    # Buat window dan sambungkan URL web dengan objek js_api
    webview.create_window(
        'Insurance Database and Comparison Maker', 
        'https://issamimar10.github.io/Database-and-Comparison-Insurance/',
        js_api=api,
        width=1200,
        height=800
    )
    
    webview.start()