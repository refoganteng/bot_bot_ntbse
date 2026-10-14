import pyautogui
import time
import sys

print("=== DETEKSI KOORDINAT MOUSE ===")
print("Arahkan mouse ke tombol yang diinginkan, lalu catat angkanya.")
print("Tekan Ctrl+C untuk berhenti.\n")

try:
    while True:
        # Mengambil posisi x dan y kursor saat ini
        x, y = pyautogui.position()
        
        # Format teks agar rapi dan menimpa baris sebelumnya di terminal
        posisi_teks = f"X: {str(x).ljust(5)} | Y: {str(y).ljust(5)}"
        sys.stdout.write("\r" + posisi_teks)
        sys.stdout.flush()
        
        time.sleep(0.1) # Update setiap 0.1 detik biar gak berat
except KeyboardInterrupt:
    print("\n\nDeteksi dihentikan.")
