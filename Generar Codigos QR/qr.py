import qrcode

# Texto o enlace que contendrá el QR
datos = "https://maps.app.goo.gl/FeHyfDxYSRAp8u3ZA"

# Crear QR
qr = qrcode.QRCode(
    version=1,
    error_correction=qrcode.constants.ERROR_CORRECT_H,
    box_size=10,
    border=4,
)

qr.add_data(datos)
qr.make(fit=True)

# Generar imagen
img = qr.make_image(fill_color="black", back_color="white")

# Guardar
img.save("codigo_qr.png")

print("QR generado correctamente.")